"""Byte-preserving cell surgery on .xlsx / .xlsm packages.

An OOXML workbook is a zip of XML parts. The obvious way to edit one — load it
with openpyxl and call ``save()`` — does not *patch* the package, it
**regenerates** it: every part openpyxl does not model is silently dropped.
On a real SharePoint-hosted workbook that took 21 zip entries down to 11,
losing ``customXml/*`` (the library's content-type schema and managed-metadata
columns) along with SharePoint's own ``[trash]/0000.dat``. Charts, pivot tables,
images, cell comments and their VML anchors, macros and in-cell rich text go the
same way.

So this module never round-trips the package. It patches the one worksheet part
the edit touches (plus ``xl/workbook.xml`` when a recalculation flag is needed)
and copies every other zip entry through byte-for-byte, preserving entry order
and per-entry compression. Anything it cannot rewrite with confidence, it
refuses — a clean error beats a quietly mangled workbook.

Two rules follow from that and are load-bearing throughout:

* **Read parts as bytes.** ``ZipFile.read()``, never text-mode ``open()`` on an
  extracted copy: universal-newline translation rewrites the ``\\r\\n`` that
  Excel puts in shared strings and after the XML declaration.
* **Never re-serialize XML.** ``ElementTree`` is used to *read* the workbook and
  relationship parts, and to *validate* a patched part, but its output is never
  written — it would rewrite the declaration, drop ``standalone="yes"``, and
  re-prefix the ``mc:``/``x14ac:``/``xr:`` namespaces real Excel files carry.
  Worksheet edits are anchored string surgery on the raw part text.
"""

import math
import posixpath
import re
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape as xml_escape

NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RT_OFFICE_DOCUMENT = NS_DOC_REL + "/officeDocument"
RT_WORKSHEET = NS_DOC_REL + "/worksheet"

MAX_ROW = 1048576
MAX_COLUMN = 16384

# Elements that follow calcPr in CT_Workbook's sequence. A new <calcPr> has to
# go before the first of these that is present, or the file fails validation.
CALCPR_SUCCESSORS = (
    "oleSize",
    "customWorkbookViews",
    "pivotCaches",
    "smartTagPr",
    "smartTagTypes",
    "webPublishing",
    "fileRecoveryPr",
    "webPublishObjects",
    "extLst",
)


class XlsxError(Exception):
    """A user-facing failure; editdoc re-raises it as an EditError."""


# --- tag and attribute matching ---

# A '>' inside a quoted attribute value is legal XML, so a naive [^>]* would cut
# the tag short. This consumes quoted runs atomically instead.
_ATTRS = r"""(?:"[^"]*"|'[^']*'|[^>"'])*"""


def _open_tag(name):
    """Match one <name …> or <name …/> tag, however its attributes are ordered.

    The ``(?=[\\s/>])`` lookahead is load-bearing: without it ``<c`` also matches
    ``<cols>`` and ``<calcPr>``, and ``<row`` matches ``<rowBreaks>``.
    """
    return re.compile("<" + name + r"(?=[\s/>])" + _ATTRS + ">")


def _attr(name):
    """Match one attribute inside an already-isolated tag string.

    Anchoring on a preceding boundary is what stops ``ht="76.5"`` also matching
    inside ``customHeight="1"``, and ``r="1"`` matching inside ``xr:uid``.
    """
    return re.compile(r"(?:^|\s)" + name + r'="([^"]*)"')


_SHEETDATA_TAG = _open_tag("sheetData")
_ROW_TAG = _open_tag("row")
_CELL_TAG = _open_tag("c")
_DIMENSION_TAG = _open_tag("dimension")
_CALCPR_TAG = _open_tag("calcPr")
_MERGECELL_TAG = _open_tag("mergeCell")

_R_ATTR = _attr("r")
_S_ATTR = _attr("s")
_SPANS_ATTR = _attr("spans")
_REF_ATTR = _attr("ref")
_FULLCALC_ATTR = _attr("fullCalcOnLoad")

_FORMULA_OPEN = re.compile(r"<f(?=[\s/>])")
_FORMULA_TEXT = re.compile(r"<f" + _ATTRS + r">(.*?)</f>", re.DOTALL)

_REF_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$")
_RANGE_RE = re.compile(r"^([A-Z]+[0-9]+)(?::([A-Z]+[0-9]+))?$")

# Characters XML 1.0 cannot represent at all, and Excel's own escape sequence:
# Excel decodes _xHHHH_ back to a character on read, so writing one literally is
# silent corruption. Both are refused rather than mangled.
_ILLEGAL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_X_ESCAPE = re.compile(r"_x[0-9A-Fa-f]{4}_")


# --- cell references ---


def column_index(letters):
    """'A' -> 1, 'AA' -> 27. ``letters`` must already be upper-case."""
    index = 0
    for char in letters:
        index = index * 26 + (ord(char) - 64)
    return index


def column_letters(index):
    """The inverse of column_index: 27 -> 'AA'."""
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def parse_ref(ref):
    """'$b$2' -> ('B', 2), rejecting anything Excel cannot address."""
    if not isinstance(ref, str):
        raise XlsxError(
            "invalid cell coordinate: %r (expected a reference like B2)." % (ref,)
        )
    match = _REF_RE.match(ref.strip())
    if match is None:
        raise XlsxError(
            "invalid cell coordinate: %r (expected a reference like B2)." % (ref,)
        )
    letters = match.group(1).upper()
    row = int(match.group(2))
    if row < 1 or row > MAX_ROW:
        raise XlsxError(
            "cell coordinate out of range: %r (rows go up to %d)." % (ref, MAX_ROW)
        )
    if column_index(letters) > MAX_COLUMN:
        raise XlsxError(
            "cell coordinate out of range: %r (the last column is %s)."
            % (ref, column_letters(MAX_COLUMN))
        )
    return letters, row


# --- rendering a single <c> element ---


def _render_number(value):
    if not math.isfinite(value):
        raise XlsxError("value must be a finite number (got %r)." % (value,))
    return str(value) if isinstance(value, int) else repr(value)


def _check_text(value):
    match = _ILLEGAL_CHARS.search(value)
    if match is not None:
        raise XlsxError(
            "value contains a control character (0x%02x) that XML cannot represent."
            % ord(match.group(0))
        )
    match = _X_ESCAPE.search(value)
    if match is not None:
        raise XlsxError(
            "value contains %r, which Excel reads back as an escape sequence rather "
            "than literal text; editdoc won't write it." % match.group(0)
        )


def _needs_space_preserve(value):
    return value != value.strip() or bool(set("\n\r\t") & set(value))


def render_cell(ref, style, value):
    """Render one <c> element. ``value`` is already coerced by editdoc."""
    attrs = ' r="%s"' % ref
    if style is not None:
        attrs += ' s="%s"' % style
    if value is None:
        # Style survives, type and content go — that is what "clear" means here.
        return "<c%s/>" % attrs
    # bool must be tested before int: isinstance(True, int) is True.
    if isinstance(value, bool):
        return '<c%s t="b"><v>%d</v></c>' % (attrs, 1 if value else 0)
    if isinstance(value, (int, float)):
        # No t= attribute: numeric is the OOXML default and what Excel writes.
        return "<c%s><v>%s</v></c>" % (attrs, _render_number(value))
    if isinstance(value, str):
        _check_text(value)
        space = ' xml:space="preserve"' if _needs_space_preserve(value) else ""
        return '<c%s t="inlineStr"><is><t%s>%s</t></is></c>' % (
            attrs,
            space,
            xml_escape(value),
        )
    raise XlsxError(
        "value must be a string, number, boolean, or null (got %s)."
        % type(value).__name__
    )


# --- worksheet surgery ---


def _unusual(detail):
    return XlsxError(
        "unusual worksheet XML that editdoc won't rewrite (%s); edit this file in "
        "Excel instead." % detail
    )


def _scan(text, start, end, pattern, close):
    """Collect sibling elements in text[start:end] as (start, tag_end, end, tag).

    ``<row>`` and ``<c>`` cannot nest, and a literal ``</row>`` or ``</c>``
    cannot appear in character data (a raw '<' is always escaped), so locating
    the close tag by plain search is exact.
    """
    found = []
    pos = start
    while True:
        match = pattern.search(text, pos, end)
        if match is None:
            return found
        tag = match.group(0)
        if tag.endswith("/>"):
            elem_end = match.end()
        else:
            close_at = text.find(close, match.end(), end)
            if close_at < 0:
                raise _unusual("unclosed %s element" % close)
            elem_end = close_at + len(close)
        found.append((match.start(), match.end(), elem_end, tag))
        pos = elem_end


def _required_r(tag, what):
    match = _R_ATTR.search(tag)
    if match is None:
        raise _unusual("a %s element with no r= position" % what)
    return match.group(1)


def _sheet_data_span(text):
    """Return (content_start, content_end, is_self_closing) for <sheetData>."""
    match = _SHEETDATA_TAG.search(text)
    if match is None:
        raise _unusual("no sheetData element")
    if match.group(0).endswith("/>"):
        return match.start(), match.end(), True
    close_at = text.find("</sheetData>", match.end())
    if close_at < 0:
        raise _unusual("unclosed sheetData element")
    return match.end(), close_at, False


def _widen_spans(tag, col):
    """Widen a <row>'s spans hint to include ``col``, if it carries one."""
    match = _SPANS_ATTR.search(tag)
    if match is None:
        return tag
    try:
        first, last = (int(part) for part in match.group(1).split(":"))
    except ValueError:
        return tag
    if first <= col <= last:
        return tag
    widened = "%d:%d" % (min(first, col), max(last, col))
    return tag[: match.start(1)] + widened + tag[match.end(1) :]


def _new_row(ref, col, row_num, value):
    # A bare <row r="N"> is valid and inherits the sheet defaults; copying a
    # neighbour's spans/ht/dyDescent would be a guess.
    return '<row r="%d">%s</row>' % (row_num, render_cell(ref, None, value))


def _formula_of(text, tag_end, elem_end):
    match = _FORMULA_TEXT.search(text, tag_end, elem_end)
    return match.group(1) if match is not None else ""


def _replace_existing_cell(text, cell, ref, value):
    cell_start, tag_end, cell_end, tag = cell
    if _FORMULA_OPEN.search(text, tag_end, cell_end) is not None:
        formula = _formula_of(text, tag_end, cell_end)
        raise XlsxError(
            "cell %s contains a formula (=%s); editdoc won't overwrite formula cells "
            "(it would strand the workbook's calculation chain). Edit the formula's "
            "inputs instead, or clear it in Excel first." % (ref, formula)
        )
    style_match = _S_ATTR.search(tag)
    style = style_match.group(1) if style_match is not None else None
    return text[:cell_start] + render_cell(ref, style, value) + text[cell_end:]


def _write_into_row(text, row, ref, col, row_num, value):
    row_start, tag_end, row_end, tag = row
    if tag.endswith("/>"):
        # <row r="4" spans="1:3"/> — expand it around the new cell.
        opened = _widen_spans(tag[:-2] + ">", col)
        return (
            text[:row_start]
            + opened
            + render_cell(ref, None, value)
            + "</row>"
            + text[row_end:]
        )

    inner_end = row_end - len("</row>")
    # Position every cell before touching anything: an out-of-order sibling
    # further along the row would make the insertion point below meaningless.
    cells = _positions(
        _scan(text, tag_end, inner_end, _CELL_TAG, "</c>"), row_num, "cell"
    )
    insert_at = inner_end
    for number, cell in cells:
        if number == col:
            text = _replace_existing_cell(text, cell, ref, value)
            return _replace_row_tag(text, row_start, tag_end, tag, col)
        if number > col:
            insert_at = min(insert_at, cell[0])
    text = text[:insert_at] + render_cell(ref, None, value) + text[insert_at:]
    return _replace_row_tag(text, row_start, tag_end, tag, col)


def _replace_row_tag(text, row_start, tag_end, tag, col):
    widened = _widen_spans(tag, col)
    if widened == tag:
        return text
    return text[:row_start] + widened + text[tag_end:]


def _positions(elements, row_num, what):
    """Number each scanned <row>/<c>, insisting the siblings ascend.

    Out-of-order siblings are legal-ish XML but break every assumption about
    where a new element belongs, so they are refused rather than guessed at.
    """
    numbered = []
    previous = 0
    for element in elements:
        raw = _required_r(element[3], what)
        if what == "row":
            try:
                number = int(raw)
            except ValueError:
                raise _unusual("a row numbered %r" % raw)
        else:
            letters, cell_row = parse_ref(raw)
            if cell_row != row_num:
                raise _unusual("cell %s sits in the row numbered %d" % (raw, row_num))
            number = column_index(letters)
        if number <= previous:
            raise _unusual("%ss are not in ascending order" % what)
        previous = number
        numbered.append((number, element))
    return numbered


def _write_cell(text, ref, col, row_num, value):
    start, end, self_closing = _sheet_data_span(text)
    if self_closing:
        expanded = "<sheetData>" + _new_row(ref, col, row_num, value) + "</sheetData>"
        return text[:start] + expanded + text[end:]

    rows = _positions(_scan(text, start, end, _ROW_TAG, "</row>"), row_num, "row")
    insert_at = end
    for number, row in rows:
        if number == row_num:
            return _write_into_row(text, row, ref, col, row_num, value)
        if number > row_num:
            insert_at = min(insert_at, row[0])
    return text[:insert_at] + _new_row(ref, col, row_num, value) + text[insert_at:]


def _check_merged(text, ref, col, row_num):
    for _, _, _, tag in _scan(text, 0, len(text), _MERGECELL_TAG, "</mergeCell>"):
        match = _REF_ATTR.search(tag)
        if match is None:
            continue
        span = _RANGE_RE.match(match.group(1))
        if span is None or span.group(2) is None:
            continue
        anchor = span.group(1)
        top_left, top_row = parse_ref(anchor)
        bottom_right, bottom_row = parse_ref(span.group(2))
        inside = top_row <= row_num <= bottom_row and column_index(
            top_left
        ) <= col <= column_index(bottom_right)
        if inside and ref != anchor:
            raise XlsxError(
                "cell %s is inside the merged range %s; only its anchor cell %s holds "
                "a value. Write to %s instead." % (ref, match.group(1), anchor, anchor)
            )


def _widen_dimension(text, col, row_num):
    match = _DIMENSION_TAG.search(text)
    if match is None:
        return text
    tag = match.group(0)
    ref_match = _REF_ATTR.search(tag)
    if ref_match is None:
        return text
    span = _RANGE_RE.match(ref_match.group(1))
    if span is None:
        return text
    first_letters, first_row = parse_ref(span.group(1))
    last_letters, last_row = parse_ref(span.group(2) or span.group(1))
    first_col, last_col = column_index(first_letters), column_index(last_letters)
    new_first = column_letters(min(first_col, col)) + str(min(first_row, row_num))
    new_last = column_letters(max(last_col, col)) + str(max(last_row, row_num))
    widened = new_first + ":" + new_last
    if widened == ref_match.group(1):
        return text
    tag = tag[: ref_match.start(1)] + widened + tag[ref_match.end(1) :]
    return text[: match.start()] + tag + text[match.end() :]


def _assert_well_formed(data, part):
    """Last line of defence: a surgery slip becomes a refusal, not a broken file."""
    try:
        ET.fromstring(data)
    except ET.ParseError as error:
        raise XlsxError(
            "internal error: refusing to write malformed XML into %s (%s). "
            "Please report this." % (part, error)
        )


def set_cell(sheet_xml, ref, value):
    """Return ``sheet_xml`` with one cell set, everything else byte-identical."""
    text = sheet_xml.decode("utf-8")
    letters, row_num = parse_ref(ref)
    col = column_index(letters)
    ref = letters + str(row_num)
    _check_merged(text, ref, col, row_num)
    text = _write_cell(text, ref, col, row_num, value)
    text = _widen_dimension(text, col, row_num)
    patched = text.encode("utf-8")
    _assert_well_formed(patched, "the worksheet")
    return patched


# --- workbook-level recalculation flag ---


def force_full_recalc(workbook_xml):
    """Set calcPr/@fullCalcOnLoad so Excel recomputes cached formula results.

    Every formula cell carries a cached ``<v>`` from the last calculation. A
    patched input leaves those stale, and Excel shows the stale number until
    something forces a recalculation — openpyxl's rewrite always set this flag,
    so a zip patch has to set it too or the edit silently displays wrong values.
    Idempotent.
    """
    text = workbook_xml.decode("utf-8")
    match = _CALCPR_TAG.search(text)
    if match is None:
        insert_at = _calcpr_insertion_point(text)
        addition = '<calcPr calcId="191029" fullCalcOnLoad="1"/>'
        text = text[:insert_at] + addition + text[insert_at:]
    else:
        tag = match.group(0)
        existing = _FULLCALC_ATTR.search(tag)
        if existing is not None:
            if existing.group(1) in ("1", "true"):
                return workbook_xml
            tag = tag[: existing.start(1)] + "1" + tag[existing.end(1) :]
        else:
            cut = -2 if tag.endswith("/>") else -1
            tag = tag[:cut] + ' fullCalcOnLoad="1"' + tag[cut:]
        text = text[: match.start()] + tag + text[match.end() :]
    patched = text.encode("utf-8")
    _assert_well_formed(patched, "the workbook")
    return patched


def _calcpr_insertion_point(text):
    for name in CALCPR_SUCCESSORS:
        match = _open_tag(name).search(text)
        if match is not None:
            return match.start()
    close_at = text.find("</workbook>")
    if close_at < 0:
        raise XlsxError("unusual workbook XML: no closing workbook element.")
    return close_at


# --- the package ---


def _rels_path(part):
    return posixpath.join(
        posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels"
    )


def _resolve(base_part, target):
    """Resolve a relationship target against the part that declares it."""
    if target.startswith("/"):
        return target[1:]
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


class XlsxPackage:
    """An open .xlsx/.xlsm zip, with staged part edits written out in one go."""

    def __init__(self, path):
        try:
            self._zin = zipfile.ZipFile(path)
            self._names = set(self._zin.namelist())
        except Exception as error:  # noqa: BLE001 - any open failure becomes XlsxError
            raise XlsxError(str(error))
        self._staged = {}
        self._workbook_part = self._find_workbook_part()
        self._sheets = self._read_sheet_map()

    # -- reading the package structure --

    def _parse(self, part):
        try:
            return ET.fromstring(self._zin.read(part))
        except KeyError:
            raise XlsxError("the workbook is missing its %s part." % part)
        except ET.ParseError as error:
            raise XlsxError("%s is not valid XML (%s)." % (part, error))

    def _relationships(self, part):
        """Map rId -> (resolved target, type) for one part's .rels."""
        rels_part = _rels_path(part)
        if rels_part not in self._names:
            return {}
        rels = {}
        for rel in self._parse(rels_part).findall("{%s}Relationship" % NS_PKG_REL):
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target")
            if not target:
                continue
            rels[rel.get("Id")] = (_resolve(part, target), rel.get("Type"))
        return rels

    def _find_workbook_part(self):
        # Follow the package relationships rather than hardcoding xl/workbook.xml:
        # the officeDocument part is what actually defines it.
        for target, rel_type in self._relationships("").values():
            if rel_type == RT_OFFICE_DOCUMENT:
                return target
        raise XlsxError("the package declares no workbook part.")

    def _read_sheet_map(self):
        rels = self._relationships(self._workbook_part)
        sheets = {}
        for sheet in self._parse(self._workbook_part).iter("{%s}sheet" % NS_MAIN):
            name = sheet.get("name")
            rel_id = sheet.get("{%s}id" % NS_DOC_REL)
            if name is not None:
                sheets[name] = rels.get(rel_id)
        return sheets

    # -- the public surface --

    @property
    def sheet_names(self):
        """Sheet names in workbook order."""
        return list(self._sheets)

    @property
    def workbook_part(self):
        return self._workbook_part

    def sheet_part(self, name):
        """The zip path of a worksheet, by sheet name."""
        if name not in self._sheets:
            raise XlsxError(
                "sheet not found: %r (available: %s)."
                % (name, ", ".join(self.sheet_names))
            )
        entry = self._sheets[name]
        if entry is None:
            raise XlsxError("the workbook's entry for sheet %r has no target." % name)
        target, rel_type = entry
        if rel_type != RT_WORKSHEET:
            raise XlsxError(
                "%r is not a worksheet (it is a chart or dialog sheet), so it has no "
                "cells to edit." % name
            )
        if target not in self._names:
            raise XlsxError(
                "the workbook references a worksheet part (%s) that is not in the "
                "file." % target
            )
        return target

    def read(self, part):
        """Part bytes — the staged version if this part has already been edited."""
        if part in self._staged:
            return self._staged[part]
        try:
            return self._zin.read(part)
        except KeyError:
            raise XlsxError("the workbook is missing its %s part." % part)

    def stage(self, part, data):
        """Queue patched bytes for a part that already exists in the archive."""
        if part not in self._names:
            raise XlsxError("cannot add a new part (%s) to the package." % part)
        self._staged[part] = data

    def has_formulas(self):
        if "xl/calcChain.xml" in self._names:
            return True
        return any(
            b"<f" in self.read(part)
            for part in self._names
            if part.startswith("xl/worksheets/") and part.endswith(".xml")
        )

    def save(self, dst):
        """Write the package to ``dst``, staged parts patched, the rest verbatim."""
        # Materialise everything first, then release the source handle: the
        # caller atomically replaces the original, which needs no open handles.
        entries = []
        for info in self._zin.infolist():
            staged = self._staged.get(info.filename)
            entries.append(
                (info, self._zin.read(info.filename) if staged is None else staged)
            )
        self.close()
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for info, data in entries:
                # Passing the original ZipInfo (not just a name) carries the
                # entry's compress_type, timestamp and extra field through —
                # SharePoint's [trash]/0000.dat is STORED while the rest of the
                # package is DEFLATEd, and entry order is preserved as-is.
                zout.writestr(info, data)

    def close(self):
        self._zin.close()
