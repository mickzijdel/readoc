"""Unit tests for bin/_xlsxzip.py — the byte-preserving xlsx surgery layer.

These drive the pure functions directly with hand-written worksheet XML, so they
can exercise shapes openpyxl never produces: Excel's ``standalone="yes"``
prologue, ``mc:Ignorable`` namespaces, ``t="s"`` shared-string cells, self-closing
cells and rows, row gaps, and the attribute lookalikes that a careless regex
would corrupt. The end-to-end behaviour is covered in test_editdoc.py.
"""

import zipfile
from xml.etree import ElementTree as ET

import pytest

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

# Deliberately Excel-shaped, not openpyxl-shaped: CRLF after the declaration and
# the markup-compatibility namespaces that ElementTree would rewrite on output.
SHEET_HEAD = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<worksheet xmlns="%s" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'mc:Ignorable="x14ac" '
    'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
) % NS_MAIN

# Row 1 mixes every cell type; row 2 has a gap at B and a self-closing cell;
# there is no row 3, so row-insertion has somewhere to land.
ROWS = (
    '<row r="1" spans="1:3" ht="76.5" customHeight="1" x14ac:dyDescent="0.25">'
    '<c r="A1" s="1" t="s"><v>0</v></c>'
    '<c r="B1" s="2"><v>1000</v></c>'
    '<c r="C1" t="inlineStr"><is><t>plain</t></is></c>'
    "</row>"
    '<row r="2" spans="1:3">'
    '<c r="A2" s="4"/>'
    '<c r="C2" t="b"><v>1</v></c>'
    "</row>"
    '<row r="4"><c r="A4" s="9" t="s"><v>0</v></c></row>'
)


def make_sheet(rows=ROWS, dimension='<dimension ref="A1:C4"/>', tail=""):
    return (
        SHEET_HEAD
        + dimension
        + "<sheetData>"
        + rows
        + "</sheetData>"
        + tail
        + "</worksheet>"
    ).encode("utf-8")


def cell_xml(sheet_bytes, ref):
    """Return the serialized <c r="ref"> element from a patched sheet."""
    text = sheet_bytes.decode("utf-8")
    start = text.index('<c r="%s"' % ref)
    if text[text.index(">", start) - 1] == "/":
        return text[start : text.index(">", start) + 1]
    return text[start : text.index("</c>", start) + 4]


def rows_in(sheet_bytes):
    """Row numbers in document order — proves insertion position."""
    root = ET.fromstring(sheet_bytes)
    data = root.find("{%s}sheetData" % NS_MAIN)
    return [int(row.get("r")) for row in data]


# --- parse_ref / column_index ---


@pytest.mark.parametrize(
    "ref, expected",
    [
        ("A1", ("A", 1)),
        ("$b$2", ("B", 2)),
        (" c3 ", ("C", 3)),
        ("XFD1048576", ("XFD", 1048576)),
    ],
)
def test_parse_ref_accepts_valid_references(xlsxzip_mod, ref, expected):
    assert xlsxzip_mod.parse_ref(ref) == expected


@pytest.mark.parametrize(
    "ref", ["1", "A", "", "A0", "A1048577", "XFE1", "A1:B2", 7, None]
)
def test_parse_ref_rejects_invalid_or_out_of_range(xlsxzip_mod, ref):
    with pytest.raises(xlsxzip_mod.XlsxError):
        xlsxzip_mod.parse_ref(ref)


@pytest.mark.parametrize(
    "letters, index", [("A", 1), ("Z", 26), ("AA", 27), ("XFD", 16384)]
)
def test_column_index_round_trips(xlsxzip_mod, letters, index):
    assert xlsxzip_mod.column_index(letters) == index
    assert xlsxzip_mod.column_letters(index) == letters


# --- render_cell: one case per value type ---


@pytest.mark.parametrize(
    "value, expected",
    [
        (None, '<c r="B2" s="3"/>'),
        (True, '<c r="B2" s="3" t="b"><v>1</v></c>'),
        (False, '<c r="B2" s="3" t="b"><v>0</v></c>'),
        (1250, '<c r="B2" s="3"><v>1250</v></c>'),
        (3.5, '<c r="B2" s="3"><v>3.5</v></c>'),
        ("hi", '<c r="B2" s="3" t="inlineStr"><is><t>hi</t></is></c>'),
    ],
)
def test_render_cell_by_value_type(xlsxzip_mod, value, expected):
    assert xlsxzip_mod.render_cell("B2", "3", value) == expected


def test_render_cell_without_style_omits_s_attribute(xlsxzip_mod):
    assert xlsxzip_mod.render_cell("B2", None, 1) == '<c r="B2"><v>1</v></c>'


def test_render_cell_escapes_xml_special_characters(xlsxzip_mod):
    rendered = xlsxzip_mod.render_cell("A1", None, 'a<b & c>d "q"')
    assert '<t>a&lt;b &amp; c&gt;d "q"</t>' in rendered


@pytest.mark.parametrize("value", [" leading", "trailing ", "two\nlines", "a\tb"])
def test_render_cell_preserves_significant_whitespace(xlsxzip_mod, value):
    assert 'xml:space="preserve"' in xlsxzip_mod.render_cell("A1", None, value)


def test_render_cell_omits_space_preserve_for_ordinary_text(xlsxzip_mod):
    assert "xml:space" not in xlsxzip_mod.render_cell("A1", None, "two words")


@pytest.mark.parametrize(
    "value", [["a"], {"a": 1}, float("inf"), float("nan"), "bell\x07", "_x0041_"]
)
def test_render_cell_rejects_unwritable_values(xlsxzip_mod, value):
    with pytest.raises(xlsxzip_mod.XlsxError):
        xlsxzip_mod.render_cell("A1", None, value)


# --- set_cell: replacing an existing cell ---


@pytest.mark.parametrize(
    "ref, style, description",
    [("A1", "1", "shared string"), ("B1", "2", "numeric"), ("A2", "4", "self-closing")],
)
def test_set_cell_replaces_any_cell_type_and_keeps_its_style(
    xlsxzip_mod, ref, style, description
):
    out = xlsxzip_mod.set_cell(make_sheet(), ref, "new")
    assert cell_xml(out, ref) == (
        '<c r="%s" s="%s" t="inlineStr"><is><t>new</t></is></c>' % (ref, style)
    ), description


def test_set_cell_numeric_replacement_drops_the_type_attribute(xlsxzip_mod):
    out = xlsxzip_mod.set_cell(make_sheet(), "A1", 42)
    assert cell_xml(out, "A1") == '<c r="A1" s="1"><v>42</v></c>'


def test_set_cell_null_clears_to_self_closing_keeping_style(xlsxzip_mod):
    out = xlsxzip_mod.set_cell(make_sheet(), "C1", None)
    assert cell_xml(out, "C1") == '<c r="C1"/>'


def test_set_cell_only_changes_the_target_cell(xlsxzip_mod):
    before = make_sheet()
    after = xlsxzip_mod.set_cell(before, "B1", "x")
    old, new = before.decode("utf-8"), after.decode("utf-8")
    assert old.replace('<c r="B1" s="2"><v>1000</v></c>', "") == new.replace(
        cell_xml(after, "B1"), ""
    )


def test_set_cell_does_not_disturb_a_cell_sharing_the_same_shared_string(xlsxzip_mod):
    # A1 and A4 both point at <si> index 0. Writing A1 inline must leave A4 alone.
    out = xlsxzip_mod.set_cell(make_sheet(), "A1", "changed")
    assert cell_xml(out, "A4") == '<c r="A4" s="9" t="s"><v>0</v></c>'


# --- set_cell: inserting cells and rows ---


def test_set_cell_inserts_cell_in_column_order(xlsxzip_mod):
    out = xlsxzip_mod.set_cell(make_sheet(), "B2", "mid").decode("utf-8")
    row = out[out.index('<row r="2"') : out.index("</row>", out.index('<row r="2"'))]
    assert row.index('r="A2"') < row.index('r="B2"') < row.index('r="C2"')


def test_set_cell_appends_cell_after_the_last_in_its_row(xlsxzip_mod):
    out = xlsxzip_mod.set_cell(make_sheet(), "D2", "end").decode("utf-8")
    row = out[out.index('<row r="2"') : out.index("</row>", out.index('<row r="2"'))]
    assert row.index('r="C2"') < row.index('r="D2"')


@pytest.mark.parametrize("ref, expected", [("A3", [1, 2, 3, 4]), ("A9", [1, 2, 4, 9])])
def test_set_cell_inserts_row_in_row_order(xlsxzip_mod, ref, expected):
    assert rows_in(xlsxzip_mod.set_cell(make_sheet(), ref, "v")) == expected


def test_set_cell_expands_a_self_closing_sheetdata(xlsxzip_mod):
    sheet = SHEET_HEAD.encode() + b'<dimension ref="A1"/><sheetData/></worksheet>'
    out = xlsxzip_mod.set_cell(sheet, "B3", "v")
    assert rows_in(out) == [3]
    assert cell_xml(out, "B3") == '<c r="B3" t="inlineStr"><is><t>v</t></is></c>'


def test_set_cell_expands_a_self_closing_row(xlsxzip_mod):
    out = xlsxzip_mod.set_cell(make_sheet(rows='<row r="4" spans="1:3"/>'), "B4", "v")
    assert cell_xml(out, "B4") == '<c r="B4" t="inlineStr"><is><t>v</t></is></c>'
    assert b'<row r="4" spans="1:3">' in out


# --- set_cell: the hints Excel keeps accurate ---


@pytest.mark.parametrize(
    "ref, expected",
    [("E9", 'ref="A1:E9"'), ("B2", 'ref="A1:C4"')],
)
def test_set_cell_widens_the_dimension_only_when_needed(xlsxzip_mod, ref, expected):
    assert expected.encode() in xlsxzip_mod.set_cell(make_sheet(), ref, "v")


def test_set_cell_leaves_an_absent_dimension_absent(xlsxzip_mod):
    out = xlsxzip_mod.set_cell(make_sheet(dimension=""), "E9", "v")
    assert b"<dimension" not in out


def test_set_cell_widens_row_spans_when_present(xlsxzip_mod):
    out = xlsxzip_mod.set_cell(make_sheet(), "F2", "v").decode("utf-8")
    assert '<row r="2" spans="1:6">' in out


def test_set_cell_leaves_absent_row_spans_absent(xlsxzip_mod):
    out = xlsxzip_mod.set_cell(make_sheet(), "F4", "v").decode("utf-8")
    assert '<row r="4">' in out


def test_set_cell_does_not_corrupt_attribute_lookalikes(xlsxzip_mod):
    # ht="76.5" also appears inside customHeight="1"; a naive substitution turns
    # that into customHeight="6" and Excel refuses the file.
    out = xlsxzip_mod.set_cell(make_sheet(), "A1", "v").decode("utf-8")
    assert 'ht="76.5" customHeight="1" x14ac:dyDescent="0.25"' in out


def test_set_cell_ref_match_is_exact_not_a_prefix(xlsxzip_mod):
    rows = '<row r="12"><c r="AB12" s="7"><v>1</v></c></row>'
    out = xlsxzip_mod.set_cell(make_sheet(rows=rows), "B12", "v")
    assert cell_xml(out, "AB12") == '<c r="AB12" s="7"><v>1</v></c>'
    assert cell_xml(out, "B12") == '<c r="B12" t="inlineStr"><is><t>v</t></is></c>'


def test_set_cell_ignores_lookalike_elements_outside_sheetdata(xlsxzip_mod):
    # <cols> and <rowBreaks> must not be mistaken for <c> and <row>.
    sheet = make_sheet(tail='<rowBreaks count="1"><brk id="1"/></rowBreaks>')
    sheet = sheet.replace(
        b"<sheetData>", b'<cols><col min="1" max="1" width="9"/></cols><sheetData>'
    )
    out = xlsxzip_mod.set_cell(sheet, "A1", "v")
    assert b'<cols><col min="1" max="1" width="9"/></cols>' in out
    assert b'<rowBreaks count="1">' in out


@pytest.mark.parametrize("ref", ["A1", "B2", "A3", "D9", "C1", "A4"])
def test_set_cell_always_produces_well_formed_xml(xlsxzip_mod, ref):
    ET.fromstring(xlsxzip_mod.set_cell(make_sheet(), ref, "value"))


# --- set_cell: refusals ---


def test_set_cell_refuses_a_formula_cell_and_names_the_formula(xlsxzip_mod):
    rows = '<row r="1"><c r="A1"><f>SUM(B1:C1)</f><v>7</v></c></row>'
    with pytest.raises(xlsxzip_mod.XlsxError) as excinfo:
        xlsxzip_mod.set_cell(make_sheet(rows=rows), "A1", 1)
    assert "formula" in str(excinfo.value)
    assert "SUM(B1:C1)" in str(excinfo.value)


@pytest.mark.parametrize(
    "rows",
    [
        '<row><c r="A1"><v>1</v></c></row>',  # row with no r=
        '<row r="1"><c><v>1</v></c></row>',  # cell with no r=
        '<row r="2"><c r="A2"/></row><row r="1"><c r="A1"/></row>',  # rows unordered
        '<row r="1"><c r="A5"/></row>',  # cell in the wrong row
    ],
)
def test_set_cell_refuses_unusual_worksheet_shapes(xlsxzip_mod, rows):
    with pytest.raises(xlsxzip_mod.XlsxError) as excinfo:
        xlsxzip_mod.set_cell(make_sheet(rows=rows), "A1", "v")
    assert "unusual worksheet XML" in str(excinfo.value)


def test_set_cell_refuses_a_sheet_without_sheetdata(xlsxzip_mod):
    with pytest.raises(xlsxzip_mod.XlsxError):
        xlsxzip_mod.set_cell(SHEET_HEAD.encode() + b"</worksheet>", "A1", "v")


def test_set_cell_refuses_a_merged_cell_that_is_not_the_anchor(xlsxzip_mod):
    sheet = make_sheet(
        tail='<mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>'
    )
    with pytest.raises(xlsxzip_mod.XlsxError) as excinfo:
        xlsxzip_mod.set_cell(sheet, "B1", "v")
    assert "merged" in str(excinfo.value) and "A1" in str(excinfo.value)


def test_set_cell_allows_the_anchor_of_a_merged_range(xlsxzip_mod):
    sheet = make_sheet(
        tail='<mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>'
    )
    out = xlsxzip_mod.set_cell(sheet, "A1", "v")
    assert cell_xml(out, "A1") == '<c r="A1" s="1" t="inlineStr"><is><t>v</t></is></c>'


# --- force_full_recalc ---

WORKBOOK_HEAD = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<workbook xmlns="%s">'
    % NS_MAIN
)


@pytest.mark.parametrize(
    "body",
    [
        '<sheets/><calcPr calcId="191028"/>',  # existing calcPr, no flag
        '<sheets/><calcPr calcId="191028" fullCalcOnLoad="0"/>',  # flag turned off
        "<sheets/>",  # no calcPr at all
        '<sheets/><extLst><ext uri="x"/></extLst>',  # must land before extLst
    ],
)
def test_force_full_recalc_sets_the_flag(xlsxzip_mod, body):
    out = xlsxzip_mod.force_full_recalc((WORKBOOK_HEAD + body + "</workbook>").encode())
    root = ET.fromstring(out)
    calc = root.find("{%s}calcPr" % NS_MAIN)
    assert calc is not None and calc.get("fullCalcOnLoad") == "1"
    assert [child.tag for child in root][-1].endswith("extLst") or "extLst" not in body


def test_force_full_recalc_is_idempotent(xlsxzip_mod):
    source = (
        WORKBOOK_HEAD + '<sheets/><calcPr fullCalcOnLoad="1"/></workbook>'
    ).encode()
    assert xlsxzip_mod.force_full_recalc(source) is source


def test_force_full_recalc_preserves_the_declaration(xlsxzip_mod):
    out = xlsxzip_mod.force_full_recalc(
        (WORKBOOK_HEAD + "<sheets/></workbook>").encode()
    )
    assert out.startswith(
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    )


# --- XlsxPackage ---


def build_package(path, sheet_targets, rel_type=None, entries=None):
    """Write a minimal workbook whose sheets map to the given rel targets.

    ``sheet_targets`` is a list of (sheet name, rel Target) in workbook order —
    deliberately decoupled from the sheetN.xml numbering so the tests can prove
    the mapping comes from the relationships, not from position.
    """
    rel_type = rel_type or NS_DOC_REL + "/worksheet"
    rels = "".join(
        '<Relationship Id="rId%d" Type="%s" Target="%s"/>' % (i + 1, rel_type, target)
        for i, (_, target) in enumerate(sheet_targets)
    )
    sheets = "".join(
        '<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (name, i + 1, i + 1)
        for i, (name, _) in enumerate(sheet_targets)
    )
    parts = {
        "[Content_Types].xml": '<?xml version="1.0"?><Types/>',
        "_rels/.rels": (
            '<Relationships xmlns="%s"><Relationship Id="rIdW" Type="%s" '
            'Target="xl/workbook.xml"/></Relationships>'
        )
        % (NS_PKG_REL, NS_DOC_REL + "/officeDocument"),
        "xl/workbook.xml": (
            '<workbook xmlns="%s" xmlns:r="%s"><sheets>%s</sheets></workbook>'
            % (NS_MAIN, NS_DOC_REL, sheets)
        ),
        "xl/_rels/workbook.xml.rels": (
            '<Relationships xmlns="%s">%s</Relationships>' % (NS_PKG_REL, rels)
        ),
    }
    for _, target in sheet_targets:
        # Targets are resolved against xl/workbook.xml, so a relative one gains
        # the xl/ prefix and an absolute one just loses its leading slash.
        resolved = target[1:] if target.startswith("/") else "xl/" + target
        parts[resolved] = make_sheet().decode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in parts.items():
            zout.writestr(name, data)
        for name, data, compress in entries or []:
            zout.writestr(zipfile.ZipInfo(name), data, compress_type=compress)
    return path


@pytest.mark.parametrize(
    "target", ["worksheets/sheet1.xml", "/xl/worksheets/sheet1.xml"]
)
def test_sheet_part_resolves_relative_and_absolute_targets(
    xlsxzip_mod, tmp_path, target
):
    path = build_package(tmp_path / "w.xlsx", [("Data", target)])
    assert (
        xlsxzip_mod.XlsxPackage(path).sheet_part("Data") == "xl/worksheets/sheet1.xml"
    )


def test_sheet_part_uses_relationships_not_sheet_order(xlsxzip_mod, tmp_path):
    path = build_package(
        tmp_path / "w.xlsx",
        [("Second", "worksheets/sheet2.xml"), ("First", "worksheets/sheet1.xml")],
    )
    pkg = xlsxzip_mod.XlsxPackage(path)
    assert pkg.sheet_names == ["Second", "First"]
    assert pkg.sheet_part("Second") == "xl/worksheets/sheet2.xml"
    assert pkg.sheet_part("First") == "xl/worksheets/sheet1.xml"


def test_sheet_part_rejects_an_unknown_sheet_and_lists_the_real_ones(
    xlsxzip_mod, tmp_path
):
    path = build_package(tmp_path / "w.xlsx", [("Data", "worksheets/sheet1.xml")])
    with pytest.raises(xlsxzip_mod.XlsxError) as excinfo:
        xlsxzip_mod.XlsxPackage(path).sheet_part("Nope")
    assert "sheet not found" in str(excinfo.value) and "Data" in str(excinfo.value)


def test_sheet_part_rejects_a_chartsheet(xlsxzip_mod, tmp_path):
    path = build_package(
        tmp_path / "w.xlsx",
        [("Chart1", "worksheets/sheet1.xml")],
        rel_type=NS_DOC_REL + "/chartsheet",
    )
    with pytest.raises(xlsxzip_mod.XlsxError) as excinfo:
        xlsxzip_mod.XlsxPackage(path).sheet_part("Chart1")
    assert "not a worksheet" in str(excinfo.value)


def test_package_rejects_a_file_that_is_not_a_zip(xlsxzip_mod, tmp_path):
    path = tmp_path / "broken.xlsx"
    path.write_bytes(b"this is not a zip")
    with pytest.raises(xlsxzip_mod.XlsxError):
        xlsxzip_mod.XlsxPackage(path)


def test_package_refuses_to_add_a_new_part(xlsxzip_mod, tmp_path):
    path = build_package(tmp_path / "w.xlsx", [("Data", "worksheets/sheet1.xml")])
    with pytest.raises(xlsxzip_mod.XlsxError):
        xlsxzip_mod.XlsxPackage(path).stage("xl/brand-new.xml", b"<x/>")


def test_package_read_returns_staged_bytes_so_edits_compose(xlsxzip_mod, tmp_path):
    path = build_package(tmp_path / "w.xlsx", [("Data", "worksheets/sheet1.xml")])
    pkg = xlsxzip_mod.XlsxPackage(path)
    part = pkg.sheet_part("Data")
    pkg.stage(part, xlsxzip_mod.set_cell(pkg.read(part), "A1", "one"))
    pkg.stage(part, xlsxzip_mod.set_cell(pkg.read(part), "B1", "two"))
    assert b"<t>one</t>" in pkg.read(part) and b"<t>two</t>" in pkg.read(part)


def test_package_save_preserves_entry_order_and_compression(xlsxzip_mod, tmp_path):
    path = build_package(
        tmp_path / "w.xlsx",
        [("Data", "worksheets/sheet1.xml")],
        entries=[("[trash]/0000.dat", b"scratch", zipfile.ZIP_STORED)],
    )
    before = zipfile.ZipFile(path)
    names, kinds = before.namelist(), [i.compress_type for i in before.infolist()]
    out = tmp_path / "out.xlsx"
    pkg = xlsxzip_mod.XlsxPackage(path)
    part = pkg.sheet_part("Data")
    pkg.stage(part, xlsxzip_mod.set_cell(pkg.read(part), "A1", "v"))
    pkg.save(out)
    after = zipfile.ZipFile(out)
    assert after.namelist() == names
    assert [i.compress_type for i in after.infolist()] == kinds
    assert after.read("[trash]/0000.dat") == b"scratch"
    assert [n for n in names if after.read(n) != before.read(n)] == [part]


def test_package_has_formulas_detects_a_formula_cell(xlsxzip_mod, tmp_path):
    plain = build_package(tmp_path / "plain.xlsx", [("Data", "worksheets/sheet1.xml")])
    assert xlsxzip_mod.XlsxPackage(plain).has_formulas() is False
    with_formula = tmp_path / "calc.xlsx"
    with zipfile.ZipFile(plain) as zin, zipfile.ZipFile(with_formula, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.endswith("sheet1.xml"):
                data = data.replace(b'<c r="B1" s="2">', b'<c r="B1" s="2"><f>A1</f>')
            zout.writestr(info, data)
    assert xlsxzip_mod.XlsxPackage(with_formula).has_formulas() is True
