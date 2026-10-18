"""Tests for the editdoc CLI (Edit-tool-equivalent for .docx / .xlsx).

editdoc reads a JSON edit spec (a single object or an array of objects) from
stdin and applies it to the target file given as an argv. Fixtures are built on
the fly with python-docx / openpyxl; results are verified by re-opening the file.
"""

import hashlib
import shutil
import subprocess
import zipfile

import docx
import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import Workbook, load_workbook

# Stand-ins for the SharePoint parts that openpyxl used to drop: the library's
# content-type schema, and the scratch blob SharePoint appends to every OOXML
# file uploaded through the Graph API (stored uncompressed, as the real one is).
CUSTOM_XML = b'<?xml version="1.0"?><properties xmlns="http://example.invalid/ct"/>'
TRASH_PART = b"\x00" * 64
VBA_PROJECT = b"\xd0\xcf\x11\xe0macro-project-stand-in"


# --- Fixture builders specific to editing tests ---


def add_hyperlink_run(paragraph, text, url="https://example.com"):
    """Append a real <w:hyperlink> (with a nested run) to ``paragraph``.

    python-docx's ``paragraph.runs`` does NOT include runs nested inside a
    hyperlink, so this builds the exact shape that a naive run-only rewrite would
    silently scramble — the substrate for the hyperlink-safety tests.
    """
    part = paragraph.part
    r_id = part.relate_to(
        url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    run = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    run.append(t)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def make_runs_docx(path, runs, style=None):
    """Write a .docx whose single body paragraph is composed of explicit runs.

    ``runs`` is a list of ``(text, bold)`` tuples, so tests can build a paragraph
    that is deliberately fragmented across runs with mixed formatting — the case
    that defeats a naive find/replace on the raw XML.
    """
    doc = Document()
    p = doc.add_paragraph()
    if style:
        p.style = style
    for text, bold in runs:
        r = p.add_run(text)
        r.bold = bold
    doc.save(str(path))
    return path


def make_paras_docx(path, paras):
    """Write a .docx with one plain paragraph per string in ``paras``."""
    doc = Document()
    for text in paras:
        doc.add_paragraph(text)
    doc.save(str(path))
    return path


def make_xlsx(path, rows=(("Quarter", "Amount"), ("Q1", 1000)), sheet="Budget"):
    """Write a small .xlsx with the given rows on a named sheet."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    for row in rows:
        ws.append(list(row))
    wb.save(str(path))
    return path


def make_xlsx_with_chart(path, sheet="Budget"):
    """Write an .xlsx that embeds a bar chart — a feature openpyxl drops on save,
    so editdoc must refuse to edit it without --force."""
    from openpyxl.chart import BarChart, Reference

    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(["Quarter", "Amount"])
    ws.append(["Q1", 1000])
    ws.append(["Q2", 1500])
    chart = BarChart()
    chart.add_data(
        Reference(ws, min_col=2, min_row=1, max_row=3), titles_from_data=True
    )
    ws.add_chart(chart, "E5")
    wb.save(str(path))
    return path


def zip_names(path):
    with zipfile.ZipFile(str(path)) as z:
        return z.namelist()


def add_zip_parts(path, parts):
    """Rewrite a package with extra zip entries appended verbatim.

    ``parts`` is a list of (name, bytes, compress_type). Used to bolt on the
    SharePoint-side parts that no Python writer produces, so the tests can prove
    an edit carries them through untouched.
    """
    with zipfile.ZipFile(str(path)) as zin:
        entries = [(info, zin.read(info.filename)) for info in zin.infolist()]
    with zipfile.ZipFile(str(path), "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in entries:
            zout.writestr(info, data)
        for name, data, compress in parts:
            zout.writestr(zipfile.ZipInfo(name), data, compress_type=compress)
    return path


def make_rich_xlsx(path, sheet="Budget"):
    """A workbook carrying everything the old openpyxl round-trip destroyed.

    A chart, an image, a cell comment (which is what puts a VML drawing under
    xl/drawings/ and used to trigger the blanket refusal), a formula, a merged
    range, plus SharePoint's customXml and [trash] parts.
    """
    from openpyxl.chart import BarChart, Reference
    from openpyxl.comments import Comment
    from openpyxl.drawing.image import Image

    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(["Quarter", "Amount"])
    ws.append(["Q1", 1000])
    ws.append(["Q2", 1500])
    ws["B4"] = "=SUM(B2:B3)"
    ws["A1"].comment = Comment("check this figure", "Reviewer")
    ws.merge_cells("D1:E1")
    chart = BarChart()
    chart.add_data(
        Reference(ws, min_col=2, min_row=1, max_row=3), titles_from_data=True
    )
    ws.add_chart(chart, "G5")
    png = path.parent / "dot.png"
    # A 1x1 PNG, written by hand so the suite needs no image library.
    png.write_bytes(
        b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
        b"\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\nIDATx\x9cc\x00\x01"
        b"\x00\x00\x05\x00\x01\r\n-\xb4\x00\x00\x00\x00IEND\xaeB`\x82"
    )
    ws.add_image(Image(str(png)), "G20")
    wb.save(str(path))
    return add_zip_parts(
        path,
        [
            ("customXml/item1.xml", CUSTOM_XML, zipfile.ZIP_DEFLATED),
            ("[trash]/0000.dat", TRASH_PART, zipfile.ZIP_STORED),
        ],
    )


# Hand-written because openpyxl never emits this shape: it writes inline strings
# and no sharedStrings part at all, so the t="s" path that real Excel and
# SharePoint files use would otherwise go completely untested.
EXCEL_SHAPED_PARTS = {
    "[Content_Types].xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        "</Types>"
    ),
    "_rels/.rels": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    ),
    "xl/workbook.xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
        '<calcPr calcId="191029"/></workbook>'
    ),
    "xl/_rels/workbook.xml.rels": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        # A relative target, as Excel writes them (openpyxl writes absolute ones).
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        "</Relationships>"
    ),
    # One <si>, referenced from two cells — the deduplication trap.
    "xl/sharedStrings.xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'count="2" uniqueCount="1"><si><t>shared</t></si></sst>'
    ),
    "xl/worksheets/sheet1.xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'mc:Ignorable="x14ac" '
        'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
        '<dimension ref="A1:C5"/><sheetData>'
        '<row r="1" spans="1:3" ht="30" customHeight="1" x14ac:dyDescent="0.25">'
        '<c r="A1" t="s"><v>0</v></c><c r="C1"/></row>'
        '<row r="2" spans="1:3"><c r="B2" t="s"><v>0</v></c></row>'
        '<row r="5" spans="1:3"><c r="A5"><v>99</v></c></row>'
        "</sheetData>"
        '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>'
        "</worksheet>"
    ),
}


def make_excel_shaped_xlsx(path, formula=False):
    """Write the Excel-shaped package; ``formula`` adds a cached formula cell.

    Its <calcPr> deliberately carries no fullCalcOnLoad, so the formula variant
    exercises editdoc adding the flag (openpyxl's own writer always sets it, so
    an openpyxl-built fixture cannot reach that path).
    """
    with zipfile.ZipFile(str(path), "w", zipfile.ZIP_DEFLATED) as zout:
        for name, text in EXCEL_SHAPED_PARTS.items():
            if formula and name == "xl/worksheets/sheet1.xml":
                text = text.replace(
                    '<row r="5"',
                    '<row r="3"><c r="A3"><f>A5*2</f><v>198</v></c></row><row r="5"',
                )
            zout.writestr(name, text)
    return path


def make_xlsm(path, sheet="Budget"):
    """A macro-enabled workbook: an .xlsx package carrying xl/vbaProject.bin.

    Zip surgery leaves the macro blob alone, which is what makes .xlsm editable
    at all — openpyxl's rewrite dropped it unless explicitly asked not to.
    """
    make_xlsx(path, sheet=sheet)
    return add_zip_parts(path, [("xl/vbaProject.bin", VBA_PROJECT, zipfile.ZIP_STORED)])


def cell_value(path, sheet, coord):
    wb = load_workbook(str(path))
    return wb[sheet][coord].value


def body_texts(path):
    """Return the visible text of every body paragraph in document order."""
    return [p.text for p in Document(str(path)).paragraphs]


def sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


# --- Task 1: skeleton (argv, stdin JSON, dispatch, error paths) ---


def test_no_args_shows_usage():
    # editdoc with zero argv should print usage and exit 1; reuse run_cli directly.
    from conftest import run_cli

    out, err, code = run_cli("editdoc", stdin="{}")
    assert code == 1
    assert "Usage" in out or "Usage" in err


@pytest.mark.parametrize("flag", ["--help", "-h"])
def test_help_prints_usage_and_edits_nothing(flag, tmp_path):
    # Agents run `<tool> --help` to learn it, so help must win even when a file
    # and a valid spec are also given.
    from conftest import run_cli

    f = make_paras_docx(tmp_path / "x.docx", ["hello"])
    before = f.read_bytes()
    out, err, code = run_cli(
        "editdoc", flag, f, stdin='{"old_string": "hello", "new_string": "hi"}'
    )
    assert code == 0, err
    assert "Usage" in out
    assert f.read_bytes() == before


def test_missing_file_errors(editdoc, tmp_path):
    out, err, code = editdoc(
        tmp_path / "nope.docx", {"old_string": "a", "new_string": "b"}
    )
    assert code == 1
    assert "not found" in err.lower()


def test_unsupported_extension_errors(editdoc, tmp_path):
    f = tmp_path / "x.pdf"
    f.write_text("hi")
    out, err, code = editdoc(f, {"old_string": "a", "new_string": "b"})
    assert code == 1
    assert "unsupported" in err.lower()


def test_invalid_json_errors(editdoc, tmp_path):
    f = make_paras_docx(tmp_path / "x.docx", ["hello world"])
    out, err, code = editdoc(f, "{bad json")
    assert code == 1
    assert "json" in err.lower()


def test_empty_spec_is_noop_success(editdoc, tmp_path):
    # An empty array of edits is a valid no-op: nothing to do, exit 0, file intact.
    f = make_paras_docx(tmp_path / "x.docx", ["hello world"])
    before = sha256(f)
    out, err, code = editdoc(f, [])
    assert code == 0, err
    assert sha256(f) == before


# --- Task 2: docx in-paragraph exact-match replace ---


def test_docx_basic_replace(editdoc, tmp_path):
    f = make_paras_docx(tmp_path / "x.docx", ["The report is due Monday."])
    out, err, code = editdoc(f, {"old_string": "Monday", "new_string": "Friday"})
    assert code == 0, err
    assert body_texts(f) == ["The report is due Friday."]
    assert "✓" in out  # success line


def test_docx_replace_is_whitespace_exact(editdoc, tmp_path):
    # A near-match with different internal spacing must NOT match.
    f = make_paras_docx(tmp_path / "x.docx", ["The  report  is due."])  # two spaces
    out, err, code = editdoc(
        f,
        {"old_string": "The report is", "new_string": "X"},  # single spaces
    )
    assert code == 1
    assert "not found" in err.lower()
    assert body_texts(f) == ["The  report  is due."]  # untouched


def test_docx_replace_across_fragmented_runs_inherits_first_run_format(
    editdoc, tmp_path
):
    # Visible text "The report is due" is split across three runs, the middle one
    # bold. Replacing a span that starts in the first (non-bold) run must succeed
    # and the replacement must inherit the FIRST spanned run's formatting.
    f = make_runs_docx(
        tmp_path / "x.docx",
        runs=[("The ", False), ("report", True), (" is due", False)],
    )
    out, err, code = editdoc(
        f, {"old_string": "The report is due", "new_string": "The memo is ready"}
    )
    assert code == 0, err
    doc = Document(str(f))
    para = doc.paragraphs[0]
    assert para.text == "The memo is ready"
    # The run carrying the replacement text inherits the first run's format (not bold).
    carrier = next(r for r in para.runs if "memo" in r.text)
    assert carrier.bold in (False, None)


def test_docx_replace_preserves_untouched_run_formatting(editdoc, tmp_path):
    # Replacing text wholly inside the first run leaves the bold middle run intact.
    f = make_runs_docx(
        tmp_path / "x.docx",
        runs=[("The ", False), ("report", True), (" is due", False)],
    )
    out, err, code = editdoc(f, {"old_string": "The ", "new_string": "A "})
    assert code == 0, err
    doc = Document(str(f))
    para = doc.paragraphs[0]
    assert para.text == "A report is due"
    bold_run = next(r for r in para.runs if r.text == "report")
    assert bold_run.bold is True


def test_docx_replace_not_found_errors(editdoc, tmp_path):
    f = make_paras_docx(tmp_path / "x.docx", ["hello world"])
    out, err, code = editdoc(f, {"old_string": "goodbye", "new_string": "hi"})
    assert code == 1
    assert "not found" in err.lower()
    assert body_texts(f) == ["hello world"]


def test_docx_replace_ambiguous_without_replace_all_errors(editdoc, tmp_path):
    f = make_paras_docx(
        tmp_path / "x.docx", ["alpha here", "alpha there"]
    )  # 'alpha' twice
    out, err, code = editdoc(f, {"old_string": "alpha", "new_string": "beta"})
    assert code == 1
    assert "2" in err  # mentions the count
    assert "replace_all" in err or "context" in err.lower()
    assert body_texts(f) == ["alpha here", "alpha there"]  # untouched


def test_docx_replace_all(editdoc, tmp_path):
    f = make_paras_docx(tmp_path / "x.docx", ["alpha here", "alpha there"])
    out, err, code = editdoc(
        f, {"old_string": "alpha", "new_string": "beta", "replace_all": True}
    )
    assert code == 0, err
    assert body_texts(f) == ["beta here", "beta there"]


def test_docx_replace_disambiguate_with_context(editdoc, tmp_path):
    # Adding surrounding context makes an otherwise-duplicated word unique.
    f = make_paras_docx(tmp_path / "x.docx", ["alpha here", "alpha there"])
    out, err, code = editdoc(f, {"old_string": "alpha here", "new_string": "beta here"})
    assert code == 0, err
    assert body_texts(f) == ["beta here", "alpha there"]


def test_docx_old_string_with_newline_errors(editdoc, tmp_path):
    f = make_paras_docx(tmp_path / "x.docx", ["line one", "line two"])
    out, err, code = editdoc(f, {"old_string": "line one\nline two", "new_string": "x"})
    assert code == 1
    assert "newline" in err.lower() or "paragraph" in err.lower()


def test_docx_replace_inside_table_cell(editdoc, tmp_path):
    # The edit surface includes table cells.
    doc = Document()
    t = doc.add_table(rows=1, cols=2)
    t.rows[0].cells[0].text = "Name"
    t.rows[0].cells[1].text = "old value"
    f = tmp_path / "x.docx"
    doc.save(str(f))
    out, err, code = editdoc(f, {"old_string": "old value", "new_string": "new value"})
    assert code == 0, err
    cell_text = Document(str(f)).tables[0].rows[0].cells[1].text
    assert cell_text == "new value"


# --- Task 3: docx paragraph / range replace ---


def test_docx_single_paragraph_replace(editdoc, tmp_path):
    f = make_paras_docx(
        tmp_path / "x.docx", ["Keep me", "Quarterly results were strong", "Keep me too"]
    )
    out, err, code = editdoc(
        f, {"start_contains": "Quarterly results", "new_text": "Replaced line"}
    )
    assert code == 0, err
    assert body_texts(f) == ["Keep me", "Replaced line", "Keep me too"]


def test_docx_paragraph_split_into_many(editdoc, tmp_path):
    # new_text with newlines turns one paragraph into several.
    f = make_paras_docx(tmp_path / "x.docx", ["before", "TARGET para", "after"])
    out, err, code = editdoc(
        f, {"start_contains": "TARGET", "new_text": "one\ntwo\nthree"}
    )
    assert code == 0, err
    assert body_texts(f) == ["before", "one", "two", "three", "after"]


def test_docx_range_replace_five_to_three(editdoc, tmp_path):
    f = make_paras_docx(
        tmp_path / "x.docx",
        ["head", "P1 start", "P2", "P3", "P4", "P5 end", "tail"],
    )
    out, err, code = editdoc(
        f,
        {
            "start_contains": "P1 start",
            "end_contains": "P5 end",
            "new_text": "new A\nnew B\nnew C",
        },
    )
    assert code == 0, err
    assert body_texts(f) == ["head", "new A", "new B", "new C", "tail"]


def test_docx_range_delete_with_empty_new_text(editdoc, tmp_path):
    f = make_paras_docx(
        tmp_path / "x.docx", ["keep", "drop 1", "drop 2", "drop 3", "keep too"]
    )
    out, err, code = editdoc(
        f,
        {"start_contains": "drop 1", "end_contains": "drop 3", "new_text": ""},
    )
    assert code == 0, err
    assert body_texts(f) == ["keep", "keep too"]


def test_docx_single_paragraph_delete(editdoc, tmp_path):
    f = make_paras_docx(tmp_path / "x.docx", ["keep", "remove me", "keep too"])
    out, err, code = editdoc(f, {"start_contains": "remove me", "new_text": ""})
    assert code == 0, err
    assert body_texts(f) == ["keep", "keep too"]


def test_docx_paragraph_split_inherits_style(editdoc, tmp_path):
    # Splitting a styled paragraph: every produced paragraph keeps the style.
    doc = Document()
    doc.add_paragraph("plain")
    doc.add_paragraph("Heading target", style="Heading 1")
    f = tmp_path / "x.docx"
    doc.save(str(f))
    out, err, code = editdoc(
        f, {"start_contains": "Heading target", "new_text": "H one\nH two"}
    )
    assert code == 0, err
    styled = [
        p.text for p in Document(str(f)).paragraphs if p.style.name == "Heading 1"
    ]
    assert styled == ["H one", "H two"]


def test_docx_paragraph_anchor_not_unique_errors(editdoc, tmp_path):
    f = make_paras_docx(tmp_path / "x.docx", ["dup line", "dup line"])
    out, err, code = editdoc(f, {"start_contains": "dup line", "new_text": "x"})
    assert code == 1
    assert "2" in err
    assert body_texts(f) == ["dup line", "dup line"]


def test_docx_paragraph_anchor_not_found_errors(editdoc, tmp_path):
    f = make_paras_docx(tmp_path / "x.docx", ["a", "b"])
    out, err, code = editdoc(f, {"start_contains": "nope", "new_text": "x"})
    assert code == 1
    assert "not found" in err.lower()


def test_docx_range_end_before_start_errors(editdoc, tmp_path):
    f = make_paras_docx(tmp_path / "x.docx", ["first", "second"])
    out, err, code = editdoc(
        f,
        {"start_contains": "second", "end_contains": "first", "new_text": "x"},
    )
    assert code == 1
    assert "before" in err.lower() or "order" in err.lower()


def test_docx_range_cross_container_errors(editdoc, tmp_path):
    doc = Document()
    doc.add_paragraph("start here in body")
    t = doc.add_table(rows=1, cols=1)
    t.rows[0].cells[0].text = "end here in cell"
    f = tmp_path / "x.docx"
    doc.save(str(f))
    out, err, code = editdoc(
        f,
        {
            "start_contains": "start here",
            "end_contains": "end here",
            "new_text": "x",
        },
    )
    assert code == 1
    assert "container" in err.lower() or "same" in err.lower()


# --- Task 4: xlsx cell edit ---


def test_xlsx_set_existing_cell(editdoc, tmp_path):
    f = make_xlsx(tmp_path / "x.xlsx", rows=(("Quarter", "Amount"), ("Q1", 1000)))
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "B2", "value": "1250"})
    assert code == 0, err
    assert cell_value(f, "Budget", "B2") == 1250  # numeric coercion
    assert "Budget!B2" in out


def test_xlsx_set_empty_cell(editdoc, tmp_path):
    f = make_xlsx(tmp_path / "x.xlsx", rows=(("Quarter", "Amount"), ("Q1", 1000)))
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "C5", "value": "hello"})
    assert code == 0, err
    assert cell_value(f, "Budget", "C5") == "hello"


def test_xlsx_numeric_vs_string_coercion(editdoc, tmp_path):
    f = make_xlsx(tmp_path / "x.xlsx")
    editdoc(f, {"sheet": "Budget", "cell": "A5", "value": "3.5"})
    editdoc(f, {"sheet": "Budget", "cell": "A6", "value": "N/A"})
    assert cell_value(f, "Budget", "A5") == 3.5
    assert cell_value(f, "Budget", "A6") == "N/A"


def test_xlsx_numeric_passed_as_json_number(editdoc, tmp_path):
    f = make_xlsx(tmp_path / "x.xlsx")
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "A7", "value": 42})
    assert code == 0, err
    assert cell_value(f, "Budget", "A7") == 42


def test_xlsx_bad_sheet_errors(editdoc, tmp_path):
    f = make_xlsx(tmp_path / "x.xlsx")
    out, err, code = editdoc(f, {"sheet": "Nonexistent", "cell": "A1", "value": "x"})
    assert code == 1
    assert "sheet" in err.lower()


def test_xlsx_bad_coordinate_errors(editdoc, tmp_path):
    f = make_xlsx(tmp_path / "x.xlsx")
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "1", "value": "x"})
    assert code == 1
    assert "cell" in err.lower() or "coordinate" in err.lower()


def test_xlsx_missing_value_errors(editdoc, tmp_path):
    f = make_xlsx(tmp_path / "x.xlsx")
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "A1"})
    assert code == 1


# --- Task 5: batch atomicity (validate-all-then-apply) ---


def test_docx_batch_all_or_nothing_on_failure(editdoc, tmp_path):
    # A valid edit followed by an invalid one must leave the file byte-identical:
    # the first edit's (in-memory) change is never written when a later edit fails.
    f = make_paras_docx(tmp_path / "x.docx", ["alpha unique", "beta line"])
    before = sha256(f)
    spec = [
        {"old_string": "alpha unique", "new_string": "ALPHA"},  # valid
        {"old_string": "does-not-exist", "new_string": "x"},  # invalid
    ]
    out, err, code = editdoc(f, spec)
    assert code == 1
    assert "edit 2" in err
    assert sha256(f) == before  # nothing written


def test_docx_batch_all_valid_applies_every_edit(editdoc, tmp_path):
    f = make_paras_docx(tmp_path / "x.docx", ["alpha", "beta", "gamma"])
    spec = [
        {"old_string": "alpha", "new_string": "A"},
        {"old_string": "gamma", "new_string": "G"},
    ]
    out, err, code = editdoc(f, spec)
    assert code == 0, err
    assert body_texts(f) == ["A", "beta", "G"]


def test_docx_batch_multiple_range_deletes(editdoc, tmp_path):
    # Two independent range deletes in one batch both succeed; deleting an earlier
    # block must not invalidate the second block's anchors.
    f = make_paras_docx(
        tmp_path / "x.docx",
        ["keep0", "d1a", "d1b", "keep1", "d2a", "d2b", "keep2"],
    )
    spec = [
        {"start_contains": "d1a", "end_contains": "d1b", "new_text": ""},
        {"start_contains": "d2a", "end_contains": "d2b", "new_text": ""},
    ]
    out, err, code = editdoc(f, spec)
    assert code == 0, err
    assert body_texts(f) == ["keep0", "keep1", "keep2"]


def test_xlsx_batch_all_or_nothing_on_failure(editdoc, tmp_path):
    f = make_xlsx(tmp_path / "x.xlsx")
    before = sha256(f)
    spec = [
        {"sheet": "Budget", "cell": "A5", "value": "1"},  # valid
        {"sheet": "Nonexistent", "cell": "A1", "value": "x"},  # invalid
    ]
    out, err, code = editdoc(f, spec)
    assert code == 1
    assert sha256(f) == before


# --- Review fixes: hyperlink safety, merged cells, empty anchors ---


def test_docx_edit_paragraph_with_hyperlink_refuses_loudly(editdoc, tmp_path):
    # A paragraph whose text includes a hyperlink (whose runs are invisible to
    # paragraph.runs) must NOT be rewritten — that silently scrambles it. Refuse
    # loudly and leave the file byte-identical.
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("See the ")
    add_hyperlink_run(p, "docs")
    p.add_run(" for details on Monday.")
    f = tmp_path / "x.docx"
    doc.save(str(f))
    before = sha256(f)
    out, err, code = editdoc(f, {"old_string": "Monday", "new_string": "Friday"})
    assert code == 1
    assert "hyperlink" in err.lower() or "safely" in err.lower()
    assert sha256(f) == before  # nothing scrambled


def test_docx_hyperlink_in_other_paragraph_does_not_block_edit(editdoc, tmp_path):
    # The guard is scoped to the edited paragraph: a hyperlink elsewhere in the
    # document must not block an edit to a clean paragraph.
    doc = Document()
    hp = doc.add_paragraph()
    hp.add_run("Visit ")
    add_hyperlink_run(hp, "our site")
    doc.add_paragraph("Plain paragraph with target word.")
    f = tmp_path / "x.docx"
    doc.save(str(f))
    out, err, code = editdoc(f, {"old_string": "target word", "new_string": "FIXED"})
    assert code == 0, err
    texts = [p.text for p in Document(str(f)).paragraphs]
    assert "Plain paragraph with FIXED." in texts
    # The hyperlinked paragraph is untouched and intact.
    assert any("Visit our site" == t for t in texts)


def test_docx_paragraph_op_on_hyperlink_paragraph_refuses(editdoc, tmp_path):
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("Anchor ")
    add_hyperlink_run(p, "linked")
    p.add_run(" tail")
    f = tmp_path / "x.docx"
    doc.save(str(f))
    before = sha256(f)
    out, err, code = editdoc(f, {"start_contains": "Anchor", "new_text": "replaced"})
    assert code == 1
    assert "hyperlink" in err.lower() or "safely" in err.lower()
    assert sha256(f) == before


def test_docx_replace_in_merged_cell(editdoc, tmp_path):
    # A horizontally-merged cell is one logical cell; a unique string in it must
    # not be reported as matching twice.
    doc = Document()
    t = doc.add_table(rows=1, cols=2)
    merged = t.rows[0].cells[0].merge(t.rows[0].cells[1])
    merged.text = "MERGED unique target"
    f = tmp_path / "x.docx"
    doc.save(str(f))
    out, err, code = editdoc(
        f, {"old_string": "MERGED unique target", "new_string": "done"}
    )
    assert code == 0, err
    assert Document(str(f)).tables[0].rows[0].cells[0].text == "done"


def test_docx_empty_start_contains_rejected(editdoc, tmp_path):
    f = make_paras_docx(tmp_path / "x.docx", ["only para"])
    before = sha256(f)
    out, err, code = editdoc(f, {"start_contains": "", "new_text": "HIJACKED"})
    assert code == 1
    assert body_texts(f) == ["only para"]
    assert sha256(f) == before


def test_docx_empty_end_contains_rejected(editdoc, tmp_path):
    f = make_paras_docx(tmp_path / "x.docx", ["alpha", "beta"])
    out, err, code = editdoc(
        f, {"start_contains": "alpha", "end_contains": "", "new_text": "x"}
    )
    assert code == 1


# --- Review round 2: xlsx coercion edges + lossy-feature detection ---


def test_xlsx_leading_zero_kept_as_string(editdoc, tmp_path):
    # Zip/account codes must not be truncated to ints.
    f = make_xlsx(tmp_path / "x.xlsx")
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "A5", "value": "00501"})
    assert code == 0, err
    assert cell_value(f, "Budget", "A5") == "00501"


def test_xlsx_nan_inf_kept_as_string_not_blanked(editdoc, tmp_path):
    # float("nan")/("inf") parse but openpyxl can't store them — they must NOT
    # silently blank the cell. Keep them as text.
    f = make_xlsx(tmp_path / "x.xlsx")
    for cell, val in (("A5", "nan"), ("A6", "inf"), ("A7", "-inf")):
        out, err, code = editdoc(f, {"sheet": "Budget", "cell": cell, "value": val})
        assert code == 0, err
    assert cell_value(f, "Budget", "A5") == "nan"
    assert cell_value(f, "Budget", "A6") == "inf"
    assert cell_value(f, "Budget", "A7") == "-inf"


def test_xlsx_scientific_and_underscore_kept_as_string(editdoc, tmp_path):
    f = make_xlsx(tmp_path / "x.xlsx")
    editdoc(f, {"sheet": "Budget", "cell": "A5", "value": "1e3"})
    editdoc(f, {"sheet": "Budget", "cell": "A6", "value": "1_000"})
    assert cell_value(f, "Budget", "A5") == "1e3"
    assert cell_value(f, "Budget", "A6") == "1_000"


def test_xlsx_plain_integer_still_coerced(editdoc, tmp_path):
    # The canonical-round-trip rule must still coerce ordinary numbers.
    f = make_xlsx(tmp_path / "x.xlsx")
    editdoc(f, {"sheet": "Budget", "cell": "A5", "value": "1250"})
    editdoc(f, {"sheet": "Budget", "cell": "A6", "value": "-42"})
    editdoc(f, {"sheet": "Budget", "cell": "A7", "value": "3.5"})
    assert cell_value(f, "Budget", "A5") == 1250
    assert cell_value(f, "Budget", "A6") == -42
    assert cell_value(f, "Budget", "A7") == 3.5


def test_xlsx_null_clears_cell(editdoc, tmp_path):
    f = make_xlsx(tmp_path / "x.xlsx", rows=(("Quarter", "Amount"), ("Q1", 1000)))
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "B2", "value": None})
    assert code == 0, err
    assert cell_value(f, "Budget", "B2") is None


def test_xlsx_bool_stored(editdoc, tmp_path):
    f = make_xlsx(tmp_path / "x.xlsx")
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "A5", "value": True})
    assert code == 0, err
    assert cell_value(f, "Budget", "A5") is True


def test_xlsx_plain_workbook_not_refused(editdoc, tmp_path):
    # A chart-free data sheet edits with no friction and no flag.
    f = make_xlsx(tmp_path / "x.xlsx")
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "A5", "value": "ok"})
    assert code == 0, err


# --- Package preservation: an edit must not cost the workbook any zip part ---
#
# The reason this whole path exists. openpyxl's save() regenerates the package
# and silently drops everything it does not model — on a real SharePoint file
# that meant 21 zip entries down to 11, taking the library's customXml metadata
# with it. These edit the richest workbook we can build and check nothing left.


def test_xlsx_comment_workbook_edits_without_any_flag(editdoc, tmp_path):
    # The headline regression: Excel anchors cell comments with a VML drawing
    # under xl/drawings/, which the old lossy-feature scan matched by prefix —
    # so *every* workbook with a comment was refused, and --force then deleted
    # the comments it was warning about.
    f = make_rich_xlsx(tmp_path / "x.xlsx")
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "A2", "value": "Q1x"})
    assert code == 0, err
    names = zip_names(f)
    assert any(n.startswith("xl/drawings/") and n.endswith(".vml") for n in names)
    assert any(n.startswith("xl/comments") for n in names)


def test_xlsx_edit_preserves_every_zip_entry_in_order(editdoc, tmp_path):
    f = make_rich_xlsx(tmp_path / "x.xlsx")
    before = zip_names(f)
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "A2", "value": "Q1x"})
    assert code == 0, err
    assert zip_names(f) == before  # list *and* order


def test_xlsx_edit_preserves_customxml_and_orphan_parts(editdoc, tmp_path):
    # The incident's literal symptom: SharePoint's content-type schema and its
    # own [trash] scratch blob, neither of which openpyxl knows how to keep.
    f = make_rich_xlsx(tmp_path / "x.xlsx")
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "A2", "value": "Q1x"})
    assert code == 0, err
    with zipfile.ZipFile(f) as z:
        assert z.read("customXml/item1.xml") == CUSTOM_XML
        assert z.read("[trash]/0000.dat") == TRASH_PART
        assert z.getinfo("[trash]/0000.dat").compress_type == zipfile.ZIP_STORED


def test_xlsx_edit_leaves_untouched_parts_byte_identical(editdoc, tmp_path):
    f = make_rich_xlsx(tmp_path / "x.xlsx")
    with zipfile.ZipFile(f) as z:
        before = {n: z.read(n) for n in z.namelist()}
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "A2", "value": "Q1x"})
    assert code == 0, err
    with zipfile.ZipFile(f) as z:
        changed = [n for n in z.namelist() if z.read(n) != before[n]]
    # Just the edited worksheet. (openpyxl's writer already sets the recalc flag
    # on this workbook, so even that part comes through untouched.)
    assert changed == ["xl/worksheets/sheet1.xml"]


def test_xlsx_edit_preserves_chart_and_image_parts(editdoc, tmp_path):
    f = make_rich_xlsx(tmp_path / "x.xlsx")
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "A2", "value": "Q1x"})
    assert code == 0, err
    names = zip_names(f)
    assert any(n.startswith("xl/charts/") for n in names)
    assert any(n.startswith("xl/media/") for n in names)


def test_xlsx_edit_size_delta_stays_small(editdoc, tmp_path):
    # A ballooning file is what jammed rclone's post-upload size check; patching
    # one part must not move the total appreciably.
    f = make_rich_xlsx(tmp_path / "x.xlsx")
    before = f.stat().st_size
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "A2", "value": "Q1x"})
    assert code == 0, err
    assert abs(f.stat().st_size - before) < before * 0.05


# --- Excel-shaped workbooks: the shared-strings path openpyxl never writes ---


def test_xlsx_edit_of_excel_shaped_workbook_replaces_shared_string_cell(
    editdoc, tmp_path
):
    f = make_excel_shaped_xlsx(tmp_path / "x.xlsx")
    out, err, code = editdoc(f, {"sheet": "Sheet1", "cell": "A1", "value": "changed"})
    assert code == 0, err
    assert cell_value(f, "Sheet1", "A1") == "changed"


def test_xlsx_edit_does_not_disturb_a_cell_sharing_the_shared_string(editdoc, tmp_path):
    # Shared strings are deduplicated: A1 and B2 both point at <si> 0. Editing
    # the <si> in place would silently change both, so editdoc writes inline.
    f = make_excel_shaped_xlsx(tmp_path / "x.xlsx")
    out, err, code = editdoc(f, {"sheet": "Sheet1", "cell": "A1", "value": "changed"})
    assert code == 0, err
    assert cell_value(f, "Sheet1", "B2") == "shared"


def test_xlsx_shared_strings_part_is_untouched_by_an_edit(editdoc, tmp_path):
    f = make_excel_shaped_xlsx(tmp_path / "x.xlsx")
    with zipfile.ZipFile(f) as z:
        before = z.read("xl/sharedStrings.xml")
    out, err, code = editdoc(f, {"sheet": "Sheet1", "cell": "A1", "value": "changed"})
    assert code == 0, err
    with zipfile.ZipFile(f) as z:
        assert z.read("xl/sharedStrings.xml") == before


def test_xlsx_edit_preserves_the_xml_declaration_and_namespaces(editdoc, tmp_path):
    # Re-serializing through ElementTree would drop standalone="yes" and
    # re-prefix mc:/x14ac:, which is why the sheet is patched as text.
    f = make_excel_shaped_xlsx(tmp_path / "x.xlsx")
    out, err, code = editdoc(f, {"sheet": "Sheet1", "cell": "A1", "value": "changed"})
    assert code == 0, err
    with zipfile.ZipFile(f) as z:
        sheet = z.read("xl/worksheets/sheet1.xml")
    assert sheet.startswith(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    assert b'mc:Ignorable="x14ac"' in sheet


# --- Formulas ---


def test_xlsx_formula_cell_refused_with_a_clean_error(editdoc, tmp_path):
    f = make_rich_xlsx(tmp_path / "x.xlsx")
    before = sha256(f)
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "B4", "value": 1})
    assert code == 1
    assert "Traceback" not in err
    assert "formula" in err.lower()
    assert sha256(f) == before


def test_xlsx_edit_sets_full_calc_on_load_when_the_workbook_has_formulas(
    editdoc, tmp_path
):
    # Formula cells cache their last result; without this flag the edit would
    # display a stale number until something forced a recalculation.
    f = make_rich_xlsx(tmp_path / "x.xlsx")
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "A2", "value": "Q1x"})
    assert code == 0, err
    with zipfile.ZipFile(f) as z:
        assert b'fullCalcOnLoad="1"' in z.read("xl/workbook.xml")


def test_xlsx_edit_adds_the_recalc_flag_when_the_workbook_lacks_it(editdoc, tmp_path):
    f = make_excel_shaped_xlsx(tmp_path / "x.xlsx", formula=True)
    with zipfile.ZipFile(f) as z:
        before = {n: z.read(n) for n in z.namelist()}
    assert b"fullCalcOnLoad" not in before["xl/workbook.xml"]
    out, err, code = editdoc(f, {"sheet": "Sheet1", "cell": "A5", "value": 5})
    assert code == 0, err
    with zipfile.ZipFile(f) as z:
        changed = sorted(n for n in z.namelist() if z.read(n) != before[n])
        assert b'fullCalcOnLoad="1"' in z.read("xl/workbook.xml")
    assert changed == ["xl/workbook.xml", "xl/worksheets/sheet1.xml"]


def test_xlsx_formula_free_workbook_touches_only_the_worksheet(editdoc, tmp_path):
    f = make_excel_shaped_xlsx(tmp_path / "x.xlsx")
    with zipfile.ZipFile(f) as z:
        before = {n: z.read(n) for n in z.namelist()}
    out, err, code = editdoc(f, {"sheet": "Sheet1", "cell": "A1", "value": "changed"})
    assert code == 0, err
    with zipfile.ZipFile(f) as z:
        changed = [n for n in z.namelist() if z.read(n) != before[n]]
    assert changed == ["xl/worksheets/sheet1.xml"]


# --- Clean errors where the old openpyxl path leaked a Python traceback ---


@pytest.mark.parametrize(
    "edit, expected",
    [
        ({"sheet": "Sheet1", "cell": "B1", "value": "x"}, "merged"),
        ({"sheet": "Sheet1", "cell": "A1048577", "value": "x"}, "range"),
        ({"sheet": "Sheet1", "cell": "XFE1", "value": "x"}, "range"),
        ({"sheet": "Sheet1", "cell": "A1", "value": ["a", "b"]}, "string"),
    ],
)
def test_xlsx_bad_edits_give_clean_errors_not_tracebacks(
    editdoc, tmp_path, edit, expected
):
    f = make_excel_shaped_xlsx(tmp_path / "x.xlsx")
    before = sha256(f)
    out, err, code = editdoc(f, edit)
    assert code == 1
    assert "Traceback" not in err
    assert expected in err.lower()
    assert sha256(f) == before


def test_xlsx_batch_stops_on_a_formula_cell_leaving_the_file_untouched(
    editdoc, tmp_path
):
    f = make_rich_xlsx(tmp_path / "x.xlsx")
    before = sha256(f)
    spec = [
        {"sheet": "Budget", "cell": "A2", "value": "Q1x"},  # valid
        {"sheet": "Budget", "cell": "B4", "value": 1},  # formula: refused
    ]
    out, err, code = editdoc(f, spec)
    assert code == 1
    assert "edit 2" in err
    assert sha256(f) == before


def test_xlsx_batch_applies_several_edits_to_one_sheet(editdoc, tmp_path):
    f = make_rich_xlsx(tmp_path / "x.xlsx")
    spec = [
        {"sheet": "Budget", "cell": "A2", "value": "Q1x"},
        {"sheet": "Budget", "cell": "D9", "value": 7},
    ]
    out, err, code = editdoc(f, spec)
    assert code == 0, err
    assert cell_value(f, "Budget", "A2") == "Q1x"
    assert cell_value(f, "Budget", "D9") == 7


def test_xlsx_deprecated_force_flag_is_accepted_with_a_notice(editdoc, tmp_path):
    f = make_rich_xlsx(tmp_path / "x.xlsx")
    out, err, code = editdoc(
        f, {"sheet": "Budget", "cell": "A2", "value": "Q1x"}, "--force"
    )
    assert code == 0, err
    assert "deprecated" in err.lower()
    assert cell_value(f, "Budget", "A2") == "Q1x"


# --- .xlsm: macro-enabled workbooks, editable now that nothing is rewritten ---


def test_xlsm_cell_edit_applies_and_keeps_the_macro_project(editdoc, tmp_path):
    f = make_xlsm(tmp_path / "x.xlsm")
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "B2", "value": 1250})
    assert code == 0, err
    assert cell_value(f, "Budget", "B2") == 1250
    with zipfile.ZipFile(f) as z:
        assert z.read("xl/vbaProject.bin") == VBA_PROJECT


def test_xlsm_is_not_rejected_as_an_unsupported_type(editdoc, tmp_path):
    f = make_xlsm(tmp_path / "x.xlsm")
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "A5", "value": "ok"})
    assert code == 0, err
    assert "Unsupported" not in err


def test_unsupported_extension_still_rejected(editdoc, tmp_path):
    f = tmp_path / "notes.rtf"
    f.write_text("hello")
    out, err, code = editdoc(f, {"old_string": "hello", "new_string": "hi"})
    assert code == 1
    assert "Unsupported file type" in err


def test_docx_edit_preserves_customxml_parts(editdoc, tmp_path):
    # python-docx walks the relationship graph rather than regenerating the
    # package, so customXml survives a .docx edit. (An orphan part outside the
    # graph does not — that limitation is documented, not fixed here.)
    f = make_paras_docx(tmp_path / "x.docx", ["hello"])
    before = zip_names(f)
    assert "customXml/item1.xml" in before  # the python-docx template ships one
    out, err, code = editdoc(f, {"old_string": "hello", "new_string": "goodbye"})
    assert code == 0, err
    assert set(zip_names(f)) >= set(before)


@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice not installed")
def test_xlsx_edited_workbook_opens_in_libreoffice(editdoc, tmp_path):
    # An independent reader's opinion that the patched package is still valid.
    f = make_rich_xlsx(tmp_path / "x.xlsx")
    out, err, code = editdoc(f, {"sheet": "Budget", "cell": "A2", "value": "Q1x"})
    assert code == 0, err
    result = subprocess.run(
        [
            "soffice",
            "--headless",
            "--convert-to",
            "ods",
            "--outdir",
            str(tmp_path),
            str(f),
        ],
        capture_output=True,
        timeout=180,
    )
    assert result.returncode == 0, result.stderr
    assert (tmp_path / "x.ods").exists()


def test_unknown_flag_rejected(tmp_path):
    from conftest import run_cli

    f = make_paras_docx(tmp_path / "x.docx", ["hello"])
    out, err, code = run_cli(
        "editdoc", f, "--frce", stdin='{"old_string": "hello", "new_string": "hi"}'
    )
    assert code == 1
    assert "flag" in err.lower() or "unknown" in err.lower()


def test_corrupt_docx_clean_error(editdoc, tmp_path):
    # A file with a .docx extension but invalid content must produce a clean
    # Error:, not a raw Python traceback.
    f = tmp_path / "broken.docx"
    f.write_bytes(b"this is not a real docx package")
    out, err, code = editdoc(f, {"old_string": "a", "new_string": "b"})
    assert code == 1
    assert "Traceback" not in err
    assert "valid" in err.lower() or "could not" in err.lower()


def test_corrupt_xlsx_clean_error(editdoc, tmp_path):
    f = tmp_path / "broken.xlsx"
    f.write_bytes(b"not a real xlsx zip")
    out, err, code = editdoc(f, {"sheet": "S", "cell": "A1", "value": "x"})
    assert code == 1
    assert "Traceback" not in err
    assert "valid" in err.lower() or "could not" in err.lower()


def test_multiple_files_rejected(tmp_path):
    # editdoc edits a single file; passing more than one must error, not silently
    # edit only the first.
    from conftest import run_cli

    a = make_paras_docx(tmp_path / "a.docx", ["one"])
    b = make_paras_docx(tmp_path / "b.docx", ["two"])
    before_b = sha256(b)
    out, err, code = run_cli(
        "editdoc",
        a,
        b,
        stdin='{"old_string": "one", "new_string": "X"}',
    )
    assert code == 1
    assert "one file" in err.lower() or "single" in err.lower()
    assert sha256(b) == before_b  # second file untouched
