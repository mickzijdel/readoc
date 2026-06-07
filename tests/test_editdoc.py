"""Tests for the editdoc CLI (Edit-tool-equivalent for .docx / .xlsx).

editdoc reads a JSON edit spec (a single object or an array of objects) from
stdin and applies it to the target file given as an argv. Fixtures are built on
the fly with python-docx / openpyxl; results are verified by re-opening the file.
"""

import hashlib

from docx import Document
from openpyxl import Workbook, load_workbook


# --- Fixture builders specific to editing tests ---


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
