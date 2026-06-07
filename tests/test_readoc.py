"""Tests for the readoc CLI (single / multi-file reader)."""

import pytest


# --- Unit: pure helpers ---


@pytest.mark.parametrize(
    "size, expected",
    [
        (0, "0 B"),
        (1023, "1023 B"),
        (1024, "1.0 KB"),
        (1536, "1.5 KB"),
        (1024 * 1024, "1.0 MB"),
        (int(1.5 * 1024 * 1024), "1.5 MB"),
        (1024**3, "1.0 GB"),
    ],
)
def test_human_size(readoc_mod, size, expected):
    assert readoc_mod.human_size(size) == expected


def test_type_labels(readoc_mod):
    assert readoc_mod.TYPE_LABELS == {".docx": "Word", ".xlsx": "Excel", ".pdf": "PDF"}


def test_print_header_known_label(readoc_mod, capsys):
    readoc_mod.print_header("a/b.docx", ".docx", 2048)
    out = capsys.readouterr().out
    assert "a/b.docx (Word, 2.0 KB)" in out
    assert "═" in out


def test_print_header_unknown_ext_uppercased(readoc_mod, capsys):
    readoc_mod.print_header("weird.xyz", ".xyz", 10)
    out = capsys.readouterr().out
    assert "weird.xyz (XYZ, 10 B)" in out


# --- Integration: docx ---


def test_docx_headerbox_and_headings(readoc, docx_file):
    out, err, code = readoc(docx_file)
    assert code == 0, err
    assert "doc.docx (Word," in out
    assert "# Main Title" in out
    assert "## A Subsection" in out
    assert "Just a plain paragraph of body text." in out


def test_docx_table_rendered(readoc, docx_file):
    out, _, code = readoc(docx_file)
    assert code == 0
    assert "--- Table 1 ---" in out
    assert "Name | Value" in out
    assert "alpha | 42" in out


# --- Integration: xlsx ---


def test_xlsx_sheets_and_no_truncation(readoc, xlsx_file):
    from conftest import LONG_CELL

    out, err, code = readoc(xlsx_file)
    assert code == 0, err
    assert "=== Sheet: Data ===" in out
    assert "=== Sheet: Second ===" in out
    assert "=== Sheet: Blank ===" in out
    # The long cell must appear in full — no truncation despite the 40-char cap.
    assert LONG_CELL in out


def test_xlsx_empty_sheet_marker(readoc, xlsx_file):
    out, _, code = readoc(xlsx_file)
    assert code == 0
    assert "(empty sheet)" in out


# --- Integration: pdf ---


def test_pdf_single_page(readoc, pdf_file):
    out, err, code = readoc(pdf_file)
    assert code == 0, err
    assert "single.pdf (PDF," in out
    assert "Hello from page one." in out
    # Single page → no page separators.
    assert "--- Page" not in out


def test_pdf_multipage_separators(readoc, multipage_pdf):
    out, _, code = readoc(multipage_pdf)
    assert code == 0
    assert "--- Page 1 ---" in out
    assert "--- Page 2 ---" in out
    assert "Content of the first page." in out
    assert "Content of the second page." in out


# --- Integration: comments ---


def test_docx_comments_included_by_default(readoc, docx_with_comments):
    out, err, code = readoc(docx_with_comments)
    assert code == 0, err
    assert "--- Comments ---" in out
    assert "Alice: First reviewer note" in out
    assert "Bob: Second note" in out
    # Body text still present.
    assert "Just a plain paragraph of body text." in out


def test_docx_comments_suppressed_with_flag(readoc, docx_with_comments):
    out, err, code = readoc("--no-comments", docx_with_comments)
    assert code == 0, err
    assert "--- Comments ---" not in out
    assert "First reviewer note" not in out
    # Body text unaffected.
    assert "Just a plain paragraph of body text." in out


def test_xlsx_comments_included_by_default(readoc, xlsx_with_comments):
    out, err, code = readoc(xlsx_with_comments)
    assert code == 0, err
    assert "--- Comments ---" in out
    assert "[Data!A1]" in out
    assert "Carol: Header looks off" in out


def test_xlsx_comments_suppressed_with_flag(readoc, xlsx_with_comments):
    out, _, code = readoc("--no-comments", xlsx_with_comments)
    assert code == 0
    assert "--- Comments ---" not in out
    assert "Header looks off" not in out


def test_pdf_comments_included_by_default(readoc, pdf_with_comments):
    out, err, code = readoc(pdf_with_comments)
    assert code == 0, err
    assert "--- Comments ---" in out
    assert "[Page 1]" in out
    assert "Dave: A sticky note here" in out


def test_pdf_comments_suppressed_with_flag(readoc, pdf_with_comments):
    out, _, code = readoc("--no-comments", pdf_with_comments)
    assert code == 0
    assert "--- Comments ---" not in out
    assert "A sticky note here" not in out


def test_no_comment_block_when_no_comments(readoc, docx_file, xlsx_file, pdf_file):
    # Comment-free fixtures must not emit a spurious Comments header.
    out, _, code = readoc(docx_file, xlsx_file, pdf_file)
    assert code == 0
    assert "--- Comments ---" not in out


def test_no_comments_flag_not_treated_as_path(readoc, docx_file):
    # The flag must be stripped from the file list, not parsed as a filename.
    out, err, code = readoc("--no-comments", docx_file)
    assert code == 0, err
    assert "File not found" not in err
    assert "# Main Title" in out


# --- Integration: multiple files ---


def test_multiple_files_each_get_header(readoc, docx_file, pdf_file):
    out, _, code = readoc(docx_file, pdf_file)
    assert code == 0
    assert "doc.docx (Word," in out
    assert "single.pdf (PDF," in out


# --- Integration: error paths ---


def test_missing_file(readoc, tmp_path):
    out, err, code = readoc(tmp_path / "nope.docx")
    assert code == 1
    assert "File not found" in err


def test_unsupported_extension(readoc, tmp_path):
    f = tmp_path / "thing.png"
    f.write_bytes(b"\x89PNG")
    out, err, code = readoc(f)
    assert code == 1
    assert "Unsupported file type" in err


def test_no_args_shows_usage(readoc):
    out, err, code = readoc()
    assert code == 1
    assert "Usage: readoc" in out


def test_partial_success_still_prints_good_file(readoc, docx_file, tmp_path):
    out, err, code = readoc(docx_file, tmp_path / "missing.pdf")
    # Overall failure because one file was missing...
    assert code == 1
    assert "File not found" in err
    # ...but the readable file's content is still emitted.
    assert "# Main Title" in out
