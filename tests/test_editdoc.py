"""Tests for the editdoc CLI (Edit-tool-equivalent for .docx / .xlsx).

editdoc reads a JSON edit spec (a single object or an array of objects) from
stdin and applies it to the target file given as an argv. Fixtures are built on
the fly with python-docx / openpyxl; results are verified by re-opening the file.
"""

import hashlib

from docx import Document


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
