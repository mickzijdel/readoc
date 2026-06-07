"""Tests for the readir CLI (tree / read / search over a folder)."""

import os

import pytest

from conftest import make_docx


@pytest.fixture
def commented_tree(tmp_path):
    """A folder with a single .docx carrying a comment whose text appears nowhere
    in the body — so any match/echo of it must come from comment extraction."""
    root = tmp_path / "ctree"
    root.mkdir()
    make_docx(
        root / "reviewed.docx",
        body="Ordinary body paragraph.",
        table=False,
        comments=[("xyzzy-comment-marker", "Reviewer")],
    )
    return root


# --- Unit: parse_extensions ---


@pytest.mark.parametrize("value", [None, ""])
def test_parse_extensions_empty_is_none(readir_mod, value):
    assert readir_mod.parse_extensions(value) is None


def test_parse_extensions_normalises(readir_mod):
    assert readir_mod.parse_extensions("md,docx, .PDF") == {".md", ".docx", ".pdf"}


def test_parse_extensions_strips_whitespace_and_dots(readir_mod):
    assert readir_mod.parse_extensions("  .TXT , csv ") == {".txt", ".csv"}


# --- Unit: walk_files ---


def test_walk_files_sorted_within_directory(readir_mod, sample_tree):
    rels = [rel for rel, _full, _size in readir_mod.walk_files(str(sample_tree))]
    # walk_files yields directory-by-directory, sorting filenames within each;
    # the root-level files therefore appear in sorted order among themselves.
    root_level = [r for r in rels if os.sep not in r]
    assert root_level == sorted(root_level)
    assert "top.md" in rels
    assert os.path.join("sub", "report.docx") in rels


def test_walk_files_reports_sizes(readir_mod, sample_tree):
    sizes = {rel: size for rel, _full, size in readir_mod.walk_files(str(sample_tree))}
    assert sizes["data.csv"] > 0
    # big.txt is the deliberately large file (~3.6 KB).
    big = next(
        s
        for rel, _f, s in readir_mod.walk_files(str(sample_tree))
        if rel.endswith("big.txt")
    )
    assert big > 2000


def test_walk_files_max_depth_1_is_top_level_only(readir_mod, sample_tree):
    # --max-depth 1 = immediate contents only, like `find -maxdepth 1`.
    rels = [rel for rel, _f, _s in readir_mod.walk_files(str(sample_tree), max_depth=1)]
    assert "top.md" in rels
    assert not any(os.sep in rel for rel in rels)  # nothing below the root


def test_walk_files_max_depth_2_descends_one_level(readir_mod, sample_tree):
    # --max-depth 2 = root + one level of subdirectories, pruning deeper.
    rels = [rel for rel, _f, _s in readir_mod.walk_files(str(sample_tree), max_depth=2)]
    assert any(rel.endswith("report.docx") for rel in rels)  # depth-1 file shown
    assert not any(rel.endswith("deepfile.txt") for rel in rels)  # depth-2 pruned


def test_walk_files_filter_and_exclude(readir_mod, sample_tree):
    only_md = [
        rel
        for rel, _f, _s in readir_mod.walk_files(str(sample_tree), filter_exts={".md"})
    ]
    assert only_md == ["top.md"]

    no_png = [
        rel
        for rel, _f, _s in readir_mod.walk_files(
            str(sample_tree), exclude_exts={".png"}
        )
    ]
    assert not any(rel.endswith(".png") for rel in no_png)


# (Extension sets and read_text_file now live in bin/_readlib — see test_readlib.py.)


# --- Integration: tree ---


def test_tree_lists_files(readir, sample_tree):
    out, err, code = readir("tree", sample_tree)
    assert code == 0, err
    assert "top.md" in out
    assert "report.docx" in out


def test_tree_summary(readir, sample_tree):
    out, _, code = readir("tree", sample_tree, "--summary")
    assert code == 0
    assert "files," in out and "total" in out
    assert ".md" in out


def test_tree_max_depth(readir, sample_tree):
    # --max-depth 1 shows only the top level.
    out, _, code = readir("tree", sample_tree, "--max-depth", "1")
    assert code == 0
    assert "top.md" in out
    assert "report.docx" not in out  # depth-1 file pruned at depth 1
    assert "deepfile.txt" not in out

    # --max-depth 2 descends one level into subdirectories.
    out2, _, code2 = readir("tree", sample_tree, "--max-depth", "2")
    assert code2 == 0
    assert "report.docx" in out2  # depth-1 file now shown
    assert "deepfile.txt" not in out2  # depth-2 still pruned


def test_tree_filter(readir, sample_tree):
    out, _, code = readir("tree", sample_tree, "--filter", "docx,pdf")
    assert code == 0
    assert "report.docx" in out
    assert "top.md" not in out


def test_tree_non_directory(readir, tmp_path):
    f = tmp_path / "afile.txt"
    f.write_text("x", encoding="utf-8")
    out, err, code = readir("tree", f)
    assert code == 1
    assert "Not a directory" in err


# --- Integration: read ---


def test_read_concatenates_readable(readir, sample_tree):
    out, err, code = readir("read", sample_tree)
    assert code == 0, err
    assert "top.md" in out
    assert "Some markdown about budgets." in out
    assert "Detailed budget discussion follows." in out  # from the docx
    # Unsupported .png is reported as skipped, not read.
    assert "unsupported format" in out


def test_read_filter_and_exclude(readir, sample_tree):
    out, _, code = readir("read", sample_tree, "--filter", "md")
    assert code == 0
    assert "Some markdown about budgets." in out
    assert "Detailed budget discussion follows." not in out

    out2, _, code2 = readir("read", sample_tree, "--exclude", "docx")
    assert code2 == 0
    assert "Detailed budget discussion follows." not in out2


def test_read_max_size_skips_and_lists(readir, sample_tree):
    out, _, code = readir("read", sample_tree, "--max-size", "1")
    assert code == 0
    # big.txt (~3.6 KB) exceeds the 1 KB ceiling → skipped and listed, not truncated.
    assert "big.txt" in out
    assert "too large" in out


def test_read_no_skip_report(readir, sample_tree):
    out, _, code = readir("read", sample_tree, "--max-size", "1", "--no-skip-report")
    assert code == 0
    assert "Skipped" not in out


def test_read_empty_dir(readir, tmp_path):
    empty = tmp_path / "empty"
    empty.mkdir()
    out, _, code = readir("read", empty)
    assert code == 0
    assert "(no files found)" in out


# --- Integration: read comments ---


def test_read_includes_comments_by_default(readir, commented_tree):
    out, err, code = readir("read", commented_tree)
    assert code == 0, err
    assert "--- Comments ---" in out
    assert "Reviewer: xyzzy-comment-marker" in out
    assert "Ordinary body paragraph." in out


def test_read_no_comments_flag_suppresses(readir, commented_tree):
    out, _, code = readir("read", commented_tree, "--no-comments")
    assert code == 0
    assert "--- Comments ---" not in out
    assert "xyzzy-comment-marker" not in out
    assert "Ordinary body paragraph." in out


# --- Integration: search ---


def test_search_case_insensitive_across_formats(readir, sample_tree):
    out, err, code = readir("search", sample_tree, "budget")
    assert code == 0, err
    assert "notes.txt" in out
    assert "top.md" in out
    assert "report.docx" in out  # "BUDGET" heading matched case-insensitively
    assert "total match" in out


def test_search_context(readir, sample_tree):
    out, _, code = readir("search", sample_tree, "budget", "--context", "5")
    assert code == 0
    assert ">>>" in out  # the match marker


def test_search_filter(readir, sample_tree):
    out, _, code = readir("search", sample_tree, "budget", "--filter", "txt")
    assert code == 0
    assert "notes.txt" in out
    assert "report.docx" not in out


def test_search_no_matches(readir, sample_tree):
    out, _, code = readir("search", sample_tree, "zzz-nonexistent-zzz")
    assert code == 0
    assert "No matches found" in out


def test_search_matches_comment_text_by_default(readir, commented_tree):
    out, err, code = readir("search", commented_tree, "xyzzy-comment-marker")
    assert code == 0, err
    assert "reviewed.docx" in out
    assert "total match" in out


def test_search_no_comments_excludes_comment_text(readir, commented_tree):
    out, _, code = readir(
        "search", commented_tree, "xyzzy-comment-marker", "--no-comments"
    )
    assert code == 0
    assert "No matches found" in out


def test_search_non_directory(readir, tmp_path):
    f = tmp_path / "afile.txt"
    f.write_text("x", encoding="utf-8")
    out, err, code = readir("search", f, "query")
    assert code == 1
    assert "Not a directory" in err


# --- Integration: help / no command ---


def test_help_text(readir):
    out, _, code = readir("help")
    assert code == 0
    assert "readir — Read entire folders of documents" in out


def test_no_command_shows_help(readir):
    out, _, code = readir()
    assert code == 0
    assert "Usage:" in out


# --- Integration: search context modes (paragraphs / chars) ---


def test_search_lines_is_default_mode(readir, search_tree):
    """With no context flag, prose search uses the numbered-line gutter."""
    out, err, code = readir("search", search_tree, "budget", "--filter", "docx")
    assert code == 0, err
    assert "report.docx (" in out
    assert ">>>" in out
    assert "budget" in out.lower()


def test_search_paragraphs_includes_neighbour_excludes_far(readir, search_tree):
    """--context-paragraphs 1 shows the hit ± one paragraph, marking units with ¶."""
    out, err, code = readir(
        "search",
        search_tree,
        "budget",
        "--context-paragraphs",
        "1",
        "--filter",
        "docx",
    )
    assert code == 0, err
    assert "¶" in out
    assert "»budget«" in out  # matched term highlighted
    assert "neighbour paragraph" in out  # immediate neighbour included
    assert "FARWORD" not in out  # two paragraphs away → excluded


def test_search_chars_window_and_marker(readir, search_tree):
    """--context-chars N shows a character window with the match wrapped in »…«."""
    out, err, code = readir(
        "search", search_tree, "budget", "--context-chars", "20", "--filter", "txt"
    )
    assert code == 0, err
    assert "blob.txt (" in out
    assert "»budget«" in out
    assert "…" in out  # truncated window delimiter


def test_search_context_modes_mutually_exclusive(readir, search_tree):
    out, err, code = readir(
        "search", search_tree, "budget", "--context", "2", "--context-chars", "50"
    )
    assert code == 2
    assert "not allowed with" in err


# --- Integration: search spreadsheets (cell-aware) ---


def test_search_xlsx_reports_coordinate_and_headers(readir, search_tree):
    out, err, code = readir("search", search_tree, "overrun", "--filter", "xlsx")
    assert code == 0, err
    assert "cell match" in out
    assert "Budget!C3" in out
    assert 'col "Notes"' in out
    assert 'row "Q2"' in out


def test_search_xlsx_shows_surrounding_rows(readir, search_tree):
    out, _, code = readir(
        "search", search_tree, "overrun", "--filter", "xlsx", "--context-rows", "1"
    )
    assert code == 0
    assert ">>>    3 |" in out  # matched row marked
    assert "Q1" in out  # adjacent row (2) shown within the ±1 window


def test_search_xlsx_searches_cell_comments(readir, search_tree):
    out, _, code = readir("search", search_tree, "double-check", "--filter", "xlsx")
    assert code == 0
    assert "--- Comments ---" in out
    # The matched term is highlighted, so the comment shows as »double-check« …
    assert "»double-check«" in out
    assert "Reviewer:" in out
    assert "this amount" in out


def test_search_xlsx_no_comments_flag_skips_comments(readir, search_tree):
    out, _, code = readir(
        "search", search_tree, "double-check", "--filter", "xlsx", "--no-comments"
    )
    assert code == 0
    assert "No matches found" in out
