"""Tests for the readir CLI (tree / read / search over a folder)."""

import os

import pytest


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
    big = next(s for rel, _f, s in readir_mod.walk_files(str(sample_tree))
               if rel.endswith("big.txt"))
    assert big > 2000


def test_walk_files_max_depth_prunes(readir_mod, sample_tree):
    # max_depth=1 descends one level (root + sub/) but prunes sub/deeper/.
    rels = [rel for rel, _f, _s in readir_mod.walk_files(str(sample_tree), max_depth=1)]
    assert "top.md" in rels
    assert any(rel.endswith("report.docx") for rel in rels)
    assert not any(rel.endswith("deepfile.txt") for rel in rels)


def test_walk_files_filter_and_exclude(readir_mod, sample_tree):
    only_md = [rel for rel, _f, _s in
               readir_mod.walk_files(str(sample_tree), filter_exts={".md"})]
    assert only_md == ["top.md"]

    no_png = [rel for rel, _f, _s in
              readir_mod.walk_files(str(sample_tree), exclude_exts={".png"})]
    assert not any(rel.endswith(".png") for rel in no_png)


# --- Unit: extension sets & text reading ---

def test_readable_sets(readir_mod):
    assert ".md" in readir_mod.READABLE_TEXT
    assert ".docx" in readir_mod.READABLE_SPECIAL
    assert readir_mod.READABLE_ALL == readir_mod.READABLE_TEXT | readir_mod.READABLE_SPECIAL


def test_read_text_file_utf8(readir_mod, tmp_path):
    f = tmp_path / "u.txt"
    f.write_text("héllo wörld\n", encoding="utf-8")
    assert readir_mod.read_text_file(str(f)) == "héllo wörld\n"


def test_read_text_file_latin1_fallback(readir_mod, tmp_path):
    f = tmp_path / "l.txt"
    # 0xFF is invalid UTF-8 but valid latin-1 (ÿ); reader should fall back.
    f.write_bytes(b"caf\xff\n")
    result = readir_mod.read_text_file(str(f))
    assert result == "caf\xff\n"


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
    out, _, code = readir("tree", sample_tree, "--max-depth", "1")
    assert code == 0
    assert "top.md" in out
    assert "report.docx" in out      # depth-1 file still shown
    assert "deepfile.txt" not in out  # depth-2 file pruned


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
