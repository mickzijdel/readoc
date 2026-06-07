"""Unit tests for bin/_readlib — the module shared by the readoc and readir CLIs.

These cover the pure helpers that used to live in (and be tested through) the
individual scripts before the shared-module extraction.
"""

import pytest


# --- human_size ---


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
def test_human_size(readlib_mod, size, expected):
    assert readlib_mod.human_size(size) == expected


# --- extension sets ---


def test_readable_sets(readlib_mod):
    assert ".md" in readlib_mod.READABLE_TEXT
    assert ".docx" in readlib_mod.READABLE_SPECIAL
    assert (
        readlib_mod.READABLE_ALL
        == readlib_mod.READABLE_TEXT | readlib_mod.READABLE_SPECIAL
    )


# --- read_text_file ---


def test_read_text_file_utf8(readlib_mod, tmp_path):
    f = tmp_path / "u.txt"
    f.write_text("héllo wörld\n", encoding="utf-8")
    assert readlib_mod.read_text_file(str(f)) == "héllo wörld\n"


def test_read_text_file_latin1_fallback(readlib_mod, tmp_path):
    f = tmp_path / "l.txt"
    # 0xFF is invalid UTF-8 but valid latin-1 (ÿ); reader should fall back.
    f.write_bytes(b"caf\xff\n")
    assert readlib_mod.read_text_file(str(f)) == "caf\xff\n"


# --- resolve_context (mode selection from the mutually exclusive flags) ---


class _Args:
    def __init__(self, context=None, context_paragraphs=None, context_chars=None):
        self.context = context
        self.context_paragraphs = context_paragraphs
        self.context_chars = context_chars


def test_resolve_context_defaults_to_lines(readlib_mod):
    assert readlib_mod.resolve_context(_Args()) == ("lines", 2)


def test_resolve_context_paragraphs(readlib_mod):
    assert readlib_mod.resolve_context(_Args(context_paragraphs=3)) == ("paragraphs", 3)


def test_resolve_context_chars(readlib_mod):
    assert readlib_mod.resolve_context(_Args(context_chars=80)) == ("chars", 80)


def test_resolve_context_explicit_lines(readlib_mod):
    assert readlib_mod.resolve_context(_Args(context=5)) == ("lines", 5)


# --- _mark (match highlighting) ---


def test_mark_wraps_case_insensitively_preserving_case(readlib_mod):
    assert readlib_mod._mark("The Budget is tight", "budget") == "The »Budget« is tight"


def test_mark_no_match_is_identity(readlib_mod):
    assert readlib_mod._mark("nothing here", "zzz") == "nothing here"
