# readoc

A Claude Code plugin for reading Office documents and folders of mixed documents
as plain text — without truncation.

## Tools

- **`readoc`** — read one or more `.docx` / `.xlsx` / `.pdf` files. Each file gets
  a header box (`path (Type, size)`); headings, tables, and sheets are preserved.
- **`readir`** — explore (`tree`), concatenate (`read`), or grep (`search`) an
  entire folder of mixed documents (`.md .txt .csv .docx .xlsx .pdf .json .yaml .yml`).

Both are exposed as the `reading-documents` skill (`readoc:reading-documents`) and
are symlinked onto `PATH` for direct CLI use.

## No truncation

Spreadsheet cells, PDF pages, and tables are emitted in full — long cells overflow
their column rather than being clipped. `readir read` will *skip* (and list) files
above `--max-size` (default 500 KB) instead of truncating them.

## Requirements

[`uv`](https://docs.astral.sh/uv/) on your `PATH`, that's the only prerequisite.
Both CLIs declare their Python dependencies (`python-docx`, `openpyxl`, `pymupdf`)
inline via [PEP 723](https://peps.python.org/pep-0723/) and run through
`uv run --script`, so uv installs them into a cached environment automatically on
first use. No manual `pip install` and no venv to manage.

## Development / Testing

A `pytest` suite covers both CLIs (`tests/`). Install the dev dependencies and run it:

```bash
uv venv .venv                                      # or: python -m venv .venv
uv pip install --python .venv -r tests/requirements-dev.txt
.venv/bin/python -m pytest tests/ -q               # or just: pytest -q
```

The tests generate their own sample `.docx`/`.xlsx`/`.pdf`/text fixtures in a temp
directory (nothing binary is committed) and drive `bin/readoc` and `bin/readir`
end-to-end as subprocesses, alongside unit tests of the pure helpers. The dev
requirements mirror the CLIs' runtime libs (`python-docx`, `openpyxl`, `pymupdf`)
plus `pytest`.

## Layout

Source lives in `~/Stack/Programmeren/readoc` and is symlinked into
`~/.claude/skills/readoc` (auto-loads as `readoc@skills-dir`). The `bin/` scripts
are symlinked into `~/.local/bin`.

## License

MIT
