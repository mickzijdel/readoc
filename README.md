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

`python-docx`, `openpyxl`, `pymupdf` (PyMuPDF).

## Layout

Source lives in `~/Stack/Programmeren/readoc` and is symlinked into
`~/.claude/skills/readoc` (auto-loads as `readoc@skills-dir`). The `bin/` scripts
are symlinked into `~/.local/bin`.

## License

MIT
