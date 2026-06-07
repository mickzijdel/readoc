---
name: reading-documents
description: Use when you need to read the contents of Office documents (.docx, .xlsx, .pdf) or explore/search a folder of mixed documents. Provides the readoc and readir CLIs. Reach for this whenever a task involves a Word doc, Excel sheet, PDF, or a directory of such files.
---

# Reading Documents

Two CLIs ship with this plugin under `${CLAUDE_PLUGIN_ROOT}/bin/`. They are also
symlinked onto `PATH` as `readoc` and `readir`, so you can usually call them by
bare name.

The standard `Read` tool cannot parse `.docx`/`.xlsx`/`.pdf`. Use these instead.
**Neither tool truncates content** — spreadsheet cells, pages, and tables are
emitted in full.

## `readoc` — read specific files

```bash
readoc file.docx                 # read a Word document
readoc file.xlsx                 # read an Excel spreadsheet
readoc file.pdf                  # read a PDF
readoc file1.docx file2.xlsx     # read several; each gets a header box
```

Each file is preceded by an unambiguous header box (`path (Type, size)`), so
concatenated output is never confusing. Headings, tables, and sheets are
preserved; long cells overflow their column rather than being cut.

## `readir` — explore / read / search a folder

```bash
# Tree — explore structure
readir tree path/to/folder                    # file tree with sizes
readir tree path/to/folder --summary          # + extension/size summary
readir tree path/to/folder --max-depth 1      # depth: 1 = top level only, 2 = one level deep
readir tree path/to/folder --filter docx,pdf  # only certain extensions

# Read — concatenate all readable files (same header-box convention as readoc)
readir read path/to/folder
readir read path/to/folder --filter md,docx   # only specific types
readir read path/to/folder --exclude pdf      # skip certain types
readir read path/to/folder --max-size 2000    # KB ceiling; files over it are
                                              # SKIPPED (and listed), not truncated
readir read path/to/folder --max-depth 1
readir read path/to/folder --no-skip-report

# Search — grep across all docs (incl. docx/xlsx/pdf)
readir search path/to/folder "query"
readir search path/to/folder "query" --context 5   # lines of context (default 2)
readir search path/to/folder "query" --filter docx
```

Supported formats: `.md .txt .csv .docx .xlsx .pdf .json .yaml .yml`.

### The one deliberate limit
`readir read` defaults to skipping files larger than **500 KB** to avoid dumping
huge files unintentionally. Skipped files are **listed** in the skip report
(never silently dropped, never truncated). Raise the ceiling with
`--max-size <KB>` when you genuinely need a large file.

## Requirements
[`uv`](https://docs.astral.sh/uv/) must be on `PATH`, it's the only prerequisite.
Both CLIs declare their Python dependencies (`python-docx`, `openpyxl`, `pymupdf`,
for the `.docx`, `.xlsx`, and `.pdf` readers respectively) inline via PEP 723 and
run through `uv run --script`, so uv installs them into a cached environment
automatically on first use.
