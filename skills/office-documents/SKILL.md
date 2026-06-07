---
name: office-documents
description: Use when you need to read OR edit the contents of Office documents (.docx, .xlsx, .pdf) or explore/search a folder of mixed documents. Provides the readoc, readir, and editdoc CLIs. Reach for this whenever a task involves reading or modifying a Word doc, Excel sheet, PDF, or a directory of such files.
---

# Reading & Editing Documents

Three CLIs ship with this plugin under `${CLAUDE_PLUGIN_ROOT}/bin/`. They are also
symlinked onto `PATH` as `readoc`, `readir`, and `editdoc`, so you can usually
call them by bare name.

The standard `Read`/`Edit`/`Write` tools cannot parse or modify
`.docx`/`.xlsx`/`.pdf`. Use these instead. **The readers never truncate
content** — spreadsheet cells, pages, and tables are emitted in full.

## `readoc` — read specific files

```bash
readoc file.docx                 # read a Word document
readoc file.xlsx                 # read an Excel spreadsheet
readoc file.pdf                  # read a PDF
readoc file1.docx file2.xlsx     # read several; each gets a header box
readoc --no-comments file.docx   # omit comments (included by default)

# Search inside one file (structure-aware context — see "Search modes" below)
readoc search file.docx "deadline"                       # default: line context
readoc search file.docx "deadline" --context-paragraphs 1
readoc search report.pdf "risk"    --context-chars 120
readoc search budget.xlsx "Q3"                           # cell-coordinate aware
```

Each file is preceded by an unambiguous header box (`path (Type, size)`), so
concatenated output is never confusing. Headings, tables, and sheets are
preserved; long cells overflow their column rather than being cut.

### Comments
Document comments are **included by default**, appended per file in a trailing
`--- Comments ---` block with the author and an anchor — Word comments
(`author: text`), Excel cell comments (`[Sheet!A1] author: text`), and PDF
annotations / sticky notes (`[Page N] author: text`). Pass `--no-comments` to
suppress them (body text is unaffected). The block is only emitted when a file
actually has comments.

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
readir read path/to/folder --no-comments      # omit comments (included by default)

# Search — grep across all docs (incl. docx/xlsx/pdf)
readir search path/to/folder "query"
readir search path/to/folder "query" --context 5            # lines of context (default 2)
readir search path/to/folder "query" --context-paragraphs 1 # whole paragraphs of context
readir search path/to/folder "query" --context-chars 120    # character window of context
readir search path/to/folder "query" --context-rows 3       # rows of context (spreadsheets)
readir search path/to/folder "query" --filter docx
readir search path/to/folder "query" --no-comments          # don't search comment text
```

### Search modes
Both `readir search` and `readoc search` are **structure-aware**, doing what a
plain `grep` cannot:

- **Prose** (`.docx`, `.pdf`, `.md`, `.txt`, …) — choose one context unit
  (mutually exclusive, default `--context` lines):
  - `--context N` — N lines around each match (the original behaviour).
  - `--context-paragraphs N` — the matching paragraph ± N whole paragraphs
    (units marked `¶`); respects real paragraph boundaries.
  - `--context-chars N` — a ±N character window around each match, merging
    overlapping windows. The matched term is wrapped in `»…«`.
- **Spreadsheets** (`.xlsx`) — always searched **cell by cell**. Each hit reports
  `Sheet!Cell` plus its column header (row 1) and row header (column A), e.g.
  `Budget!C7  [col "Amount" / row "Q3"]  <value>`, followed by a window of the
  matching rows ± `--context-rows` (default 2). Cell comments are searched too.

You can still pipe to ordinary tools for line-oriented work
(`readoc file.docx | grep -A3 deadline`, `… | head`, `find ./docs -name '*.docx'`)
— the built-in modes add the structure-aware context that piping can't.

Comments are extracted by default for `readir read` (and are searchable by both
search commands), in the same `--- Comments ---` form as `readoc`. Use
`--no-comments` to opt out.

Supported formats: `.md .txt .csv .docx .xlsx .pdf .json .yaml .yml`.

### The one deliberate limit
`readir read` defaults to skipping files larger than **500 KB** to avoid dumping
huge files unintentionally. Skipped files are **listed** in the skip report
(never silently dropped, never truncated). Raise the ceiling with
`--max-size <KB>` when you genuinely need a large file.

## `editdoc` — edit .docx / .xlsx files

`editdoc` is to binary Office documents what the built-in `Edit` tool is to text
files. It takes the target file as an argv and reads a JSON **edit spec from
stdin** — a single object, or an array of objects applied as one atomic batch.
The edit type is chosen by which keys the object carries:

```bash
# 1) Word — exact in-paragraph replace (mirrors the Edit tool)
echo '{"old_string": "due Monday", "new_string": "due Friday"}' | editdoc report.docx
echo '{"old_string": "TODO", "new_string": "Done", "replace_all": true}' | editdoc report.docx

# 2) Word — paragraph / range replace (structural)
echo '{"start_contains": "Quarterly results", "new_text": "Replaced paragraph."}' | editdoc report.docx
echo '{"start_contains": "First old para", "end_contains": "last old para", "new_text": "New A\nNew B"}' | editdoc report.docx
echo '{"start_contains": "Stale section", "new_text": ""}' | editdoc report.docx   # "" deletes

# 3) Excel — set a cell
echo '{"sheet": "Budget", "cell": "B2", "value": "1250"}' | editdoc book.xlsx

# Batch: an array is applied all-or-nothing
echo '[{"old_string":"a","new_string":"b"},{"sheet":"S","cell":"A1","value":1}]' | editdoc f.docx

# Force editing an xlsx that has charts/macros/etc (see Limitations)
echo '{"sheet":"Budget","cell":"B2","value":42}' | editdoc book.xlsx --force
```

### docx edit types

- **In-paragraph replace** — `{"old_string", "new_string", "replace_all"?}`.
  `old_string` is matched **exactly** (whitespace included) against each
  paragraph's visible text, across body paragraphs **and table cells**. It must
  be **unique** in the document; if it appears more than once you get a loud
  error — add surrounding context to disambiguate, or set `"replace_all": true`.
  A match cannot span a paragraph break (no newline in `old_string`).
- **Paragraph / range replace** — `{"start_contains", "end_contains"?, "new_text"}`.
  `start_contains` (and optional `end_contains`) each uniquely identify a
  paragraph; the inclusive block between them is replaced by `new_text`, split on
  `\n` into one-or-more paragraphs that inherit the original's style. Empty
  `new_text` deletes the block. A range must stay within one container (the body,
  or a single table cell).

### xlsx edit type and value coercion

- **Cell set** — `{"sheet", "cell", "value"}`. The sheet must exist and the cell
  must be a valid reference (`B2`). A **string** `value` becomes a number **only
  when it round-trips exactly**: `"1250"`→`1250`, `"-42"`→`-42`, `"3.5"`→`3.5`,
  but codes that would be corrupted stay text — leading zeros (`"00501"`),
  underscores (`"1_000"`), scientific notation (`"1e3"`), and non-finite values
  (`"nan"`, `"inf"`). Pass a JSON number (`42`, `3.5`) to force numeric storage; a
  JSON boolean stores a boolean; JSON `null` clears the cell.

### Why a naive find/replace fails on Word (and how editdoc handles it)

Word stores a visible sentence as a chain of formatting **runs** — `"the report
is due"` might be `["the ", "report", " is due"]` with the middle run bold. A
literal search on the raw XML usually misses it. `editdoc` reconstructs each
paragraph's text, matches there, and rewrites the runs so the replacement
**inherits the first matched run's formatting** while untouched text keeps its
own. After each edit it prints a confirmation line plus the affected paragraph
rendered as Markdown (`**bold**`, `*italic*`, heading `#`) so you can verify the
formatting landed as intended.

### Multi-paragraph rewrites

In-paragraph replace can't cross a paragraph break. To rewrite several
consecutive paragraphs at once, use a **single range replace**: anchor the first
paragraph with `start_contains`, the last with `end_contains`, and pass the full
new text in `new_text` (newline-separated for multiple paragraphs). This turns,
say, a five-paragraph block into three paragraphs in one atomic edit — no
per-paragraph delete chain.

### Safety

Every edit is validated before the file is written; if any edit in a batch fails,
**nothing** is written (the file stays byte-identical) and `editdoc` exits
non-zero with a specific error. Saves go through a temp file + atomic replace, so
a crash never corrupts the original. Read the result back with `readoc` to
confirm.

### Limitations

- **Hyperlinks / fields / footnote refs:** a paragraph whose text isn't carried
  entirely by ordinary runs (it contains a hyperlink, field, or footnote
  reference) is **refused** rather than rewritten — editing it would scramble the
  non-run content. Delete the whole paragraph (range op with empty `new_text`) or
  adjust the document instead. Such a paragraph can still be matched/deleted, just
  not rewritten in place.
- **Nested tables:** only top-level table cells are reached; text inside a table
  nested within a cell is not found.
- **xlsx fidelity:** openpyxl rewrites the whole workbook on save and can drop
  features it doesn't model. `editdoc` **scans the workbook and refuses** when it
  detects charts, pivot tables, drawings/images, or macros, naming what would be
  lost — pass `--force` (or `--allow-lossy`) to edit anyway and accept the loss.
  Plain data workbooks edit with no flag. (Detection is by file contents, so it
  can't see every lossy feature; `--force` skips the check entirely.)
- **No `.pdf` editing** (no clean text reflow).

## Requirements
[`uv`](https://docs.astral.sh/uv/) must be on `PATH`, it's the only prerequisite.
Each CLI declares its Python dependencies inline via PEP 723 and runs through
`uv run --script`, so uv installs them into a cached environment automatically on
first use: `readoc`/`readir` use `python-docx`, `openpyxl`, and `pymupdf` (for
`.docx`, `.xlsx`, and `.pdf`); `editdoc` uses `python-docx` and `openpyxl`.
