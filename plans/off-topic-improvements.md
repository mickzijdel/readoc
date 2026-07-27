# Off-topic improvements

Things noticed while working on something else. Not urgent, not in scope at the
time — pick them up when convenient.

## README structure (noticed 2026-07-27)

`README.md` fails the `writing:github-readme` audit on structure, and has since
before the xlsx rework: it has **4 H1 headings** where there should be one, and
no section literally titled *Installation* or *Usage* (the content is there,
under other names). Fixing it means restructuring the whole document, which is
too big to bundle into an unrelated change.

## Wider xlsx edit operations (noticed 2026-07-27)

`editdoc`'s xlsx support is still a single operation: set one cell. Now that the
zip-surgery layer (`bin/_xlsxzip.py`) exists, these become tractable and would
each be a small addition on top of `set_cell`:

- setting a **range** in one edit (`"cell": "A1:C1"` with a list of values)
- inserting or deleting a **row**, which also means shifting every `r=` below it
  and rewriting formula references — considerably harder, and probably where the
  "edit it in Excel" line belongs
- a **find/replace within a sheet**, mirroring the docx `old_string` edit

## Overwriting formula cells (noticed 2026-07-27)

Currently refused outright. If a real need appears, the way in is to remove the
`<f>` *and* the cell's `xl/calcChain.xml` entry (or drop the part along with its
`[Content_Types].xml` Override and workbook rel so Excel regenerates it).
Deferred because the refusal is the safe default and the demand is hypothetical.
