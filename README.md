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

## Comments

Document comments are extracted by default and appended per file in a trailing
`--- Comments ---` block with the author and an anchor: Word comments
(`author: text`), Excel cell comments (`[Sheet!A1] author: text`), and PDF
annotations / sticky notes (`[Page N] author: text`). Pass `--no-comments` to
`readoc`, `readir read`, or `readir search` to omit them.

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

readoc carries a small `pyproject.toml` for **dev tooling only** — the CLIs
themselves stay self-contained PEP 723 scripts with no runtime dependencies. It
declares a dev dependency-group (`pytest`, `ruff`, and the doc libs) and is never
built or installed (`package = false`). Contributors get a reproducible
test/lint environment via uv:

```bash
uv sync                  # install the dev group into .venv
uv run pytest            # run the test suite
uv run ruff check .      # lint
uv run ruff format .     # format
```

The tests generate their own sample `.docx`/`.xlsx`/`.pdf`/text fixtures in a temp
directory (nothing binary is committed) and drive `bin/readoc` and `bin/readir`
end-to-end as subprocesses, alongside unit tests of the pure helpers.

### Git hooks (hk)

A [`mise`](https://mise.jdx.dev/) manifest pins the hook tooling (`hk`, `pkl`,
`uv`) and an `hk.pkl` runs ruff + pytest on commit. Set it up once:

```bash
mise install     # provision hk, pkl, uv
hk install       # install the pre-commit hook
```

The same checks run in CI ([.github/workflows/ci.yml](.github/workflows/ci.yml))
on every push and PR to `master`: `uv run ruff check .`, `uv run ruff format
--check .`, and `uv run pytest`.

## Layout

Source lives in `~/Stack/Programmeren/readoc` and is symlinked into
`~/.claude/skills/readoc` (auto-loads as `readoc@skills-dir`). The `bin/` scripts
are symlinked into `~/.local/bin`.

## License

MIT
