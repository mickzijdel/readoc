Make sure to check all of the following and make sure they are up-to-date after making changes;
1. tool-specific documentation for tools you edited
2. skills for tools you edited
3. plugin.json
4. README.md
5. CLAUDE.md
6. tests/ — keep the pytest suite green and add coverage for behaviour you change

## Tests

The `bin/readoc` and `bin/readir` CLIs are covered by a pytest suite in `tests/`.
Before changing either script, establish a green baseline:

```bash
uv pip install --python .venv -r tests/requirements-dev.txt   # first time
.venv/bin/python -m pytest tests/ -q
```

Tests generate their own fixtures in a temp dir and run the CLIs as subprocesses;
see `README.md` (Development / Testing) for details.
