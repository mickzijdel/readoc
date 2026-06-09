Make sure to check all of the following and make sure they are up-to-date after making changes;
1. tool-specific documentation for tools you edited
2. skills for tools you edited
3. plugin.json
4. README.md
5. CLAUDE.md
6. tests/ — keep the pytest suite green and add coverage for behaviour you change

Bump the plugin version on every commit. Patch version for small fixes, minor version for more substantial changes (new skill or tool).

## Tests

The `bin/readoc`, `bin/readir`, and `bin/editdoc` CLIs are covered by a pytest
suite in `tests/`.
Dev tooling is managed by `pyproject.toml` (a dev-only virtual uv project,
`package = false`). Before changing any of these scripts, establish a green
baseline:

```bash
uv sync          # first time / after dependency changes
uv run pytest
```

Lint and format with `uv run ruff check .` and `uv run ruff format .`. Tests
generate their own fixtures in a temp dir and run the CLIs as subprocesses; see
`README.md` (Development / Testing) for details.

`hk.pkl` (via `hk`, provisioned by `mise.toml`) runs ruff-check, ruff-format,
pytest, the audits (vulture dead-code, jscpd duplication), gitleaks secret
scanning, and a large-file guard on pre-commit; the same checks run in CI
(`.github/workflows/ci.yml`) on push/PR to `master`. After changing the dev
workflow, keep all three in sync. Run the whole suite locally with
`hk run check --all`.

This repo follows Mick's dev-env standard (`dev-hooks:dev-env-setup`), version
stamped via `DEV_ENV_VERSION` in `mise.toml` (currently **v10**). Tool versions
are pinned reproducibly in the committed `mise.lock`; bump them with
`mise upgrade` (commit the diff). `uv` enforces a 4-day dependency cooldown
(`[tool.uv] exclude-newer`). Key dev tooling: hk + pkl (pre-commit), uv (Python
env, 3.12 in CI), ruff (lint/format), pytest, vulture (dead code), jscpd
(duplication, via npx), gitleaks (secrets). `.gitleaks.toml` allowlists
gitignored runtime paths so `gitleaks dir`'s whole-tree scan stays green.
