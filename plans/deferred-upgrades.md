# Deferred upgrades

## ruff 0.15 → 0.16+ (held at 0.15.22 in uv.lock, 2026-10-29)

**Why:** ruff 0.16 enables a much larger default rule set (413 rules, up from 59) and formats
Markdown code blocks by default. This repo sets no `[tool.ruff.lint] select`, so ruff 0.17.2
reports 151 findings (RUF059 82, UP031 47, I001 8, BLE001 6, RUF100 3, ISC004 2, PLW1510 2,
FURB167 1). Fixing them is a code change, not a dependency bump.

**Manual steps (pick one):**

1. Keep today's rules: add `[tool.ruff.lint] select = ["E4", "E7", "E9", "F"]` (the pre-0.16
   default) and decide on Markdown formatting (`extend-exclude = ["*.md"]` to keep it off).
2. Adopt the new defaults: `uv lock -P ruff && uv sync`, then `uv run ruff check --fix .`, review
   the unsafe fixes by hand, and run `hk run check --all`.

CI's lint job runs `uv run ruff format --check .` over the whole tree, so under 0.16+ it also
checks the Markdown files' code blocks unless they are excluded.
