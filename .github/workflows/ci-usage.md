# CI Workflow

`ci.yml` runs on every push (all branches), every pull request, and manually via
`workflow_dispatch`. It is the primary quality gate — release workflows never run
if this one is red.

## Triggers

| Event               | Branches |
| ------------------- | -------- |
| `push`              | all      |
| `pull_request`      | all      |
| `workflow_dispatch` | manual   |

Concurrent runs on the same ref are cancelled automatically (`cancel-in-progress: true`,
group `ci-$ref`). Deploy jobs are never affected because they live in separate workflows.

## Jobs

```text
lint ──┐
       ├──► test (Python 3.13)
type-check ──┘
```

`lint` and `type-check` run in parallel; `test` needs both.

### `lint` — Lint & Format Check

Timeout: 15 min. Steps:

1. `uv sync --frozen --extra dev`
2. `ruff check .` — lint
3. `ruff format --check .` — format check (read-only, never auto-fixes)

### `type-check` — Type Check & Import Contracts

Timeout: 15 min. Steps:

1. `uv sync --frozen --extra dev`
2. `uv run ty check` — type check (configured via `[tool.ty]` in `pyproject.toml`)
3. `uv run pyright` — type check complémentaire (configured via `[tool.pyright]`)
4. `PYTHONPATH=src uv run lint-imports` — import-linter contracts (couches hexagonales)

### `test` — pytest

Needs: `[lint, type-check]`. Timeout: 25 min. Python **3.13 only** (`fail-fast: false`).

Steps:

1. `uv sync --frozen --extra dev`
2. `uv run pytest` — reads `pyproject.toml`, applies `--cov-fail-under=70`
3. Upload `coverage.xml` + `htmlcov/` as artifact — **always** (even on failure), retained 7 days

## Local equivalent

```bash
# Quality gate
uv run ruff check .
uv run ruff format --check .
uv run ty check
uv run pyright
PYTHONPATH=src uv run lint-imports

# Tests
uv run pytest
```

## Notes

- The install is always frozen (`uv sync --frozen`). Never mutates `uv.lock`.
- `ruff check` and `ruff format --check` are read-only — they never auto-fix.
- Both `ty` and `pyright` are configured in `pyproject.toml` and run in CI.
- Import contracts are enforced by import-linter against the hexagonal layer
  contracts defined in `[tool.importlinter]`. Adding a cross-layer import
  without updating the contracts will fail this job.
