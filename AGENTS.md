# AGENTS.md — EzXl

Instructions for AI agents (Claude, Codex, Gemini, Copilot, etc.) working on this repository.

---

## Project Overview

**EzXl** is a Python library (3.11+) for Excel automation on Windows via COM (`win32com`), closed-file manipulation via `openpyxl`, and fast file conversion via `polars` / `fastexcel`. It also exposes a `click`-based CLI.

- **PyPI**: `ezxl` — [https://pypi.org/project/ezxl/](https://pypi.org/project/ezxl/)
- **Docs**: [https://neuraaak.github.io/ezxl/](https://neuraaak.github.io/ezxl/)
- **Package manager**: `uv`
- **Build backend**: `hatchling`

---

## Repository Layout

```text
src/ezxl/
├── __init__.py          # Public API, platform guard (win32 vs cross-platform)
├── _version.py          # Single source of truth for version string
├── exceptions.py        # All custom exceptions (EzXlError hierarchy)
├── cli/                 # Click CLI entry point (ezxl command)
│   ├── main.py
│   └── commands/
├── core/                # COM automation — Windows only
│   ├── _excel_app.py    # ExcelApp (opens/closes Excel via win32com)
│   ├── _workbook.py     # WorkbookProxy
│   └── _sheet.py        # SheetProxy, CellProxy, RangeProxy
├── gui/                 # GUI/ribbon automation — Windows only
│   ├── _protocols.py    # Abstract backends (ABCs)
│   ├── _gui_proxy.py    # GUIProxy dispatcher
│   ├── win32com/        # COM-based backends (ribbon, dialog, menu, backstage)
│   └── pywinauto/       # UIA-based backends (locale-independent)
├── io/                  # Cross-platform file I/O
│   ├── _converters.py   # read_excel, read_csv, xlsx_to_csv, csv_to_xlsx, read_sheet
│   └── _formatters.py   # ExcelFormatter (openpyxl closed-file formatting)
└── utils/               # Internal helpers
    ├── _com_utils.py    # COM helpers — Windows only
    └── _pywintypes_compat.py
tests/
├── unit/                # Fast, no Excel instance required
├── integration/         # Require a live Excel process (marked `excel`)
└── robustness/
```

---

## Import Layer Contract

Enforced by `import-linter` in CI. **Never violate this dependency order:**

```text
cli  →  core  →  gui  →  io | utils  →  exceptions
```

- `exceptions` has no internal imports.
- `io` and `utils` must not import from `core`, `gui`, or `cli`.
- `gui` must not import from `core` or `cli` — it types the application
  through the `ExcelAppLike` protocol instead.
- `core` may import `gui` (`ExcelApp.gui` builds a `GUIProxy`), and does so
  inside the property so the GUI backends are not pulled into every `core`
  import.
- Cross-cutting imports trigger a CI failure (`lint-imports`).

---

## Platform Guard

COM-dependent modules (`core`, `gui`, `utils/_com_utils.py`) are conditionally imported:

```python
if sys.platform == "win32":
    from .core._excel_app import ExcelApp
    ...
```

**Do not remove this guard.** Cross-platform utilities (`io`, `exceptions`) must remain importable on Linux/macOS (e.g., for CI running on `ubuntu-latest`).

---

## Development Setup

```bash
# Create venv and install all dev dependencies
uv sync --group dev

# Mandatory post-install step for pywin32 (Windows only, once)
python .venv/Scripts/pywin32_postinstall.py -install

# Install in editable mode (required for import-linter)
pip install -e .
```

### Corporate / offline environment

```bash
pip install --no-index --find-links=\\share\wheels -e . --group dev
```

No PyPI access is assumed in restricted environments. Use local `.whl` files.

---

## Code Conventions

### Style

- **Formatter**: `ruff format` (line length 88, double quotes, space indent)
- **Linter**: `ruff check` (E, W, F, I, B, C4, UP, S, T20, ARG, PIE, SIM)
- **Type hints**: mandatory on all public functions and methods
- **Docstrings**: Google style, on all public symbols
- **String formatting**: f-strings only
- **Paths**: `pathlib.Path`, never `os.path`
- **Logging**: `logging` module only — never `print()`
- **No global variables**, no commented-out code, no hard-coded paths or credentials

### Type Checking

Two tools are configured — both must pass:

| Tool      | Command           | Mode                        |
| --------- | ----------------- | --------------------------- |
| `pyright` | `uv run pyright`  | comprehensive, `basic` mode |
| `ty`      | `uv run ty check` | fast, minimal               |

Both run in order via `uv run poe typecheck`.

### Error Handling Pattern

```python
try:
    result = risky_operation()
except SpecificError as e:
    logger.error(f"Operation failed: {e}")
    raise
```

Only validate at system boundaries (user input, COM calls, file I/O). Do not add defensive checks for scenarios the type system already prevents.

---

## Testing

```bash
# What CI runs: excludes tests requiring live Excel, applies the 85% threshold
uv run poe test-ci

# Full suite, live-Excel tests included (Windows + Excel required)
uv run poe test

# Unit tests only
uv run poe test-unit

# Parallel, fast subset (no slow, no excel)
uv run poe test-fast
```

### Test Naming

```text
test_should_<expected_behavior>_when_<condition>
```

### Markers

| Marker        | Meaning                                          |
| ------------- | ------------------------------------------------ |
| `unit`        | Default — no external dependencies               |
| `integration` | Requires file system or real data                |
| `excel`       | Requires a live Excel process (excluded from CI) |
| `slow`        | Long-running tests                               |
| `cli`         | CLI-related tests                                |
| `robustness`  | Edge cases and fault tolerance                   |

Tests in `tests/unit/` must not require a running Excel instance. Mark any test that does with `@pytest.mark.excel` so it is excluded from CI.

---

## Linting & Security

```bash
# Lint and format check
uv run poe lint
uv run poe format-check

# Import layer contracts (the task carries PYTHONPATH=src itself)
uv run poe imports

# Security scan: bandit + pip-audit
uv run poe security

# The whole gate, as CI runs it
uv run poe check
```

Every CI step runs a `poe` task from `[tool.poe.tasks]`, so the local command
and the CI command are the same string. Run `uv run poe --help` for the list.

Pre-commit hooks run ruff + bandit automatically on `git commit`.

---

## CI/CD Pipeline

Numbered reusable workflows, one responsibility each, chained on push to `main`:

```text
00-install-deps  # primes the uv cache (keyed on uv.lock)
01-ci            # lint | type-check | security -> test matrix; the quality gate
                 #   the matrix runs on Ubuntu and Windows
02-tag-sync      # orchestrator: CI gate -> tag_action -> cascade 03 then 04
03-publish       # build always; upload to PyPI on `create` only (OIDC)
04-docs          # build always; deploy latest / dev / none by tag_action
```

`02-tag-sync` derives a three-state `tag_action` (`create` | `skip` | `preview`)
and calls `03`/`04` directly — see `.github/workflows/README.md` for the full
decision table.

**Version source of truth**: `pyproject.toml` → `[project] version`. The `src/ezxl/_version.py` file is synced from it by the pre-commit hook.

To release: bump `version` in `pyproject.toml` and push to `main`. The pipeline handles the rest.

---

## Public API Rules

The public API is defined in `src/ezxl/__init__.py` via `__all__`. When adding new public symbols:

1. Export from the appropriate submodule.
2. Add to `__all__` in `__init__.py` (inside the `win32` guard if Windows-only).
3. Add or update docstrings.
4. Add or update tests.

Do not expose internal helpers (prefixed `_`) in `__all__`.

---

## Commit Convention

Conventional Commits format is required (enforced by `cliff.toml` for changelog generation):

```text
feat: add xlsx_to_parquet converter
fix: handle empty sheet in read_sheet
refactor: extract COM retry logic to _com_utils
docs: update ExcelApp docstring
test: add robustness tests for CellProxy
chore: bump version to 1.2.0
```

Scope is optional. Breaking changes: append `!` after the type (`feat!: ...`) and add a `BREAKING CHANGE:` footer.

---

## What Agents Should NOT Do

- Do not add `print()` statements — use `logging`.
- Do not remove the `sys.platform == "win32"` guard in `__init__.py`.
- Do not import `core` or `gui` from `io` or `utils` (violates layer contract).
- Do not commit commented-out code.
- Do not hard-code file paths or credentials.
- Do not run `pytest` without the `-m "not excel"` flag unless a live Excel instance is confirmed available.
- Do not publish to PyPI manually — the CI pipeline owns this.
- Do not amend pushed commits on `main`.
