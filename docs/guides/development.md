# Development

This guide documents the development workflow for contributors to `ezxl`.

---

## Environment setup

Follow the [Configuration guide](configuration.md) for virtual environment creation, dependency installation, and pywin32 post-install. Then register the git hooks:

```bash
git config core.hooksPath .hooks
```

---

## Linting and formatting

`ezxl` uses [ruff](https://docs.astral.sh/ruff/) for both linting and formatting. The configuration is in `pyproject.toml` under `[tool.ruff]`.

```bash
# Check for lint errors
ruff check src/

# Auto-fix fixable errors
ruff check src/ --fix

# Format all source files
ruff format src/

# Check formatting without modifying files
ruff format src/ --check
```

Ruff enforces: pycodestyle errors and warnings (`E`, `W`), pyflakes (`F`), isort-compatible import sorting (`I`), flake8-bugbear (`B`), pyupgrade (`UP`), bandit security rules (`S`), and several others. See `[tool.ruff.lint]` in `pyproject.toml` for the full rule set.

---

## Type checking

Two type checkers are configured:

**ty** (fast, minimal):

```bash
ty check src/
```

**pyright** (comprehensive, detailed reporting):

```bash
pyright src/
```

Both are configured in `pyproject.toml`. `ty` is used for rapid feedback during development; `pyright` is used for thorough pre-commit checks. Neither is a substitute for the other — they catch different classes of errors.

!!! note "pywin32 type stubs"
`pywintypes` ships without complete type stubs. Attributes such as `pywintypes.TimeType` and `pywintypes.com_error` are accessed via `getattr` in the source and annotated with `# type: ignore[attr-defined]` where necessary. This is expected and does not indicate a type error in the library logic.

---

## Security scanning

Bandit static security analysis is included in the ruff rule set (`S` rules) and runs as part of `ruff check`. For a standalone Bandit report:

```bash
bandit -r src/ezxl/
```

The `S101` (assert usage) and `S106` (hardcoded passwords in function args) rules are suppressed globally because they produce false positives in test code and configuration defaults.

---

## Import layer contracts

[import-linter](https://import-linter.readthedocs.io/) is configured in `pyproject.toml` under `[tool.importlinter]` to enforce the layer dependency contracts. Run the checks with:

```bash
lint-imports
```

The package must be installed in editable mode (`pip install -e .`) before running import-linter, because import-linter 2.x requires the package to be importable from `sys.path`.

---

## Pre-commit hooks

The `.hooks/` directory contains shell scripts registered as git hooks via `git config core.hooksPath .hooks`.

The pre-commit hook runs:

1. `ruff check` — lint gate
2. `ruff format --check` — formatting gate

A commit is rejected if either check fails. Fix the issues and stage the changes before retrying.

---

## Conventional commits

All commit messages must follow the [Conventional Commits](https://www.conventionalcommits.org/en/v1.0.0/) specification. The pre-commit hook validates the message format.

```text
<type>(<scope>): <short description>

[optional body]

[optional footer]
```

Accepted types: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `chore`, `perf`, `ci`, `build`.

Examples:

```text
feat(core): add ExcelApp.wait_ready() timeout parameter
fix(gui): handle COM E_INVALIDARG in RibbonProxy.is_pressed()
docs: add pywinauto backend example to guides
chore: update ruff to 0.4.x
```

Breaking changes must include `BREAKING CHANGE:` in the footer:

```text
feat(exceptions)!: rename ComError to COMOperationError

BREAKING CHANGE: COMOperationError replaces ComError in all public raises.
```

---

## Version bump

The version string has a single source of truth: the `version` field in `pyproject.toml`. Do not edit `version.py` manually.

Use the version bump script to update the version in `pyproject.toml`, synchronise `version.py`, and update the README version badge atomically:

```bash
python .scripts/dev/update_version.py 0.2.0
```

After running the script, review the diff, commit with `chore: bump version to 0.2.0`, and tag the release:

```bash
git tag v0.2.0
git push origin main --tags
```

---

## Building a distribution

```bash
python -m build
```

Output wheels and sdist are written to `dist/`. Verify the package before publishing:

```bash
twine check dist/*
```
