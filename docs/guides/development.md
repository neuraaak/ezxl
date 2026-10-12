# How to work on ezxl locally

Use this guide to create a contributor-ready workspace with the current lint, typing, test, and packaging tooling.

## 🔧 Prerequisites

- A cloned EzXl repository
- `uv` installed locally
- Python 3.11 or later

## 📝 Steps

1. Sync the workspace with every contributor dependency group.

    === "pip"

        ```bash
        pip install -e . --group dev
        ```

    === "uv"

        ```bash
        uv sync --all-groups
        ```

    The groups follow [PEP 735](https://peps.python.org/pep-0735/): `test`, `docs`, and `dev` (which includes both). `pip install --group` needs pip 25.1 or later.

2. Install the git hooks managed by `pre-commit`.

    === "pip"

        ```bash
        pre-commit install
        ```

    === "uv"

        ```bash
        uv run pre-commit install
        ```

3. Run the formatter and linter.

    === "pip"

        ```bash
        ruff format src tests
        ruff check src tests
        ```

    === "uv"

        ```bash
        uv run poe format
        uv run poe lint
        ```

4. Run both configured type checkers.

    === "pip"

        ```bash
        pyright
        ty check
        ```

    === "uv"

        ```bash
        uv run poe typecheck
        ```

5. Enforce the import-layer contract.

    === "pip"

        ```bash
        lint-imports
        ```

    === "uv"

        ```bash
        uv run poe imports
        ```

6. Run the test suite, then build the distributable artifacts.

    === "pip"

        ```bash
        pytest -m "not excel"
        python -m build
        twine check dist/*
        ```

    === "uv"

        ```bash
        uv run poe test-ci
        uv run poe build
        uv run twine check dist/*
        ```

    `poe test-ci` is exactly what CI runs: it deselects the `excel` marker
    (those tests drive a live Excel instance over COM) and applies the
    coverage threshold.

7. Replay the whole gate in one command before opening a pull request.

    ```bash
    uv run poe check     # format-check -> lint -> typecheck -> imports -> test
    uv run poe security  # bandit + pip-audit
    ```

## Variations

If you only need the fast code-quality loop during feature work:

=== "pip"

    ```bash
    ruff check . --fix
    ty check
    ```

=== "uv"

    ```bash
    uv run poe lint-fix
    uv run poe typecheck
    ```

EzXl centralizes runtime-only `pywintypes` members in compatibility helpers so that `ty` and `pyright` can type-check the COM boundary without scattered local workarounds.

If you need to bump the release version, edit `version` under `[project]` in
`pyproject.toml` — it is the single source of truth — then sync `_version.py`
from it:

=== "pip"

    ```bash
    python .scripts/dev/update_version.py
    ```

=== "uv"

    ```bash
    uv run poe version
    ```

The `sync-version` pre-commit hook runs the same script, so a bump committed
without this step is corrected at commit time.

## ✅ Result

You have a local workspace aligned with the repository tooling, including formatting, linting, type checking, import contract validation, packaging, and installed git hooks.
