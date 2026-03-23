# How to work on ezxl locally

Use this guide to create a contributor-ready workspace with the current lint, typing, test, and packaging tooling.

## 🔧 Prerequisites

- A cloned EzXl repository
- `uv` installed locally
- Python 3.11 or later

## 📝 Steps

1. Sync the workspace with all contributor extras.

    === "pip"

        ```bash
        pip install -e ".[dev,test,docs]"
        ```

    === "uv"

        ```bash
        uv sync --extra dev --extra test --extra docs
        ```

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
        ruff format .
        ruff check .
        ```

    === "uv"

        ```bash
        uv run ruff format .
        uv run ruff check .
        ```

4. Run both configured type checkers.

    === "pip"

        ```bash
        ty check
        pyright src/
        ```

    === "uv"

        ```bash
        uv run ty check
        uv run pyright src/
        ```

5. Enforce the import-layer contract.

    === "pip"

        ```bash
        lint-imports
        ```

    === "uv"

        ```bash
        uv run lint-imports
        ```

6. Build the distributable artifacts.

    === "pip"

        ```bash
        python -m build
        twine check dist/*
        ```

    === "uv"

        ```bash
        uv build
        uv run twine check dist/*
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
    uv run ruff check . --fix
    uv run ty check
    ```

EzXl centralizes runtime-only `pywintypes` members in compatibility helpers so that `ty` and `pyright` can type-check the COM boundary without scattered local workarounds.

If you need to bump the release version:

=== "pip"

    ```bash
    python .scripts/dev/update_version.py 1.2.0
    ```

=== "uv"

    ```bash
    uv run python .scripts/dev/update_version.py 1.2.0
    ```

## ✅ Result

You have a local workspace aligned with the repository tooling, including formatting, linting, type checking, import contract validation, packaging, and installed git hooks.
