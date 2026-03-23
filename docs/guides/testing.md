# How to run tests locally

Use this guide to run the current pytest suite, generate the coverage artifacts used by the docs build, and choose the right markers during local work.

## 🔧 Prerequisites

- A synced workspace with the `test` extra installed
- Windows and Excel only if you intend to opt into Excel-marked tests

## 📝 Steps

1. Run the default local suite.

    === "pip"

        ```bash
        pytest -m "not excel"
        ```

    === "uv"

        ```bash
        uv run pytest -m "not excel"
        ```

2. Generate the coverage files consumed by the documentation workflow.

    === "pip"

        ```bash
        pytest --cov=src/ezxl --cov-report=xml --cov-report=html -m "not excel"
        ```

    === "uv"

        ```bash
        uv run pytest --cov=src/ezxl --cov-report=xml --cov-report=html -m "not excel"
        ```

3. Use the helper script when you want a preset test mode.

    === "pip"

        ```bash
        python tests/run_tests.py --type unit --verbose
        ```

    === "uv"

        ```bash
        uv run python tests/run_tests.py --type unit --verbose
        ```

## 🧪 Current suite

- [x] Unit tests for converters, exceptions, exported symbols, and GUI protocols
- [x] Marker registration for `unit`, `integration`, `slow`, and `excel`
- [ ] Committed Excel integration scenarios in `tests/integration/`

??? note "🧪 Current repository scope"
     The `integration` and `excel` markers are already part of the test contract, but the committed suite in this repository is currently centered on `tests/unit/`. Keep the markers when you add new coverage so CI and local filtering stay consistent.

## ✏️ Add a new test

Mark each new test explicitly so selection stays predictable:

```python
import pytest


@pytest.mark.unit
def test_should_export_all_exception_symbols() -> None:
     ...
```

Use `@pytest.mark.excel` only for tests that require a live Excel installation, and keep those scenarios out of the default local run.

## ✅ Result

You can run the default suite, produce `coverage.xml` for the docs build, and choose markers that match the repository's existing test contract.
