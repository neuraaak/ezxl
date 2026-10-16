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

3. Use a `poe` task when you want a preset test mode.

    === "pip"

        ```bash
        poe test-unit    # unit tests only
        poe test-fast    # parallel, skips slow and excel
        poe test-ci      # what CI runs, with the coverage gate
        ```

    === "uv"

        ```bash
        uv run poe test-unit
        uv run poe test-fast
        uv run poe test-ci
        ```

## 🧪 Current suite

- [x] Unit tests for converters, formatters, exceptions, exported symbols, and GUI protocols
- [x] Unit tests for the CLI (`click.testing.CliRunner`) and the COM error-translation helpers
- [x] Marker registration for `unit`, `integration`, `slow`, `cli`, and `excel`
- [ ] Committed Excel scenarios behind `@pytest.mark.excel`

??? note "🧪 Current repository scope"
     Every committed test lives in `tests/unit/`. The COM proxies in `core/` are
     excluded from the coverage gate because they need a running Excel; everything
     else is measured. A test that does need Excel is marked `@pytest.mark.excel`
     wherever it sits, which is what keeps it out of CI — the directory it lives in
     plays no part.

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
