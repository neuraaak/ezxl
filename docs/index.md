# EzXl

[![PyPI version](https://img.shields.io/pypi/v/ezxl?style=flat&logo=pypi&logoColor=white)](https://pypi.org/project/ezxl/)
[![Python versions](https://img.shields.io/pypi/pyversions/ezxl?style=flat&logo=python&logoColor=white)](https://pypi.org/project/ezxl/)
[![PyPI status](https://img.shields.io/pypi/status/ezxl?style=flat&logo=pypi&logoColor=white)](https://pypi.org/project/ezxl/)
[![License](https://img.shields.io/badge/license-MIT-green?style=flat&logo=github&logoColor=white)](https://github.com/neuraaak/ezxl/blob/main/LICENSE)
[![CI](https://img.shields.io/github/actions/workflow/status/neuraaak/ezxl/publish-pypi.yml?style=flat&label=publish&logo=githubactions&logoColor=white)](https://github.com/neuraaak/ezxl/actions/workflows/publish-pypi.yml)
[![Docs](https://img.shields.io/badge/docs-Github%20Pages-blue?style=flat&logo=materialformkdocs&logoColor=white)](https://neuraaak.github.io/ezxl/)
[![uv](https://img.shields.io/badge/package%20manager-uv-DE5FE9?style=flat&logo=uv&logoColor=white)](https://github.com/astral-sh/uv)
[![linter](https://img.shields.io/badge/linter-ruff-orange?style=flat&logo=ruff&logoColor=white)](https://github.com/astral-sh/ruff)
[![type checker](https://img.shields.io/badge/type%20checker-ty-orange?style=flat&logo=astral&logoColor=white)](https://github.com/astral-sh/ty)

![EzXl Logo](https://raw.githubusercontent.com/neuraaak/ezxl/refs/heads/main/docs/assets/logo-min.png)

**EzXl** is a Windows-first Python library for live Excel automation, closed-file conversion, and workbook formatting.

## 🚀 Quick start

=== "Live Excel session"

    === "pip"

        ```bash
        python -m venv .venv
        .venv\Scripts\activate
        pip install ezxl
        python .venv/Scripts/pywin32_postinstall.py -install
        ```

    === "uv"

        ```bash
        uv add ezxl
        uv run python .venv/Scripts/pywin32_postinstall.py -install
        ```

    ```python { .annotate }
    from ezxl import ExcelApp

    with ExcelApp(mode="dispatch", visible=False) as xl:
        workbook = xl.open("C:/reports/budget.xlsx")
        total = workbook.sheet("Summary").cell("B12").value  # (1)!
        print(total)
    ```

    1. Live COM automation requires Windows, Excel, and matching Python/Excel bitness.

=== "Closed-file I/O"

    ```python
    from ezxl import read_excel, xlsx_to_csv

    dataframe = read_excel("C:/reports/budget.xlsx", sheet="Summary")
    xlsx_to_csv("C:/reports/budget.xlsx", "C:/reports/budget.csv")
    print(dataframe.shape)
    ```

!!! tip "🚀 Fast path"
    Use the closed-file functions when you do not need a running Excel process.

## ✨ Key features

- Excel COM lifecycle management through `ExcelApp`, `WorkbookProxy`, and `SheetProxy`
- Closed-file read and conversion flows through `polars`, `fastexcel`, and `xlsxwriter`
- Closed-file formatting through `ExcelFormatter` and `openpyxl`
- Swappable GUI surfaces through `GUIProxy` and backend protocols
- Optional GUI navigation helpers for pywinauto-backed keystrokes and Backstage access

## 📚 Documentation

| Section                               | Description                                                        |
| :------------------------------------ | :----------------------------------------------------------------- |
| [Getting started](getting-started.md) | Install ezxl and complete a first workbook read in a few minutes.  |
| [Guides](guides/index.md)             | Task-oriented recipes for installation, development, and testing.  |
| [API reference](api/index.md)         | Curated overview of the public Python API.                         |
| [CLI reference](cli/index.md)         | Command and option tables for the `ezxl` executable.               |
| [Examples](examples/index.md)         | Copy-paste-ready snippets for common automation and I/O scenarios. |
| [Concepts](concepts/index.md)         | Design rationale behind the package split and backend strategy.    |
| [Architecture](architecture.md)       | Generated import graph and structural overview.                    |
| [Coverage](coverage.md)               | Coverage report page generated from `coverage.xml`.                |
| [Changelog](changelog.md)             | Release notes generated from Conventional Commits.                 |

## 📋 Requirements

- Python 3.11 or later
- Windows for COM automation and GUI helpers
- Microsoft Excel installed only for live COM scenarios
- Matching Python and Excel bitness for COM automation

## ⚖️ License

EzXl is distributed under the MIT license. See [LICENSE](https://github.com/neuraaak/ezxl/blob/main/LICENSE).
