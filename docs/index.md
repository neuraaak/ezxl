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

**EzXl** is a lightweight Python library for Excel automation via **COM (win32com)** on Windows — open workbooks, manipulate sheets, interact with the ribbon, and convert between formats, all from a clean Python API.

---

## Features

**COM automation**
: Open, attach, navigate, and control Excel via `win32com`. All raw COM errors are caught at the boundary and re-raised as typed Python exceptions. Thread identity is enforced at every call site.

**GUI interaction with swappable backends**
: Ribbon commands, legacy CommandBars menus, file-picker dialogs, and keystroke injection are exposed through a unified `GUIProxy` facade. Each surface (ribbon, menu, dialog, keys) accepts an alternative backend at construction time. The default backend uses COM; a `pywinauto`-based backend is provided for scenarios where COM GUI access is unavailable.

**File I/O without Excel**
: Read `.xlsx` and `.csv` files into [polars](https://pola.rs) DataFrames, convert between formats, or extract data as a row-major list — all without a running Excel process.

**Closed-file formatting**
: Apply column widths, row heights, fonts, fills, borders, and alignment to a saved workbook file via a fluent `ExcelFormatter` API backed by `openpyxl`.

---

## Requirements

- Windows only (COM dependency)
- Python 3.11 or later
- Excel bitness must match Python bitness (32-bit Python requires 32-bit Excel)
- `pywinauto` backends are optional — the COM layer has no dependency on them

---

## Installation

=== "Standard"

```bash
pip install ezxl
```

=== "Corporate (offline wheels)"

```bash
pip install --no-index --find-links /path/to/wheels ezxl
```

=== "With pywinauto backends"

```bash
pip install ezxl pywinauto
```

After installing on Windows, run the pywin32 post-install step:

```bash
python .venv/Scripts/pywin32_postinstall.py -install
```

---

## Quick usage

```python
from ezxl import ExcelApp

# Open a workbook, read a cell, and save — then quit Excel automatically.
with ExcelApp(mode="dispatch", visible=False) as xl:
    wb = xl.open("C:/reports/budget.xlsx")
    ws = wb.sheet("Summary")
    total = ws.cell("B12").value
    print(f"Total: {total}")
    wb.save()
```

---

## Where to go next

- [Getting Started](getting-started.md) — installation, bitness check, first steps
- [API Reference](api/index.md) — complete public API grouped by category
- [Examples](examples/index.md) — copy-paste-ready code for common scenarios
- [Architecture](architecture.md) — module layout and design decisions
