# EzXl

[![PyPI version](https://img.shields.io/pypi/v/ezxm?style=flat&logo=pypi&logoColor=white)](https://pypi.org/project/ezxm/)
[![Python versions](https://img.shields.io/pypi/pyversions/ezxm?style=flat&logo=python&logoColor=white)](https://pypi.org/project/ezxm/)
[![PyPI status](https://img.shields.io/pypi/status/ezxm?style=flat&logo=pypi&logoColor=white)](https://pypi.org/project/ezxm/)
[![License](https://img.shields.io/badge/license-MIT-green?style=flat&logo=github&logoColor=white)](https://github.com/neuraaak/ezxm/blob/main/LICENSE)
[![CI](https://img.shields.io/github/actions/workflow/status/neuraaak/ezxm/publish-pypi.yml?style=flat&label=publish&logo=githubactions&logoColor=white)](https://github.com/neuraaak/ezxm/actions/workflows/publish-pypi.yml)
[![Docs](https://img.shields.io/badge/docs-Github%20Pages-blue?style=flat&logo=materialformkdocs&logoColor=white)](https://neuraaak.github.io/ezxm/)
[![uv](https://img.shields.io/badge/package%20manager-uv-DE5FE9?style=flat&logo=uv&logoColor=white)](https://github.com/astral-sh/uv)
[![linter](https://img.shields.io/badge/linter-ruff-orange?style=flat&logo=ruff&logoColor=white)](https://github.com/astral-sh/ruff)
[![type checker](https://img.shields.io/badge/type%20checker-ty-orange?style=flat&logo=astral&logoColor=white)](https://github.com/astral-sh/ty)

![Logo](docs/assets/logo-min.png)

**EzXl** is a lightweight Python library for Excel automation via **COM (win32com)** on Windows — open workbooks, manipulate sheets, interact with the ribbon, and convert between formats, all from a clean Python API.

## 📦 Installation

### Standard

```bash
# Create and activate a virtual environment
python -m venv .venv
.venv\Scripts\activate

# Install in development mode
pip install -e ".[dev]"
```

### pywin32 post-install step (mandatory)

> **Warning**: pywin32 requires a post-install script to register COM components. This step is mandatory and must be run once after installation. Skipping it will cause COM dispatch to fail.

```bash
python .venv/Scripts/pywin32_postinstall.py -install
```

### Corporate offline environment

In restricted environments without PyPI access, install from local wheel files only. No network requests are made during installation. Configure your pip to point to a local wheel directory:

```bash
pip install --no-index --find-links=\\share\wheels -e ".[dev]"
```

### Optional pywinauto backend

To enable the pywinauto GUI automation backend (UIA-based, locale-independent):

```bash
pip install -e ".[pywinauto]"
```

> **Note**: Ensure your Python interpreter bitness (32-bit or 64-bit) matches the installed Excel. COM dispatch will fail on a mismatch. Verify with:
>
> ```bash
> python -c "import struct; print(struct.calcsize('P') * 8)"
> ```

## 🚀 Quick Start

```python
from ezxl import ExcelApp

# Open a workbook and read from a sheet
with ExcelApp(mode="dispatch", visible=True) as xl:
    wb = xl.open("C:/reports/budget.xlsx")
    sheet = wb.sheet("Summary")
    value = sheet.cell("B2").value
    print(value)
    wb.save()

# Attach to a running Excel instance
with ExcelApp(mode="attach") as xl:
    xl.gui.ribbon.execute("FileSave")
```

## 🎯 Key Features

- **✅ COM Automation**: open, attach, navigate Excel via win32com
- **✅ GUI Interaction**: ribbon, menus, dialogs, and SendKeys via win32com
- **✅ Swappable GUI Backends**: COM or pywinauto, injected via protocol ABCs
- **✅ File I/O**: xlsx↔csv conversion via polars+fastexcel (no Excel required)
- **✅ Closed-file Formatting**: colors, fonts, and borders via openpyxl
- **✅ Thread Safety**: COM STA model enforced, thread identity checked at construction
- **✅ Full Type Hints**: complete typing for IDEs and linters
- **✅ Corporate Ready**: offline wheel install, proxy support

## 📚 Documentation

Full documentation is available online: **[neuraaak.github.io/ezxl](https://neuraaak.github.io/ezxl/)**

- **[📖 Getting Started](https://neuraaak.github.io/ezxl/getting-started/)** – Installation, first steps, and environment setup
- **[🎯 API Reference](https://neuraaak.github.io/ezxl/api/)** – Complete class and function reference (auto-generated)
- **[🏗️ Architecture](https://neuraaak.github.io/ezxl/architecture/)** – Design decisions and component overview
- **[💡 Examples](https://neuraaak.github.io/ezxl/examples/)** – Usage examples for common scenarios
- **[🔧 Development](https://neuraaak.github.io/ezxl/guides/development/)** – Environment setup and contribution guide
- **[🧪 Testing](https://neuraaak.github.io/ezxl/guides/testing/)** – Test suite documentation _(coming in next sprint)_

## 🧪 Testing

```bash
# Install dev dependencies
pip install -e ".[dev]"

# Run unit tests (no Excel required)
pytest tests/ -m "not excel"

# Run integration tests (requires local Excel)
pytest tests/ -m excel
```

> **Note**: Integration tests marked with `@pytest.mark.excel` require a locally installed Excel instance. They are excluded from CI/CD pipelines and are intended for local verification only.

## 🛠️ Development Setup

```bash
# Clone the repository
git clone https://github.com/neuraaak/ezxl.git
cd ezxl

# Create and activate a virtual environment
python -m venv .venv
.venv\Scripts\activate

# Install in development mode with all dev dependencies
pip install -e ".[dev]"

# Mandatory: register pywin32 COM components
python .venv/Scripts/pywin32_postinstall.py -install

# Set up git hooks
git config core.hooksPath .hooks
```

See the **[Development Guide](https://neuraaak.github.io/ezxl/guides/development/)** for detailed setup instructions.

## 🔌 API Overview

### 🖥️ COM Automation (5)

| Symbol          | Description                                       |
| --------------- | ------------------------------------------------- |
| `ExcelApp`      | COM session — dispatch or attach, context manager |
| `WorkbookProxy` | Workbook open/save/close, sheet access            |
| `SheetProxy`    | Sheet navigation, cell access                     |
| `CellProxy`     | Single cell read/write                            |
| `RangeProxy`    | Range selection and bulk operations               |

### 🎛️ GUI Interaction (4)

| Symbol        | Description                                |
| ------------- | ------------------------------------------ |
| `GUIProxy`    | Unified GUI facade with swappable backends |
| `RibbonProxy` | Ribbon execution and state via COM         |
| `MenuProxy`   | Legacy CommandBar traversal via COM        |
| `DialogProxy` | File pickers and alerts via COM            |

### 🔌 GUI Backends — pywinauto (4)

| Symbol                   | Description                                        |
| ------------------------ | -------------------------------------------------- |
| `PywinautoRibbonBackend` | Ribbon via keyboard shortcuts (locale-independent) |
| `PywinautoMenuBackend`   | Menu traversal via UIA                             |
| `PywinautoDialogBackend` | File pickers via UIA + Win32                       |
| `PywinautoKeysBackend`   | Keystroke injection via pywinauto                  |

### 📂 File I/O (5)

| Symbol           | Description                                     |
| ---------------- | ----------------------------------------------- |
| `read_excel`     | Read xlsx into a polars DataFrame               |
| `read_csv`       | Read csv into a polars DataFrame                |
| `xlsx_to_csv`    | Convert xlsx to csv (no Excel required)         |
| `csv_to_xlsx`    | Convert csv to xlsx (no Excel required)         |
| `ExcelFormatter` | Closed-file formatting (colors, fonts, borders) |

## 📦 Dependencies

- **pywin32>=306** — COM driver (`win32com.client`), Windows-only
- **polars>=1.0.0** — DataFrame I/O engine
- **fastexcel>=0.11.0** — Fast xlsx reader (Rust binding)
- **xlsxwriter>=3.0.0** — xlsx write backend for polars
- **openpyxl>=3.1.0** — Closed-file formatting
- **ezplog>=2.0.0** — Structured logging

Optional:

- **pywinauto>=0.6.8** — UI Automation backend (locale-independent GUI interaction)

## 📄 License

MIT License – See [LICENSE](LICENSE) file for details.

## 🔗 Links

- **Repository**: [https://github.com/neuraaak/ezxl](https://github.com/neuraaak/ezxl)
- **Issues**: [GitHub Issues](https://github.com/neuraaak/ezxl/issues)
- **Documentation**: [neuraaak.github.io/ezxl](https://neuraaak.github.io/ezxl/)
- **Changelog**: [docs/changelog.md](docs/changelog.md)

---

**ezxl** — Excel automation made simple, reliable, and Pythonic. 🐍
