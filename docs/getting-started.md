# Getting started

This tutorial takes you from a fresh virtual environment to a first successful workbook read.

## 🔧 Prerequisites

- Windows with Microsoft Excel installed
- Python 3.11 or later
- Matching Python and Excel bitness

## 📝 Step 1 — Check Python and Excel bitness

Run the following command in the interpreter you plan to use:

```bash
python -c "import struct; print(struct.calcsize('P') * 8, 'bit')"
```

You should see `64 bit` on a standard Microsoft 365 installation. If the bitness does not match Excel, COM automation will fail before EzXl can open a workbook.

## 📝 Step 2 — Install EzXl

=== "PyPI"

    === "pip"

        ```bash
        python -m venv .venv
        .venv\Scripts\activate
        pip install ezxl
        ```

    === "uv"

        ```bash
        uv add ezxl
        ```

=== "From source"

    === "pip"

        ```bash
        git clone https://github.com/neuraaak/ezxl.git
        cd ezxl
        python -m venv .venv
        .venv\Scripts\activate
        pip install -e ".[dev]"
        ```

    === "uv"

        ```bash
        git clone https://github.com/neuraaak/ezxl.git
        cd ezxl
        uv sync --extra dev
        ```

=== "Offline wheels"

    ```bash
    python -m venv .venv
    .venv\Scripts\activate
    pip install --no-index --find-links C:\wheels ezxl
    ```

After installation, `python -c "import ezxl; print(ezxl.__version__)"` should print a version number.

## 📝 Step 3 — Run the pywin32 post-install step

Run the post-install script once in the virtual environment you just created:

```bash { .annotate }
python .venv/Scripts/pywin32_postinstall.py -install  # (1)!
```

1. Run this once per virtual environment, not once per machine.

!!! warning "🔧 Required once per virtual environment"
    If you skip this step, `win32com` imports can succeed while COM dispatch still fails later at runtime.

## 📝 Step 4 — Open a workbook and read a value

```python
from ezxl import ExcelApp

with ExcelApp(mode="dispatch", visible=False) as xl:
    workbook = xl.open("C:/data/report.xlsx")
    summary = workbook.sheet("Summary")
    revenue = summary.cell("B5").value
    print(revenue)
```

You should see the value from cell `B5` printed to the console.

## 📝 Step 5 — Save a change

```python
from ezxl import ExcelApp

with ExcelApp(mode="dispatch", visible=False) as xl:
    workbook = xl.open("C:/data/report.xlsx")
    summary = workbook.sheet("Summary")
    summary.cell("B6").value = 42_000
    workbook.save()
```

You should see the updated value in Excel the next time you open the workbook.

## ✅ What you built

You created a working EzXl environment, opened a workbook through COM, read a cell, and saved a change back to disk.

!!! danger "🔧 Stay on one thread"
    `ExcelApp` follows Excel's STA threading model. Create and use the same instance on the same thread.

## ➡️ Next steps

- [How to install ezxl](guides/configuration.md)
- [How to run tests locally](guides/testing.md)
- [API reference](api/index.md)
- [Examples](examples/index.md)
