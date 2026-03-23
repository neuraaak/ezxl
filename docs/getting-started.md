# Getting Started

This page gets you from a fresh Python environment to a working Excel automation session in under five minutes.

---

## Prerequisites

| Requirement        | Detail                                                              |
| ------------------ | ------------------------------------------------------------------- |
| Operating system   | Windows only — COM is a Windows-exclusive technology                |
| Python version     | 3.11 or later                                                       |
| Excel bitness      | Must match Python bitness exactly (both 64-bit or both 32-bit)      |
| Excel installation | Required for COM features; not required for file I/O and formatting |

---

## Verify Python bitness

Before installing, confirm that the Python interpreter you are using matches your Excel installation:

```bash
python -c "import struct; print(struct.calcsize('P') * 8, 'bit')"
```

The output should match your Excel bitness. Microsoft 365 is 64-bit by default. A bitness mismatch causes COM dispatch to fail with a `pythoncom` registration error that is difficult to diagnose without this check.

---

## Installation

=== "Standard (PyPI)"

```bash
    pip install ezxl
```

=== "Development mode (from source)"

```bash
    git clone https://github.com/neuraaak/ezxl.git
    cd ezxl
    python -m venv .venv
    .venv\Scripts\activate
    pip install -e ".[dev]"
```

=== "Corporate (offline wheels)"

In restricted environments with no PyPI access, install from a local wheel directory. All wheels must be pre-downloaded for `ezxl` and its dependencies (`pywin32`, `polars`, `fastexcel`, `openpyxl`, `xlsxwriter`).

````bash
    pip install --no-index --find-links C:\wheels ezxl
    ```

    If a proxy is required for any network operation:

    ```bash
    set HTTPS_PROXY=http://proxy.corp.example.com:8080
    pip install --proxy http://proxy.corp.example.com:8080 ezxl
````

---

## pywin32 post-install step

After any installation that includes `pywin32`, you must run its post-install script once. This step registers COM components and sets up registry entries that `win32com` depends on at runtime.

```bash
python .venv/Scripts/pywin32_postinstall.py -install
```

!!! warning "Required on every new virtual environment"
This step is not automatic. Forgetting it results in `ImportError: No module named 'pywintypes'` or silent COM dispatch failures. Run it once per virtual environment, not once per machine.

---

## Optional: pywinauto backends

The pywinauto GUI backends are an optional extension. Install `pywinauto` separately if you intend to use `PywinautoRibbonBackend`, `PywinautoMenuBackend`, `PywinautoDialogBackend`, or `PywinautoKeysBackend`:

```bash
pip install pywinauto
```

The COM layer (`ExcelApp`, `WorkbookProxy`, `GUIProxy` with default backends) has no dependency on `pywinauto`. You can use the full COM automation surface without it.

---

## First steps

### 1. Open a workbook

```python
from ezxl import ExcelApp

with ExcelApp(mode="dispatch", visible=True) as xl:
    wb = xl.open("C:/data/report.xlsx")
    print(wb.name)          # "report.xlsx"
    print(wb.sheets)        # ["Sheet1", "Data", "Summary"]
```

`ExcelApp` used as a context manager starts Excel on entry and quits it on exit. In `dispatch` mode, a new Excel process is launched. In `attach` mode (shown below), an existing process is reused and left running after the `with` block.

### 2. Read a cell

```python
from ezxl import ExcelApp

with ExcelApp(mode="dispatch", visible=False) as xl:
    wb = xl.open("C:/data/report.xlsx")
    ws = wb.sheet("Summary")

    # Single cell
    revenue = ws.cell("B5").value
    print(f"Revenue: {revenue}")

    # Range — returns list[list[Any]]
    table = ws.range("A1:D10").values
    headers = table[0]
    rows = table[1:]
```

COM date values in cells are automatically converted to `datetime` objects. Excel error cells (`#N/A`, `#VALUE!`, etc.) are returned as `None` with a warning logged.

### 3. Write a value and save

```python
from ezxl import ExcelApp

with ExcelApp(mode="dispatch", visible=False) as xl:
    wb = xl.open("C:/data/report.xlsx")
    ws = wb.sheet("Summary")

    ws.cell("B5").value = 42_000
    wb.save()
    wb.close(save=False)    # already saved; close without re-saving
```

### 4. Attach to a running Excel instance

Use `mode="attach"` when Excel is already open and you do not want ezxl to manage the process lifecycle. The Excel window remains open after the `with` block exits.

```python
from ezxl import ExcelApp

with ExcelApp(mode="attach") as xl:
    wb = xl.workbook("report.xlsx")     # must already be open
    ws = wb.sheet("Data")
    ws.cell("A1").value = "Updated"
    wb.save()
    # Excel keeps running after this block
```

---

## Threading note

!!! danger "ExcelApp is not thread-safe"
Excel COM uses the Single-Threaded Apartment (STA) model. An `ExcelApp` instance records the thread it was created on and raises `ExcelThreadViolationError` immediately if any method is called from a different thread. Always create and use an `ExcelApp` instance on the same thread. Do not share instances across threads.
