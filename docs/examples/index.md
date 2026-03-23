# Examples

Practical, copy-paste-ready examples for the most common `ezxl` use cases.

Each example is self-contained: copy the block, adjust the file paths, and run it.

---

## 1. Open a workbook and read a sheet

Open a workbook in a new (hidden) Excel process, read a range from a named sheet, then save and quit.

```python
from ezxl import ExcelApp

with ExcelApp(mode="dispatch", visible=False) as xl:
    wb = xl.open("C:/reports/sales_2026.xlsx")
    ws = wb.sheet("Q1")

    # Read all used cells as a 2D list
    data = ws.used_range.values
    headers = data[0]
    rows = data[1:]

    print(f"Sheet: {ws.name}")
    print(f"Columns: {headers}")
    print(f"Row count: {len(rows)}")

    # Read a single named cell
    total = ws.cell("F20").value
    print(f"Q1 total: {total}")

    wb.save()
# Excel quits automatically when the with block exits
```

---

## 2. Attach to a running Excel instance

Use `mode="attach"` when Excel is already open. ezxl binds to the existing process and leaves it running after the `with` block.

```python
from ezxl import ExcelApp, WorkbookNotFoundError

with ExcelApp(mode="attach") as xl:
    try:
        wb = xl.workbook("budget_2026.xlsx")
    except WorkbookNotFoundError:
        print("budget_2026.xlsx is not open in the current Excel session.")
        raise

    ws = wb.sheet("Actuals")
    ws.cell("B2").value = 145_000
    ws.cell("B3").formula = "=B2*1.1"
    wb.save()

# Excel keeps running; only the Python reference is released
```

!!! tip "List open workbooks"
If you are unsure of the workbook name, call `xl.workbook()` without arguments to get a proxy for the active workbook, or iterate `xl._get_app().Workbooks` via the raw COM object. The workbook name is the filename as shown in Excel's title bar, including the extension.

---

## 3. Execute a ribbon command via COM

`ExcelApp.gui.ribbon` wraps `Application.CommandBars` MSO methods. Use it to trigger any standard Excel ribbon button without navigating menus.

```python
from ezxl import ExcelApp

with ExcelApp(mode="attach") as xl:
    wb = xl.workbook("report.xlsx")

    # Trigger File > Save
    xl.gui.ribbon.execute("FileSave")

    # Check whether the command is available in the current state
    can_paste = xl.gui.ribbon.is_enabled("Paste")
    print(f"Paste enabled: {can_paste}")

    # Check whether a toggle command is active
    bold_on = xl.gui.ribbon.is_pressed("Bold")
    print(f"Bold active: {bold_on}")
```

Common MSO identifiers: `"FileSave"`, `"FileSaveAs"`, `"Copy"`, `"Paste"`, `"PasteValues"`, `"Undo"`, `"Redo"`, `"Bold"`, `"Italic"`.

---

## 4. Switch to the pywinauto backend for the ribbon

Replace the default COM ribbon backend with the pywinauto UI Automation backend. This is useful when the COM GUI surface is blocked or unavailable in the target environment. The rest of the `GUIProxy` surfaces continue to use their default COM implementations.

```python
from ezxl import ExcelApp, GUIProxy
from ezxl.gui.pywinauto import PywinautoRibbonBackend

with ExcelApp(mode="attach") as xl:
    # Bind the pywinauto backend to the exact Excel window managed by xl.
    # Passing hwnd prevents the backend from attaching to a different
    # Excel window if multiple instances are running simultaneously.
    ribbon_backend = PywinautoRibbonBackend(hwnd=xl.hwnd)

    gui = GUIProxy(xl, ribbon=ribbon_backend)

    # Execute via pywinauto keyboard shortcut (locale-independent)
    gui.ribbon.execute("FileSave")    # sends Ctrl+S
    gui.ribbon.execute("Bold")        # sends Ctrl+B

    # menu and dialog still use the default COM backends
    bars = gui.menu.list_bars()
    print(f"CommandBars: {bars[:3]}")
```

!!! warning "Limited MSO support"
`PywinautoRibbonBackend` maps MSO identifiers to keyboard shortcuts. Only a curated set is supported: `FileSave`, `Copy`, `Paste`, `Bold`, `Italic`, `Underline`, `Undo`, `Redo`. Passing an unmapped identifier raises `GUIOperationError`. State queries (`is_enabled`, `is_pressed`, `is_visible`) raise `NotImplementedError` — use `RibbonProxy` for those.

---

## 5. Convert xlsx to csv without Excel open

`xlsx_to_csv` operates entirely on closed files using polars and fastexcel. No running Excel process is required.

```python
from ezxl import xlsx_to_csv, read_excel

# Simple conversion — first sheet, comma separator
xlsx_to_csv(
    source="C:/data/transactions_2026.xlsx",
    dest="C:/output/transactions_2026.csv",
)

# Named sheet, semicolon separator (common in European locales)
xlsx_to_csv(
    source="C:/data/transactions_2026.xlsx",
    dest="C:/output/transactions_2026_eu.csv",
    sheet="Transactions",
    separator=";",
)

# Read the result back as a polars DataFrame for verification
df = read_excel("C:/data/transactions_2026.xlsx", sheet="Transactions")
print(df.shape)        # (rows, columns)
print(df.head())
```

---

## 6. Format a closed workbook with ExcelFormatter

`ExcelFormatter` applies formatting to an existing `.xlsx` file without opening Excel. Operations are buffered and written in a single pass when `save()` is called.

```python
from ezxl import ExcelFormatter

(
    ExcelFormatter("C:/output/report_2026.xlsx")
    # Header row: bold, large font, white text, blue background
    .column_width("A", 25)
    .column_width("B", 15)
    .column_width("C", 15)
    .row_height(1, 28)
    .font("A1:C1", bold=True, size=12, color="FFFFFF")
    .fill("A1:C1", "2E4F8A")
    .align("A1:C1", horizontal="center", vertical="center")
    # Data rows: thin border, wrap text in column A
    .border("A1:C50", style="thin")
    .align("A2:A50", wrap=True)
    # Save in place (overwrites the source file)
    .save()
)
```

To write the formatted result to a new path instead of overwriting the source:

```python
ExcelFormatter("C:/output/report_2026.xlsx") \
    .font("A1", bold=True) \
    .save("C:/output/report_2026_formatted.xlsx")
```

!!! note "Active sheet only"
`ExcelFormatter` operates on the active sheet of the workbook. To format multiple sheets, create one `ExcelFormatter` instance per target sheet and use `openpyxl` directly to activate each sheet before calling `save()`, or open a feature request for multi-sheet support.
