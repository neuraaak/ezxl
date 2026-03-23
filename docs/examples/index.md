# Examples

Copy-paste-ready examples for the most common EzXl scenarios.

!!! tip "🚀 Assumption"
    All examples assume the package is already installed in the current environment.

## 🚀 Open a workbook and read a range

```python
from ezxl import ExcelApp

with ExcelApp(mode="dispatch", visible=False) as xl:
    workbook = xl.open("C:/reports/sales_2026.xlsx")
    sheet = workbook.sheet("Q1")
    rows = sheet.used_range.values
    print(rows[0])
    print(len(rows) - 1)
```

## 💡 Attach to a running Excel instance

```python
from ezxl import ExcelApp, WorkbookNotFoundError

with ExcelApp(mode="attach") as xl:
    try:
        workbook = xl.workbook("budget_2026.xlsx")
    except WorkbookNotFoundError:
        print("budget_2026.xlsx is not open.")
        raise

    actuals = workbook.sheet("Actuals")
    actuals.cell("B2").value = 145_000
    actuals.cell("B3").formula = "=B2*1.1"
    workbook.save()
```

## 💡 Execute a ribbon command

```python
from ezxl import ExcelApp

with ExcelApp(mode="attach") as xl:
    xl.gui.ribbon.execute("FileSave")
    print(xl.gui.ribbon.is_enabled("Paste"))
    print(xl.gui.ribbon.is_pressed("Bold"))
```

## 💡 Navigate Backstage with pywinauto

```python
from ezxl import ExcelApp, GUIProxy
from ezxl.gui.pywinauto import PywinautoBackstageBackend

with ExcelApp(mode="attach") as xl:
    backstage_nav = PywinautoBackstageBackend(hwnd=xl.hwnd, locale="fr")
    gui = GUIProxy(xl, backstage_nav=backstage_nav)

    gui.backstage.save_as(path="C:/reports/budget_2026_final.xlsx")

    if gui.backstage_nav is not None:
        gui.backstage_nav.open_save_as_panel()
        gui.backstage_nav.open_options()
```

## 💡 Convert xlsx to csv without Excel

```python
from ezxl import read_excel, xlsx_to_csv

xlsx_to_csv(
    source="C:/data/transactions_2026.xlsx",
    dest="C:/output/transactions_2026.csv",
    sheet="Transactions",
)

dataframe = read_excel("C:/data/transactions_2026.xlsx", sheet="Transactions")
print(dataframe.shape)
```

## 💡 Format a closed workbook

```python
from ezxl import ExcelFormatter

(
    ExcelFormatter("C:/output/report_2026.xlsx")
    .column_width("A", 25)
    .row_height(1, 28)
    .font("A1:C1", bold=True, size=12, color="FFFFFF")
    .fill("A1:C1", "2E4F8A")
    .align("A1:C1", horizontal="center", vertical="center")
    .border("A1:C50", style="thin")
    .save("C:/output/report_2026_formatted.xlsx")
)
```
