# Auto-Generated Reference

This page renders the full public API from source docstrings via mkdocstrings.
All symbols are importable from the top-level `ezxl` package.

---

## Exceptions

::: ezxl.EzXlError
options:
show_source: false

::: ezxl.ExcelNotAvailableError
options:
show_source: false

::: ezxl.ExcelSessionLostError
options:
show_source: false

::: ezxl.ExcelThreadViolationError
options:
show_source: false

::: ezxl.WorkbookNotFoundError
options:
show_source: false

::: ezxl.SheetNotFoundError
options:
show_source: false

::: ezxl.COMOperationError
options:
show_source: false

::: ezxl.GUIOperationError
options:
show_source: false

::: ezxl.FormatterError
options:
show_source: false

---

## COM automation

::: ezxl.ExcelApp
options:
show_source: false
members: - **init** - **enter** - **exit** - open - workbook - run_macro - execute_ribbon - wait_ready - quit - gui - hwnd

::: ezxl.WorkbookProxy
options:
show_source: false
members: - name - sheets - sheet - save - save_as - close

::: ezxl.SheetProxy
options:
show_source: false
members: - name - used_range - cell - range - calculate

::: ezxl.CellProxy
options:
show_source: false
members: - address - value - formula

::: ezxl.RangeProxy
options:
show_source: false
members: - address - values

---

## GUI interaction

::: ezxl.GUIProxy
options:
show_source: false
members: - **init** - ribbon - menu - dialog - backstage - backstage_nav - send_keys

---

## GUI protocols (ABCs)

::: ezxl.AbstractRibbonBackend
options:
show_source: false
members: - execute - is_enabled - is_pressed - is_visible

::: ezxl.AbstractMenuBackend
options:
show_source: false
members: - click - list_bars - list_controls

::: ezxl.AbstractDialogBackend
options:
show_source: false
members: - get_file_open - get_file_save - alert

::: ezxl.AbstractKeysBackend
options:
show_source: false
members: - send_keys

::: ezxl.AbstractBackstageFileOps
options:
show_source: false
members: - save - save_as - open_file - close_workbook

::: ezxl.AbstractBackstageNavigator
options:
show_source: false
members: - open_options - open_save_as_panel - open_file - close_workbook

---

## GUI backends — COM (win32com)

::: ezxl.RibbonProxy
options:
show_source: false
members: - execute - is_enabled - is_pressed - is_visible

::: ezxl.MenuProxy
options:
show_source: false
members: - click - list_bars - list_controls

::: ezxl.DialogProxy
options:
show_source: false
members: - get_file_open - get_file_save - alert

::: ezxl.COMBackstageBackend
options:
show_source: false
members: - save - save_as - open_file - close_workbook

---

## GUI backends — pywinauto

::: ezxl.PywinautoKeysBackend
options:
show_source: false
members: - send_keys

::: ezxl.PywinautoBackstageBackend
options:
show_source: false
members: - **init** - open_options - open_save_as_panel - open_file - close_workbook

---

## File I/O

::: ezxl.read_excel
options:
show_source: false

::: ezxl.read_csv
options:
show_source: false

::: ezxl.xlsx_to_csv
options:
show_source: false

::: ezxl.csv_to_xlsx
options:
show_source: false

::: ezxl.read_sheet
options:
show_source: false

---

## Closed-file formatting

::: ezxl.ExcelFormatter
options:
show_source: false
members: - **init** - column_width - row_height - font - fill - border - align - save
