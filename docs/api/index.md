# API reference

Curated index of the public `ezxl` API exported from the top-level package.

## 📦 COM automation

| Symbol                                                   | Description                                                     |
| :------------------------------------------------------- | :-------------------------------------------------------------- |
| [`ExcelApp`](reference/index.md#ezxl.ExcelApp)           | Entry point for a live Excel COM session.                       |
| [`WorkbookProxy`](reference/index.md#ezxl.WorkbookProxy) | Workbook-level operations such as open, save, and sheet lookup. |
| [`SheetProxy`](reference/index.md#ezxl.SheetProxy)       | Worksheet navigation plus cell and range access.                |
| [`CellProxy`](reference/index.md#ezxl.CellProxy)         | Single-cell read and write access.                              |
| [`RangeProxy`](reference/index.md#ezxl.RangeProxy)       | Rectangular range access for bulk values.                       |

## 📦 GUI layer

| Symbol                                                                           | Description                                                            |
| :------------------------------------------------------------------------------- | :--------------------------------------------------------------------- |
| [`GUIProxy`](reference/index.md#ezxl.GUIProxy)                                   | Unified facade for ribbon, menu, dialog, keys, and Backstage surfaces. |
| [`RibbonProxy`](reference/index.md#ezxl.RibbonProxy)                             | COM-backed MSO ribbon command execution and state queries.             |
| [`MenuProxy`](reference/index.md#ezxl.MenuProxy)                                 | COM-backed legacy CommandBar traversal.                                |
| [`DialogProxy`](reference/index.md#ezxl.DialogProxy)                             | COM-backed file open, file save, and alert dialogs.                    |
| [`COMBackstageBackend`](reference/index.md#ezxl.COMBackstageBackend)             | COM-backed file operations in Excel Backstage.                         |
| [`PywinautoKeysBackend`](reference/index.md#ezxl.PywinautoKeysBackend)           | Optional pywinauto keystroke backend.                                  |
| [`PywinautoBackstageBackend`](reference/index.md#ezxl.PywinautoBackstageBackend) | Optional pywinauto Backstage navigator.                                |

## 📦 GUI contracts

| Symbol                                                                             | Description                                       |
| :--------------------------------------------------------------------------------- | :------------------------------------------------ |
| [`AbstractRibbonBackend`](reference/index.md#ezxl.AbstractRibbonBackend)           | Contract for ribbon execution and state queries.  |
| [`AbstractMenuBackend`](reference/index.md#ezxl.AbstractMenuBackend)               | Contract for legacy menu traversal.               |
| [`AbstractDialogBackend`](reference/index.md#ezxl.AbstractDialogBackend)           | Contract for file-picker and alert dialogs.       |
| [`AbstractKeysBackend`](reference/index.md#ezxl.AbstractKeysBackend)               | Contract for key injection.                       |
| [`AbstractBackstageFileOps`](reference/index.md#ezxl.AbstractBackstageFileOps)     | Contract for COM-style Backstage file operations. |
| [`AbstractBackstageNavigator`](reference/index.md#ezxl.AbstractBackstageNavigator) | Contract for UIA-style Backstage navigation.      |
| AbstractBackstageBackend                                                           | Compatibility alias kept for existing imports.    |

## 📦 File I/O and formatting

| Symbol                                                     | Description                                             |
| :--------------------------------------------------------- | :------------------------------------------------------ |
| [`read_excel`](reference/index.md#ezxl.read_excel)         | Read a workbook sheet into a polars DataFrame.          |
| [`read_csv`](reference/index.md#ezxl.read_csv)             | Read a CSV file into a polars DataFrame.                |
| [`xlsx_to_csv`](reference/index.md#ezxl.xlsx_to_csv)       | Convert a workbook sheet to CSV.                        |
| [`csv_to_xlsx`](reference/index.md#ezxl.csv_to_xlsx)       | Convert a CSV file to XLSX.                             |
| [`read_sheet`](reference/index.md#ezxl.read_sheet)         | Read a sheet into a legacy row-major list of lists.     |
| [`ExcelFormatter`](reference/index.md#ezxl.ExcelFormatter) | Apply formatting to a closed workbook through openpyxl. |

## 📦 Exceptions

| Symbol                                                                           | Description                                         |
| :------------------------------------------------------------------------------- | :-------------------------------------------------- |
| [`EzXlError`](reference/index.md#ezxl.EzXlError)                                 | Base exception for all library-originated failures. |
| [`ExcelNotAvailableError`](reference/index.md#ezxl.ExcelNotAvailableError)       | Excel could not be dispatched or attached.          |
| [`ExcelSessionLostError`](reference/index.md#ezxl.ExcelSessionLostError)         | A previously attached COM session was lost.         |
| [`ExcelThreadViolationError`](reference/index.md#ezxl.ExcelThreadViolationError) | A COM call was made from the wrong thread.          |
| [`WorkbookNotFoundError`](reference/index.md#ezxl.WorkbookNotFoundError)         | The requested workbook is not open.                 |
| [`SheetNotFoundError`](reference/index.md#ezxl.SheetNotFoundError)               | The requested sheet does not exist.                 |
| [`COMOperationError`](reference/index.md#ezxl.COMOperationError)                 | A COM call failed without a more specific mapping.  |
| [`GUIOperationError`](reference/index.md#ezxl.GUIOperationError)                 | A GUI surface call failed.                          |
| [`FormatterError`](reference/index.md#ezxl.FormatterError)                       | A closed-file formatting operation failed.          |

## 🔍 Full reference

For the complete mkdocstrings dump of the public API, see [Full reference](reference/index.md).

## 📦 Backend modules

| Module page                                           | Description                                                                                                       |
| :---------------------------------------------------- | :---------------------------------------------------------------------------------------------------------------- |
| [Win32com backends](reference/win32com-backends.md)   | Module-level reference for COM GUI backends (`RibbonProxy`, `MenuProxy`, `DialogProxy`, `COMBackstageBackend`).   |
| [Pywinauto backends](reference/pywinauto-backends.md) | Module-level reference for optional UI Automation backends (`PywinautoKeysBackend`, `PywinautoBackstageBackend`). |
