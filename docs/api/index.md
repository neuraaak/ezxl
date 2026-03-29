# API Reference

Complete public API for `ezxl`, grouped by functional category. All symbols are importable directly from the top-level `ezxl` package.

For auto-generated reference pages rendered from docstrings, see [Auto-Generated Reference](reference/index.md).

---

## COM automation

Core classes for managing an Excel COM session, navigating workbooks, and reading or writing cell data.

| Symbol                                                   | Description                                                                                   |
| -------------------------------------------------------- | --------------------------------------------------------------------------------------------- |
| [`ExcelApp`](reference/index.md#ezxl.ExcelApp)           | COM session lifecycle manager. Dispatches or attaches to an Excel process. Context manager.   |
| [`WorkbookProxy`](reference/index.md#ezxl.WorkbookProxy) | Proxy for a single open workbook. Provides `save`, `save_as`, `close`, `sheet`, and `sheets`. |
| [`SheetProxy`](reference/index.md#ezxl.SheetProxy)       | Proxy for a worksheet. Provides `cell`, `range`, `used_range`, and `calculate`.               |
| [`CellProxy`](reference/index.md#ezxl.CellProxy)         | Proxy for a single cell. Exposes `value`, `formula`, and `address`.                           |
| [`RangeProxy`](reference/index.md#ezxl.RangeProxy)       | Proxy for a rectangular cell range. Exposes `values` (read/write) and `address`.              |

---

## GUI interaction

Facade and backend classes for ribbon commands, menus, dialogs, keystroke injection, and Backstage file operations.

| Symbol                                                               | Description                                                                                                                                                                          |
| -------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| [`GUIProxy`](reference/index.md#ezxl.GUIProxy)                       | Unified GUI facade. Bundles ribbon, menu, dialog, keys, backstage (COM), and optional backstage_nav (UIA). Accepts alternative backends at construction time.                        |
| [`RibbonProxy`](reference/index.md#ezxl.RibbonProxy)                 | COM-based ribbon backend. Executes MSO commands and queries their state via `Application.CommandBars`.                                                                               |
| [`MenuProxy`](reference/index.md#ezxl.MenuProxy)                     | COM-based menu backend. Traverses legacy CommandBars by caption path.                                                                                                                |
| [`DialogProxy`](reference/index.md#ezxl.DialogProxy)                 | COM-based dialog backend. Provides `get_file_open`, `get_file_save`, and `alert`.                                                                                                    |
| [`COMBackstageBackend`](reference/index.md#ezxl.COMBackstageBackend) | COM-based Backstage backend. Implements file operations (`save`, `save_as`, `open_file`, `close_workbook`) via the Excel COM object model. Default backend for `GUIProxy.backstage`. |

---

## GUI protocols (ABCs)

Abstract base classes that define the contract for each GUI surface. Implement these to create a custom backend.

| Symbol                                                                             | Description                                                                                                                                              |
| ---------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------- |
| [`AbstractRibbonBackend`](reference/index.md#ezxl.AbstractRibbonBackend)           | Contract for ribbon execution and state queries.                                                                                                         |
| [`AbstractMenuBackend`](reference/index.md#ezxl.AbstractMenuBackend)               | Contract for CommandBar traversal and control execution.                                                                                                 |
| [`AbstractDialogBackend`](reference/index.md#ezxl.AbstractDialogBackend)           | Contract for file-open, file-save, and alert dialogs.                                                                                                    |
| [`AbstractKeysBackend`](reference/index.md#ezxl.AbstractKeysBackend)               | Contract for keystroke injection.                                                                                                                        |
| [`AbstractBackstageFileOps`](reference/index.md#ezxl.AbstractBackstageFileOps)     | Contract for Backstage file operations via COM (`save`, `save_as`, `open_file`, `close_workbook`). Implemented by `COMBackstageBackend`.                 |
| [`AbstractBackstageNavigator`](reference/index.md#ezxl.AbstractBackstageNavigator) | Contract for Backstage UIA navigation (`open_options`, `open_save_as_panel`, `open_file`, `close_workbook`). Implemented by `PywinautoBackstageBackend`. |

---

## GUI backends — pywinauto

Optional backends that operate at the OS UI Automation level instead of via COM. Require `pywinauto` to be installed separately.

| Symbol                                                                           | Description                                                                                                                                                                                                                                              |
| -------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| [`PywinautoKeysBackend`](reference/index.md#ezxl.PywinautoKeysBackend)           | Keystroke injection via `pywinauto.keyboard.send_keys`. No STA constraint.                                                                                                                                                                               |
| [`PywinautoBackstageBackend`](reference/index.md#ezxl.PywinautoBackstageBackend) | UIA-based Backstage navigator. Implements `AbstractBackstageNavigator`. Opens the Options panel, Save As panel, and other Backstage views via UI Automation direct click with Alt-sequence fallback. Requires `hwnd` to target the correct Excel window. |

---

## File I/O

Functions for reading and converting files without a running Excel process. All functions use polars backed by `fastexcel` (Rust) for Excel I/O.

| Symbol                                                     | Description                                                                                                        |
| ---------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| [`read_excel`](reference/index.md#ezxl.read_excel)         | Read an `.xlsx` / `.xlsm` sheet into a polars DataFrame.                                                           |
| [`read_csv`](reference/index.md#ezxl.read_csv)             | Read a `.csv` file into a polars DataFrame.                                                                        |
| [`xlsx_to_csv`](reference/index.md#ezxl.xlsx_to_csv)       | Convert an Excel sheet to a CSV file.                                                                              |
| [`csv_to_xlsx`](reference/index.md#ezxl.csv_to_xlsx)       | Convert a CSV file to an `.xlsx` workbook.                                                                         |
| [`read_sheet`](reference/index.md#ezxl.read_sheet)         | Read a sheet as a row-major `list[list[Any]]`. Compatibility shim for pre-polars callers.                          |
| [`ExcelFormatter`](reference/index.md#ezxl.ExcelFormatter) | Fluent closed-file formatter via openpyxl. Buffers operations and applies them in a single write pass on `save()`. |

---

## Exceptions

All exceptions inherit from `EzXlError`. Catching `EzXlError` handles any library-originated failure. Catch specific subclasses for fine-grained error handling.

| Symbol                                                                           | When raised                                                                            |
| -------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------- |
| [`EzXlError`](reference/index.md#ezxl.EzXlError)                                 | Base class for all EzXl exceptions.                                                    |
| [`ExcelNotAvailableError`](reference/index.md#ezxl.ExcelNotAvailableError)       | `mode="attach"` and no Excel instance is running, or COM registration is broken.       |
| [`ExcelSessionLostError`](reference/index.md#ezxl.ExcelSessionLostError)         | An established COM connection was lost mid-operation (Excel crashed or was closed).    |
| [`ExcelThreadViolationError`](reference/index.md#ezxl.ExcelThreadViolationError) | A COM call was attempted from a thread other than the one that created the `ExcelApp`. |
| [`WorkbookNotFoundError`](reference/index.md#ezxl.WorkbookNotFoundError)         | A workbook cannot be found by name in the current Excel session.                       |
| [`SheetNotFoundError`](reference/index.md#ezxl.SheetNotFoundError)               | A worksheet cannot be found by name in a workbook.                                     |
| [`COMOperationError`](reference/index.md#ezxl.COMOperationError)                 | An unclassified COM error that does not map to a more specific subclass.               |
| [`GUIOperationError`](reference/index.md#ezxl.GUIOperationError)                 | A COM error occurring within a GUI surface (ribbon, CommandBars, dialog).              |
| [`FormatterError`](reference/index.md#ezxl.FormatterError)                       | An openpyxl-based formatting operation failed.                                         |
