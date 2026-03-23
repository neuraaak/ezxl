# Changelog

All notable changes to this project are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).
This project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

---

## [0.1.0] — 2026-03-21

### Added

- `ExcelApp` — COM session lifecycle manager with `dispatch` and `attach` modes, context manager support, and STA thread identity enforcement.
- `WorkbookProxy` — thin COM proxy for workbook operations: `open`, `save`, `save_as` (multi-format via `XlFileFormat`), `close`, `sheet`, and `sheets`.
- `SheetProxy`, `CellProxy`, `RangeProxy` — COM proxies for worksheet, single-cell, and range access. COM date normalisation to `datetime`; error cells mapped to `None`.
- `GUIProxy` — unified GUI facade with injectable backend slots for ribbon, menu, dialog, and keys surfaces.
- `RibbonProxy`, `MenuProxy`, `DialogProxy` — default COM-based GUI backends implementing the four abstract protocols.
- `AbstractRibbonBackend`, `AbstractMenuBackend`, `AbstractDialogBackend`, `AbstractKeysBackend` — ABC contracts for custom GUI backends.
- `PywinautoRibbonBackend`, `PywinautoMenuBackend`, `PywinautoDialogBackend`, `PywinautoKeysBackend` — optional pywinauto-based GUI backends for UI Automation scenarios.
- `ExcelFormatter` — fluent closed-file formatting API via openpyxl: `column_width`, `row_height`, `font`, `fill`, `border`, `align`, `save`.
- `read_excel`, `read_csv`, `xlsx_to_csv`, `csv_to_xlsx`, `read_sheet` — polars-backed file I/O and format conversion functions. No running Excel process required.
- Full exception hierarchy: `EzXlError`, `ExcelNotAvailableError`, `ExcelSessionLostError`, `ExcelThreadViolationError`, `WorkbookNotFoundError`, `SheetNotFoundError`, `COMOperationError`, `GUIOperationError`, `FormatterError`.
- `wrap_com_error` decorator — intercepts `pywintypes.com_error` at the COM boundary and re-raises as typed EzXl exceptions.
- `assert_main_thread` — proactive STA thread check raised before any COM dispatcher call.
- `wait_until_ready` — polls `Application.Ready` until Excel becomes idle or a timeout expires.

[Unreleased]: https://github.com/neuraaak/ezxl/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/neuraaak/ezxl/releases/tag/v0.1.0
