# ///////////////////////////////////////////////////////////////
# _workbook - WorkbookProxy COM wrapper
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
WorkbookProxy — thin COM proxy over an Excel Workbook object.

Provides save, save-as (with format conversion), close, and sheet
navigation without exposing raw COM objects to callers. All COM calls
are wrapped via ``_com_utils.wrap_com_error``.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
from pathlib import Path
from typing import TYPE_CHECKING, Any

# Third-party imports
from ezplog.lib_mode import get_logger, get_printer

# Local imports
from ..exceptions import SheetNotFoundError, WorkbookNotFoundError
from ..utils._com_utils import assert_main_thread, wrap_com_error

if TYPE_CHECKING:
    from ._excel_app import ExcelApp
    from ._sheet import SheetProxy

# ///////////////////////////////////////////////////////////////
# CONSTANTS
# ///////////////////////////////////////////////////////////////

logger = get_logger(__name__)
printer = get_printer()

# Excel FileFormat constants — used when calling SaveAs with a specific format.
# These match the XlFileFormat enumeration values from the Excel object model.
_FORMAT_MAP: dict[str, int] = {
    ".xlsx": 51,  # xlOpenXMLWorkbook
    ".xlsm": 52,  # xlOpenXMLWorkbookMacroEnabled
    ".xlsb": 50,  # xlExcel12 (binary)
    ".xls": 56,  # xlExcel8
    ".csv": 6,  # xlCSV
    ".txt": 42,  # xlUnicodeText
    ".pdf": 57,  # xlTypePDF  (Export, not SaveAs — handled separately)
    ".ods": 60,  # xlOpenDocumentSpreadsheet
}

# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


class WorkbookProxy:
    """COM proxy for a single Excel Workbook.

    Instances are created by ``ExcelApp.open()`` or ``ExcelApp.workbook()``.
    Do not instantiate directly.

    All methods enforce COM thread safety by delegating to the parent
    ``ExcelApp``'s thread identity.

    Args:
        app: The ``ExcelApp`` that owns this session.
        name: The workbook name as shown in Excel's title bar
            (e.g. ``"budget.xlsx"``).

    Example:
        >>> with ExcelApp() as xl:
        ...     wb = xl.open("C:/data/report.xlsx")
        ...     print(wb.name)
        ...     wb.save()
        ...     wb.close()
    """

    # ///////////////////////////////////////////////////////////////
    # INIT
    # ///////////////////////////////////////////////////////////////

    def __init__(self, app: ExcelApp, name: str) -> None:
        self._app = app
        self._name = name

    # ///////////////////////////////////////////////////////////////
    # PRIVATE HELPERS
    # ///////////////////////////////////////////////////////////////

    def _check_thread(self) -> None:
        """Delegate thread assertion to the parent ExcelApp."""
        assert_main_thread(self._app._thread_id)

    def _get_wb(self) -> Any:
        """Resolve and return the underlying COM Workbook object.

        Returns:
            The win32com Workbook COM object.

        Raises:
            WorkbookNotFoundError: If no open workbook matches ``self._name``.
        """
        xl = self._app._get_app()
        try:
            return xl.Workbooks(self._name)
        except Exception as exc:
            raise WorkbookNotFoundError(
                f"Workbook '{self._name}' is no longer open.", cause=exc
            ) from exc

    # ///////////////////////////////////////////////////////////////
    # PROPERTIES
    # ///////////////////////////////////////////////////////////////

    @property
    def name(self) -> str:
        """The workbook filename as shown in Excel's title bar.

        Returns:
            str: Workbook name (e.g. ``"report.xlsx"``).
        """
        return self._name

    @property
    @wrap_com_error
    def sheets(self) -> list[str]:
        """List of all worksheet names in this workbook.

        Returns:
            list[str]: Sheet names in tab order.

        Raises:
            WorkbookNotFoundError: If the workbook is no longer open.
        """
        self._check_thread()
        wb = self._get_wb()
        return [wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)]

    # ///////////////////////////////////////////////////////////////
    # PUBLIC METHODS
    # ///////////////////////////////////////////////////////////////

    @wrap_com_error
    def sheet(self, name: str) -> SheetProxy:
        """Return a proxy for a named worksheet.

        Args:
            name: The worksheet name (case-sensitive, as shown on the tab).

        Returns:
            SheetProxy: A proxy bound to the named worksheet.

        Raises:
            SheetNotFoundError: If no sheet with that name exists.
            WorkbookNotFoundError: If the workbook is no longer open.

        Example:
            >>> ws = wb.sheet("Summary")
        """
        from ._sheet import SheetProxy  # local import avoids circular dep

        self._check_thread()
        wb = self._get_wb()

        # Validate existence before constructing the proxy.
        for i in range(1, wb.Sheets.Count + 1):
            if wb.Sheets(i).Name == name:
                logger.debug("Resolved sheet '%s' in '%s'.", name, self._name)
                return SheetProxy(self, name)

        available = [wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)]
        raise SheetNotFoundError(
            f"No sheet named '{name}' in '{self._name}'. Available sheets: {available}"
        )

    @wrap_com_error
    def save(self) -> None:
        """Save the workbook in place (equivalent to Ctrl+S).

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            WorkbookNotFoundError: If the workbook is no longer open.
            COMOperationError: If the save fails.

        Example:
            >>> wb.save()
        """
        self._check_thread()
        logger.debug("Saving workbook: %s", self._name)
        self._get_wb().Save()
        printer.success(f"Workbook saved: {self._name}")

    @wrap_com_error
    def save_as(self, path: str | Path, fmt: str | None = None) -> None:
        """Save the workbook to a new path, optionally changing its format.

        Uses COM ``Workbook.SaveAs`` which keeps Excel open. Suitable for
        format conversion (e.g. xlsx → csv) via an active Excel session.

        Args:
            path: Destination file path. The extension determines the format
                when ``fmt`` is ``None``.
            fmt: Explicit format override. Must be a key from the internal
                format map (e.g. ``".csv"``, ``".xlsx"``). If omitted the
                extension of ``path`` is used.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            WorkbookNotFoundError: If the workbook is no longer open.
            COMOperationError: If SaveAs fails.
            ValueError: If the file extension cannot be mapped to a COM format.

        Example:
            >>> wb.save_as("C:/output/report.csv")
            >>> wb.save_as("C:/output/report_backup.xlsx", fmt=".xlsx")
        """
        self._check_thread()
        dest = Path(path).resolve()
        extension = fmt if fmt is not None else dest.suffix.lower()

        file_format: int | None = _FORMAT_MAP.get(extension)
        if file_format is None:
            raise ValueError(
                f"Unsupported file format '{extension}'. "
                f"Supported: {list(_FORMAT_MAP.keys())}"
            )

        logger.debug(
            "SaveAs workbook '%s' → '%s' (format=%d).", self._name, dest, file_format
        )
        wb = self._get_wb()

        # PDF export uses a different COM method.
        if extension == ".pdf":
            wb.ExportAsFixedFormat(0, str(dest))
        else:
            wb.SaveAs(str(dest), FileFormat=file_format)

        printer.success(f"Workbook saved as: {dest}")

    @wrap_com_error
    def close(self, save: bool = False) -> None:
        """Close the workbook.

        Args:
            save: If ``True``, save changes before closing. Defaults to
                ``False`` (discard unsaved changes).

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            WorkbookNotFoundError: If the workbook is no longer open.
            COMOperationError: If the close operation fails.

        Example:
            >>> wb.close(save=True)
        """
        self._check_thread()
        logger.debug("Closing workbook '%s' (save=%s).", self._name, save)
        self._get_wb().Close(SaveChanges=save)
        printer.system(f"Workbook closed: {self._name}")
