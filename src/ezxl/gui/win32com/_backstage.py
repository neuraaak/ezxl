# ///////////////////////////////////////////////////////////////
# _backstage - COMBackstageBackend: Backstage file ops via COM
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
COMBackstageBackend — Excel Backstage file operations via COM.

Implements :class:`~ezxl.gui._protocols.AbstractBackstageFileOps` using the
Excel COM object model.  COM calls are direct and focus-independent — no
keyboard sequences or window handles are required.

Advantages over the pywinauto backend
--------------------------------------
- No focus requirement: COM calls work regardless of which window is active.
- Locale-independent: COM APIs are language-neutral.
- No external dependency: uses only ``win32com.client`` (already required).

Limitations
-----------
- ``open_options`` is provided as a convenience method but is **not** part
  of the :class:`~ezxl.gui._protocols.AbstractBackstageFileOps` contract.
  It may fail in restricted environments (e.g. during a macro run or when
  the ribbon is disabled).  The preferred implementation for
  ``open_options`` is :class:`~ezxl.gui.pywinauto.PywinautoBackstageBackend`
  via :class:`~ezxl.gui._protocols.AbstractBackstageNavigator`.
- ``save_as`` with ``path=None`` shows the built-in Excel Save As dialog
  (``xlDialogSaveAs``).  The dialog blocks until the user dismisses it.
- ``close_workbook`` raises :exc:`~ezxl.exceptions.WorkbookNotFoundError`
  if no workbook is currently open.

All COM calls are guarded by :func:`~ezxl.utils._com_utils.wrap_com_error`
and :func:`~ezxl.utils._com_utils.assert_main_thread`.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
from typing import Any

# Third-party imports
from ezplog.lib_mode import get_logger

# Local imports
from ...exceptions import GUIOperationError, WorkbookNotFoundError
from ...utils._com_utils import assert_main_thread, wrap_com_error
from .._protocols import AbstractBackstageFileOps, ExcelAppLike

# ///////////////////////////////////////////////////////////////
# CONSTANTS
# ///////////////////////////////////////////////////////////////

logger = get_logger(__name__)

# xlDialogOpen constant — built-in Excel dialog identifier.
_XL_DIALOG_OPEN: int = 1

# xlDialogSaveAs constant — built-in Excel dialog identifier.
_XL_DIALOG_SAVE_AS: int = 5

# Maps lowercase file extensions to Excel XlFileFormat numeric constants.
# When SaveAs is called with a path, the FileFormat is inferred from the
# extension so Excel does not silently keep the current workbook format.
# References: https://learn.microsoft.com/en-us/office/vba/api/excel.xlfileformat
_EXT_TO_FILE_FORMAT: dict[str, int] = {
    ".xlsx": 51,  # xlOpenXMLWorkbook
    ".xlsm": 52,  # xlOpenXMLWorkbookMacroEnabled
    ".xlsb": 50,  # xlExcel12 (binary)
    ".xls": 56,  # xlExcel8 (Excel 97-2003)
    ".csv": 6,  # xlCSV
    ".txt": -4158,  # xlTextWindows
    ".xml": 46,  # xlXMLSpreadsheet
}

# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


class COMBackstageBackend(AbstractBackstageFileOps):
    """Excel Backstage file operations via the COM object model.

    Implements :class:`~ezxl.gui._protocols.AbstractBackstageFileOps` using
    Excel's COM API.  All operations are focus-independent and
    locale-independent.

    This is the default backstage backend for :class:`~ezxl.gui.GUIProxy`.
    For UIA-based Backstage navigation (Options panel, visual panel
    opening), inject a
    :class:`~ezxl.gui.pywinauto.PywinautoBackstageBackend` as the
    ``backstage_nav`` argument::

        gui = GUIProxy(
            xl,
            backstage=COMBackstageBackend(xl),
            backstage_nav=PywinautoBackstageBackend(hwnd=xl.hwnd, locale="fr"),
        )

    Args:
        app: The active ``ExcelApp`` instance that owns this backend.

    Example:
        >>> backend = COMBackstageBackend(xl)
        >>> backend.save()
        >>> backend.save_as(path="C:\\\\Reports\\\\output.xlsx")
        >>> backend.open_file()
        >>> backend.close_workbook()
    """

    # ///////////////////////////////////////////////////////////////
    # INIT
    # ///////////////////////////////////////////////////////////////

    def __init__(self, app: ExcelAppLike) -> None:
        self._app = app

    # ///////////////////////////////////////////////////////////////
    # PRIVATE HELPERS
    # ///////////////////////////////////////////////////////////////

    def _check_thread(self) -> None:
        """Assert the call originates from the COM apartment thread."""
        assert_main_thread(self._app._thread_id)

    def _get_app(self) -> Any:
        """Return the underlying COM ``Application`` object."""
        return self._app._get_app()

    def _active_workbook(self) -> Any:
        """Return ``ActiveWorkbook``, raising if no workbook is open.

        Raises:
            WorkbookNotFoundError: If no workbook is currently active.
        """
        wb: Any = self._get_app().ActiveWorkbook
        if wb is None:
            raise WorkbookNotFoundError(
                "No active workbook — cannot perform Backstage action."
            )
        return wb

    # ///////////////////////////////////////////////////////////////
    # AbstractBackstageFileOps implementation
    # ///////////////////////////////////////////////////////////////

    @wrap_com_error
    def save(self) -> None:
        """Save the active workbook via ``ActiveWorkbook.Save()``.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            WorkbookNotFoundError: If no workbook is currently open.
            COMOperationError: If the COM call fails.

        Example:
            >>> backend.save()
        """
        self._check_thread()
        logger.debug("COMBackstageBackend.save")
        self._active_workbook().Save()

    @wrap_com_error
    def save_as(self, path: str | None = None) -> None:
        """Save the active workbook under a new path, or show the Save As dialog.

        If *path* is provided, calls ``ActiveWorkbook.SaveAs(Filename=path)``
        directly — no dialog is shown.  If *path* is ``None``, opens the
        built-in Excel Save As dialog (``xlDialogSaveAs``), which blocks until
        the user dismisses it.

        Args:
            path: Absolute path for the new file.  If ``None``, the built-in
                dialog is displayed for manual path selection.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            WorkbookNotFoundError: If no workbook is currently open.
            COMOperationError: If the COM call fails.

        Example:
            >>> backend.save_as()                              # opens dialog
            >>> backend.save_as(path="C:\\\\output.xlsx")     # direct save
        """
        import pathlib

        self._check_thread()
        logger.debug("COMBackstageBackend.save_as: path=%r", path)
        if path is not None:
            app = self._get_app()
            fmt = _EXT_TO_FILE_FORMAT.get(pathlib.Path(path).suffix.lower())
            app.DisplayAlerts = False
            try:
                if fmt is not None:
                    self._active_workbook().SaveAs(Filename=path, FileFormat=fmt)
                else:
                    self._active_workbook().SaveAs(Filename=path)
            finally:
                app.DisplayAlerts = True
        else:
            self._get_app().Dialogs(_XL_DIALOG_SAVE_AS).Show()

    @wrap_com_error
    def open_file(self) -> None:
        """Show the built-in Excel Open dialog (``xlDialogOpen``).

        The dialog blocks until the user selects a file or cancels.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            COMOperationError: If the COM call fails.

        Example:
            >>> backend.open_file()
        """
        self._check_thread()
        logger.debug("COMBackstageBackend.open_file")
        self._get_app().Dialogs(_XL_DIALOG_OPEN).Show()

    @wrap_com_error
    def close_workbook(self) -> None:
        """Close the active workbook without saving (``SaveChanges=False``).

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            WorkbookNotFoundError: If no workbook is currently open.
            COMOperationError: If the COM call fails.

        Example:
            >>> backend.close_workbook()
        """
        self._check_thread()
        logger.debug("COMBackstageBackend.close_workbook")
        self._active_workbook().Close(SaveChanges=False)

    # ///////////////////////////////////////////////////////////////
    # EXTRA — not part of AbstractBackstageFileOps
    # ///////////////////////////////////////////////////////////////

    @wrap_com_error
    def open_options(self) -> None:
        """Open the Excel Options dialog via ``CommandBars.ExecuteMso``.

        This method is **not** part of :class:`AbstractBackstageFileOps`.
        It is provided as a convenience fallback when
        :class:`~ezxl.gui.pywinauto.PywinautoBackstageBackend` (the
        preferred implementation via :class:`AbstractBackstageNavigator`)
        is unavailable.

        Uses the ``"ApplicationOptionsDialog"`` MSO identifier.  This may
        fail in restricted environments (e.g. when a macro is running or the
        ribbon is disabled).

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the Options dialog cannot be opened via COM.

        Example:
            >>> backend.open_options()
        """
        self._check_thread()
        logger.debug("COMBackstageBackend.open_options")
        try:
            self._get_app().CommandBars.ExecuteMso("ApplicationOptionsDialog")
        except Exception as exc:
            raise GUIOperationError(
                "Could not open Excel Options dialog via COM "
                "(CommandBars.ExecuteMso('ApplicationOptionsDialog') failed). "
                "Consider using PywinautoBackstageBackend.open_options() via "
                "GUIProxy.backstage_nav as the preferred implementation. "
                f"Original error: {exc}",
                cause=exc,
            ) from exc
