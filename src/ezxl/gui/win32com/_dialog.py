# ///////////////////////////////////////////////////////////////
# _dialog - DialogProxy: file and message dialogs via COM
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
DialogProxy — native Excel dialog helpers via ``Application`` COM methods.

Exposes three dialog types:

- ``get_file_open`` — ``Application.GetOpenFilename``; returns the
  user-selected path or ``None`` on cancel.
- ``get_file_save`` — ``Application.GetSaveAsFilename``; returns the
  user-selected path or ``None`` on cancel.
- ``alert`` — a modal message box. Primary implementation uses
  ``ctypes.windll.user32.MessageBoxW`` to avoid requiring a running
  VBA environment. Falls back gracefully if ctypes is unavailable.

All COM calls are guarded by ``wrap_com_error`` and
``assert_main_thread``.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
import contextlib
import ctypes

# Third-party imports
from ezplog.lib_mode import get_logger

# Local imports
from ...exceptions import GUIOperationError
from ...utils._com_utils import assert_main_thread, wrap_com_error
from .._protocols import AbstractDialogBackend, ExcelAppLike

# ///////////////////////////////////////////////////////////////
# CONSTANTS
# ///////////////////////////////////////////////////////////////

logger = get_logger(__name__)

# Win32 MessageBox button/icon flags.
_MB_OK: int = 0x00000000
_MB_ICONINFORMATION: int = 0x00000040

# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


class DialogProxy(AbstractDialogBackend):
    """File and message dialog helpers backed by Excel COM and Win32.

    Provides a clean Python interface for the three most common GUI
    dialogs needed during Excel automation: open-file picker, save-file
    picker, and a simple information alert.

    Args:
        app: The active ``ExcelApp`` instance that owns this proxy.

    Example:
        >>> proxy = DialogProxy(xl)
        >>> path = proxy.get_file_open(title="Select a report")
        >>> if path:
        ...     wb = xl.open(path)
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

    # ///////////////////////////////////////////////////////////////
    # PUBLIC METHODS
    # ///////////////////////////////////////////////////////////////

    @wrap_com_error
    def get_file_open(
        self,
        title: str = "Open",
        initial_dir: str | None = None,
        filter: str = "Excel Files (*.xls*), *.xls*",
    ) -> str | None:
        """Show Excel's built-in Open file picker dialog.

        Calls ``Application.GetOpenFilename``. The dialog is modal;
        this method blocks until the user confirms or cancels.

        Args:
            title: Dialog title bar text. Defaults to ``"Open"``.
            initial_dir: Directory to open the dialog in. If ``None``,
                Excel uses its current working directory.
            filter: File-type filter string in Excel's two-part format:
                ``"<description>, <wildcard>"``. Defaults to Excel files.

        Returns:
            str | None: Absolute path chosen by the user, or ``None``
                if the dialog was cancelled.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the COM call fails.

        Example:
            >>> path = dialog.get_file_open(title="Pick a workbook")
            >>> if path is not None:
            ...     xl.open(path)
        """
        self._check_thread()
        logger.debug(
            "DialogProxy.get_file_open: title=%r, initial_dir=%r, filter=%r",
            title,
            initial_dir,
            filter,
        )
        xl = self._app._get_app()

        # Optionally change the initial directory for the duration of the
        # dialog, then restore it to avoid side-effects.
        original_dir: str | None = None
        if initial_dir is not None:
            try:
                original_dir = xl.DefaultFilePath
                xl.DefaultFilePath = initial_dir
            except Exception:
                # Non-critical; continue without setting the directory.
                original_dir = None

        try:
            result = xl.GetOpenFilename(
                FileFilter=filter,
                Title=title,
            )
        except Exception as exc:
            raise GUIOperationError(
                f"GetOpenFilename failed: {exc}", cause=exc
            ) from exc
        finally:
            if original_dir is not None:
                with contextlib.suppress(Exception):
                    xl.DefaultFilePath = original_dir

        # Excel returns False (boolean) when the user cancels.
        if result is False or result == "False":
            logger.debug("DialogProxy.get_file_open: cancelled by user.")
            return None

        logger.debug("DialogProxy.get_file_open: selected=%r", result)
        return str(result)

    @wrap_com_error
    def get_file_save(
        self,
        title: str = "Save As",
        initial_dir: str | None = None,
        filter: str = "Excel Files (*.xlsx), *.xlsx",
    ) -> str | None:
        """Show Excel's built-in Save As file picker dialog.

        Calls ``Application.GetSaveAsFilename``. The dialog is modal;
        this method blocks until the user confirms or cancels.

        Args:
            title: Dialog title bar text. Defaults to ``"Save As"``.
            initial_dir: Directory to open the dialog in. If ``None``,
                Excel uses its current working directory.
            filter: File-type filter string in Excel's two-part format:
                ``"<description>, <wildcard>"``. Defaults to ``.xlsx``.

        Returns:
            str | None: Absolute path chosen by the user, or ``None``
                if the dialog was cancelled.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the COM call fails.

        Example:
            >>> path = dialog.get_file_save(title="Save report as")
            >>> if path is not None:
            ...     wb.save_as(path)
        """
        self._check_thread()
        logger.debug(
            "DialogProxy.get_file_save: title=%r, initial_dir=%r, filter=%r",
            title,
            initial_dir,
            filter,
        )
        xl = self._app._get_app()

        original_dir: str | None = None
        if initial_dir is not None:
            try:
                original_dir = xl.DefaultFilePath
                xl.DefaultFilePath = initial_dir
            except Exception:
                original_dir = None

        try:
            result = xl.GetSaveAsFilename(
                FileFilter=filter,
                Title=title,
            )
        except Exception as exc:
            raise GUIOperationError(
                f"GetSaveAsFilename failed: {exc}", cause=exc
            ) from exc
        finally:
            if original_dir is not None:
                with contextlib.suppress(Exception):
                    xl.DefaultFilePath = original_dir

        if result is False or result == "False":
            logger.debug("DialogProxy.get_file_save: cancelled by user.")
            return None

        logger.debug("DialogProxy.get_file_save: selected=%r", result)
        return str(result)

    def alert(self, message: str, title: str = "EzXl") -> None:
        """Display a modal information message box.

        Uses ``ctypes.windll.user32.MessageBoxW`` (Win32 API) directly.
        This approach avoids any dependency on a running VBA environment
        and does not require Excel to have a document open.

        The dialog shows a single OK button with an information icon.
        This method blocks until the user dismisses the dialog.

        Args:
            message: The body text displayed in the message box.
            title: Caption for the message box title bar.
                Defaults to ``"EzXl"``.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the Win32 MessageBoxW call fails.

        Example:
            >>> dialog.alert("Export complete.", title="Success")
        """
        self._check_thread()
        logger.debug("DialogProxy.alert: title=%r, message=%r", title, message)
        try:
            # HWND=0 → desktop owner (no parent window).
            ctypes.windll.user32.MessageBoxW(
                0,
                str(message),
                str(title),
                _MB_OK | _MB_ICONINFORMATION,
            )
        except Exception as exc:
            raise GUIOperationError(f"MessageBoxW failed: {exc}", cause=exc) from exc
