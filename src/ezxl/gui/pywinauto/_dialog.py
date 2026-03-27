# ///////////////////////////////////////////////////////////////
# _dialog - PywinautoDialogBackend: file dialogs via UI Automation
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
PywinautoDialogBackend — file picker dialogs via ``pywinauto`` and Win32.

Implements :class:`~ezxl.gui._protocols.AbstractDialogBackend` by triggering
the Windows common file dialogs through keyboard shortcuts and then
interacting with the resulting dialog windows via ``pywinauto``.

Dialog strategy
---------------
- **get_file_open**: sends ``Ctrl+O`` to the focused Excel window, waits for
  the "Open" Windows common dialog to appear, optionally sets the filename
  field to *initial_dir*, and returns the user-selected path or ``None`` if
  cancelled.
- **get_file_save**: sends ``Ctrl+F12`` (Save As shortcut), waits for the
  "Save As" dialog, then follows the same flow.
- **alert**: delegates directly to ``ctypes.windll.user32.MessageBoxW``
  (Win32 API) — the same approach used by the COM backend.  No ``pywinauto``
  involvement; this is a pure Win32 call.

Important limitations
---------------------
- File picker methods **have a side effect**: they send keystrokes to the
  Excel window, which briefly changes its state.  If Excel is in edit-cell
  mode, the shortcut may be intercepted differently.
- The ``filter`` parameter is accepted for API compatibility but cannot be
  applied to the Windows common dialog after it has been opened via keyboard.
  The Windows common dialog respects whatever file-type filter was last set
  inside Excel's file dialogs.
- Dialog timeout is 10 seconds.  If the dialog does not appear within this
  window, :exc:`~ezxl.exceptions.GUIOperationError` is raised.
- This backend has **no COM dependency** and **no thread constraint**.

Note:
    ``pywinauto`` is an optional dependency.  Install it with::

        pip install pywinauto
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
import ctypes
import logging
from typing import Any

# Local imports
from ...exceptions import GUIOperationError
from .._protocols import AbstractDialogBackend
from ._connect import _get_excel_window

# ///////////////////////////////////////////////////////////////
# OPTIONAL DEPENDENCY GUARD
# ///////////////////////////////////////////////////////////////

try:
    from pywinauto import timings as _pw_timings  # type: ignore[import-untyped]
    from pywinauto.application import (  # type: ignore[import-untyped]
        Application as _PWApplication,
    )
    from pywinauto.keyboard import (  # type: ignore[import-untyped]
        send_keys as _pw_send_keys,
    )
except ImportError as _pwn_import_error:
    raise ImportError(
        "pywinauto is required for the pywinauto GUI backends but is not installed. "
        "Install it with: pip install pywinauto"
    ) from _pwn_import_error

# ///////////////////////////////////////////////////////////////
# CONSTANTS
# ///////////////////////////////////////////////////////////////

logger = logging.getLogger(__name__)

# Win32 MessageBox flags — identical to the COM DialogProxy.
_MB_OK: int = 0x00000000
_MB_ICONINFORMATION: int = 0x00000040

# Maximum seconds to wait for a file dialog window to appear.
_DIALOG_TIMEOUT: float = 10.0

# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


class PywinautoDialogBackend(AbstractDialogBackend):
    """File picker and alert dialogs via ``pywinauto`` and Win32.

    Triggers Windows common file dialogs through Excel keyboard shortcuts,
    then interacts with the resulting dialog window to return the user's
    selection.  Alert dialogs are shown via the Win32 ``MessageBoxW`` API.

    This backend is a standalone alternative to the COM-based
    :class:`~ezxl.gui.DialogProxy`.  It does **not** require an
    :class:`~ezxl.core.ExcelApp` instance and carries no COM STA
    thread constraint.

    Args:
        hwnd: Win32 window handle for the Excel main window.  If ``None``,
            the backend auto-detects the first visible Excel instance whose
            title matches ``".*- Microsoft Excel$"``.

    Example:
        >>> from ezxl.gui.pywinauto import PywinautoDialogBackend
        >>> dialog = PywinautoDialogBackend()
        >>> path = dialog.get_file_open()
        >>> if path:
        ...     print(f"Selected: {path}")

        >>> dialog.alert("Export complete.", title="Done")

        >>> # Inject into GUIProxy:
        >>> from ezxl import ExcelApp, GUIProxy
        >>> with ExcelApp(mode="attach") as xl:
        ...     gui = GUIProxy(xl, dialog=PywinautoDialogBackend())
        ...     path = gui.dialog.get_file_save()
    """

    # ///////////////////////////////////////////////////////////////
    # INIT
    # ///////////////////////////////////////////////////////////////

    def __init__(self, hwnd: int | None = None) -> None:
        self._hwnd = hwnd

    # ///////////////////////////////////////////////////////////////
    # PRIVATE HELPERS
    # ///////////////////////////////////////////////////////////////

    def _get_window(self) -> Any:
        """Return the pywinauto ``WindowSpecification`` for Excel."""
        return _get_excel_window(self._hwnd)

    def _wait_for_dialog(self, title_re: str) -> Any:
        """Wait for a top-level dialog window matching *title_re* to appear.

        Polls for up to :data:`_DIALOG_TIMEOUT` seconds using
        ``pywinauto.timings.wait_until_passes``.

        Args:
            title_re: Regular expression matched against the dialog's title.

        Returns:
            Any: The pywinauto ``WindowSpecification`` for the dialog.

        Raises:
            GUIOperationError: If the dialog does not appear within the
                timeout window.
        """
        try:
            app: Any = _PWApplication(backend="uia")
            app.connect(title_re=title_re, timeout=_DIALOG_TIMEOUT)
            dialog: Any = app.window(title_re=title_re)
            return dialog
        except Exception as exc:
            raise GUIOperationError(
                f"Timed out waiting for dialog matching {title_re!r} "
                f"after {_DIALOG_TIMEOUT:.0f}s: {exc}",
                cause=exc,
            ) from exc

    def _interact_file_dialog(
        self,
        dialog: Any,
        initial_dir: str | None,
    ) -> str | None:
        """Interact with an open Windows common file dialog.

        If *initial_dir* is provided, types it into the filename field so
        the dialog opens in that directory.  Then waits for the user to
        confirm or cancel.

        Args:
            dialog: The pywinauto ``WindowSpecification`` for the dialog.
            initial_dir: Directory to navigate to before presenting the
                dialog to the user.  ``None`` leaves the current directory.

        Returns:
            str | None: The text in the filename field after the user
                clicks OK, or ``None`` if the dialog was cancelled.
        """
        try:
            if initial_dir is not None:
                # Type the directory into the filename edit field so the
                # dialog starts in the requested location.
                try:
                    filename_edit: Any = dialog.child_window(
                        auto_id="1148", control_type="Edit"
                    )
                    filename_edit.set_edit_text(initial_dir)
                    _pw_send_keys("{ENTER}")
                    # Wait briefly for the dialog to navigate.
                    import time

                    time.sleep(0.3)
                except Exception as exc:
                    logger.debug(
                        "PywinautoDialogBackend: could not set initial_dir %r: %s",
                        initial_dir,
                        exc,
                    )

            # Block until the dialog closes (user clicks OK or Cancel).
            _pw_timings.wait_until_passes(
                _DIALOG_TIMEOUT,
                0.2,
                lambda: not dialog.exists(),
            )

            # Attempt to read the filename field value before the dialog
            # disappears.  If the dialog was cancelled the field is empty.
            try:
                filename_edit = dialog.child_window(auto_id="1148", control_type="Edit")
                selected: str = filename_edit.get_value()
                if selected and selected.strip():
                    return selected.strip()
                return None
            except Exception:
                # Dialog closed without a selection (cancelled).
                return None

        except GUIOperationError:
            raise
        except Exception as exc:
            raise GUIOperationError(
                f"Error interacting with file dialog: {exc}", cause=exc
            ) from exc

    # ///////////////////////////////////////////////////////////////
    # PUBLIC METHODS
    # ///////////////////////////////////////////////////////////////

    def get_file_open(
        self,
        title: str = "Open",
        initial_dir: str | None = None,
        filter: str = "Excel Files (*.xls*), *.xls*",  # noqa: A002,ARG002
    ) -> str | None:
        """Show the Windows Open file picker dialog and return the selected path.

        Sends ``Ctrl+O`` to the Excel window to trigger the Open dialog,
        then waits for the Windows common file dialog to appear.  If
        *initial_dir* is provided, it is typed into the filename field to
        navigate to that directory.

        Side effects
        ------------
        This method transfers keyboard focus to the Excel window and sends
        ``Ctrl+O``.  Ensure Excel is not in a state that would intercept
        or redirect this shortcut (e.g., an active cell edit).

        Note:
            The *filter* parameter is accepted for API compatibility but
            cannot be applied to the Windows common dialog after it opens
            via keyboard shortcut.

        Args:
            title: Dialog title bar text.  Accepted for API compatibility;
                the Windows common dialog title is controlled by Excel.
            initial_dir: Directory to navigate to before presenting the
                dialog.  ``None`` leaves Excel's current directory.
            filter: File-type filter string.  Accepted for API compatibility
                but **not applied** by this backend.

        Returns:
            str | None: Absolute path chosen by the user, or ``None``
                if the dialog was cancelled.

        Raises:
            GUIOperationError: If the Excel window cannot be found, the
                dialog does not appear within the timeout, or interaction
                fails.

        Example:
            >>> path = dialog.get_file_open(initial_dir="C:\\\\Reports")
            >>> if path:
            ...     print(f"User selected: {path}")
        """
        logger.debug(
            "PywinautoDialogBackend.get_file_open: title=%r, initial_dir=%r",
            title,
            initial_dir,
        )
        try:
            window = self._get_window()
            window.set_focus()
            # Ctrl+O triggers the Open dialog in all modern Excel versions.
            _pw_send_keys("^o")
            open_dialog = self._wait_for_dialog(r"Open")
            return self._interact_file_dialog(open_dialog, initial_dir)
        except GUIOperationError:
            raise
        except Exception as exc:
            raise GUIOperationError(f"get_file_open failed: {exc}", cause=exc) from exc

    def get_file_save(
        self,
        title: str = "Save As",
        initial_dir: str | None = None,
        filter: str = "Excel Files (*.xlsx), *.xlsx",  # noqa: A002,ARG002
    ) -> str | None:
        """Show the Windows Save As file picker dialog and return the selected path.

        Sends ``Ctrl+F12`` (universal Save As shortcut) to the Excel window
        to trigger the Save As dialog, then waits for the Windows common
        file dialog to appear.

        Side effects
        ------------
        This method transfers keyboard focus to the Excel window and sends
        ``Ctrl+F12``.  Ensure Excel is not in a state that would intercept
        or redirect this shortcut.

        Note:
            The *filter* parameter is accepted for API compatibility but
            cannot be applied to the Windows common dialog after it opens
            via keyboard shortcut.

        Args:
            title: Dialog title bar text.  Accepted for API compatibility.
            initial_dir: Directory to navigate to before presenting the
                dialog.  ``None`` leaves Excel's current directory.
            filter: File-type filter string.  Accepted for API compatibility
                but **not applied** by this backend.

        Returns:
            str | None: Absolute path chosen by the user, or ``None``
                if the dialog was cancelled.

        Raises:
            GUIOperationError: If the Excel window cannot be found, the
                dialog does not appear within the timeout, or interaction
                fails.

        Example:
            >>> path = dialog.get_file_save(initial_dir="C:\\\\Reports")
            >>> if path:
            ...     print(f"User will save to: {path}")
        """
        logger.debug(
            "PywinautoDialogBackend.get_file_save: title=%r, initial_dir=%r",
            title,
            initial_dir,
        )
        try:
            window = self._get_window()
            window.set_focus()
            # Ctrl+F12 is the universal Save As shortcut (F12 alone also works).
            _pw_send_keys("^{F12}")
            save_dialog = self._wait_for_dialog(r"Save As")
            return self._interact_file_dialog(save_dialog, initial_dir)
        except GUIOperationError:
            raise
        except Exception as exc:
            raise GUIOperationError(f"get_file_save failed: {exc}", cause=exc) from exc

    def alert(self, message: str, title: str = "EzXl") -> None:
        """Display a modal information message box.

        Uses ``ctypes.windll.user32.MessageBoxW`` (Win32 API) directly —
        identical to the COM-based :class:`~ezxl.gui.DialogProxy`.  No
        ``pywinauto`` involvement; this is a pure Win32 call that works
        regardless of COM availability.

        The dialog shows a single OK button with an information icon.
        This method blocks until the user dismisses the dialog.

        Args:
            message: The body text displayed in the message box.
            title: Caption for the message box title bar.
                Defaults to ``"EzXl"``.

        Raises:
            GUIOperationError: If the Win32 ``MessageBoxW`` call fails.

        Example:
            >>> dialog.alert("Export complete.", title="Success")
        """
        logger.debug(
            "PywinautoDialogBackend.alert: title=%r, message=%r", title, message
        )
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
