# ///////////////////////////////////////////////////////////////
# _connect - Shared Excel window connection helper (pywinauto)
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
Shared Excel window connection helper for pywinauto backends.

Provides :func:`_get_excel_window`, a single entry point used by all four
pywinauto backend classes to obtain a ``pywinauto`` ``WindowSpecification``
for the Excel main window.

This module intentionally has **no dependency** on ``ExcelApp`` or any
COM object.  It operates purely at the OS UI level via ``pywinauto``.

Note:
    ``pywinauto`` is an optional dependency.  Install it with::

        pip install pywinauto
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
import logging

# Local imports
from ...exceptions import GUIOperationError
from ._imports import Application, WindowSpecification

# ///////////////////////////////////////////////////////////////
# CONSTANTS
# ///////////////////////////////////////////////////////////////

logger = logging.getLogger(__name__)

# Standard Excel main-window title pattern. Matches "Book1 - Microsoft Excel",
# "report.xlsx - Microsoft Excel", etc.
_EXCEL_TITLE_RE: str = r".*- Microsoft Excel$"

# ///////////////////////////////////////////////////////////////
# FUNCTIONS
# ///////////////////////////////////////////////////////////////


def _get_excel_window(hwnd: int | None = None) -> WindowSpecification:  # type: ignore[reportUnusedFunction]  -- imported by _ribbon, _menu, _dialog
    """Return a pywinauto ``WindowSpecification`` for the Excel main window.

    Connects to a running Excel instance using the ``uia`` backend, which
    supports the full UI Automation tree available in modern Excel (2013+).

    Connection strategy:

    - If *hwnd* is provided, connect directly via the window handle
      (fastest; no title disambiguation required).
    - If *hwnd* is ``None``, connect via a regular-expression title match
      ``".*- Microsoft Excel$"`` (matches the standard Excel caption format).

    Args:
        hwnd: Win32 window handle for the Excel main window.  Pass ``None``
            to auto-detect the first visible Excel instance.

    Returns:
        WindowSpecification: A ``pywinauto`` ``WindowSpecification``
        representing the Excel main window.

    Raises:
        GUIOperationError: If no Excel window can be found, or if the
            pywinauto connection attempt raises any exception.

    Warning:
        Passing ``hwnd=None`` is unsafe when multiple Excel instances are
        open simultaneously.  pywinauto will connect to whichever window
        matches ``".*- Microsoft Excel$"`` first, which may **not** be the
        workbook managed by the calling ``ExcelApp`` COM session.  This can
        silently interact with the wrong workbook.

        Always pass an explicit ``hwnd`` in production code.  Obtain it
        from the managing ``ExcelApp`` instance::

            hwnd = xl.hwnd
            win = _get_excel_window(hwnd=hwnd)

    Example:
        >>> win = _get_excel_window()
        >>> win.set_focus()

        >>> win = _get_excel_window(hwnd=131234)
    """
    logger.debug("_get_excel_window: hwnd=%r", hwnd)
    try:
        app = Application(backend="uia")
        if hwnd is not None:
            app.connect(handle=hwnd)
            window: WindowSpecification = app.window(handle=hwnd)
        else:
            app.connect(title_re=_EXCEL_TITLE_RE)
            window = app.window(title_re=_EXCEL_TITLE_RE)
        logger.debug("_get_excel_window: connected to Excel window %r", window)
        return window
    except GUIOperationError:
        raise
    except Exception as exc:
        raise GUIOperationError(
            f"Could not connect to an Excel window: {exc}. "
            "Ensure Excel is running and visible.",
            cause=exc,
        ) from exc
