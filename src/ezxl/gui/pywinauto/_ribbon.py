# ///////////////////////////////////////////////////////////////
# _ribbon - PywinautoRibbonBackend: MSO ribbon via UI Automation
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
PywinautoRibbonBackend — MSO ribbon command execution via ``pywinauto``.

Implements :class:`~ezxl.gui._protocols.AbstractRibbonBackend` using
``pywinauto`` UI Automation rather than ``Application.CommandBars``.  This
backend finds the running Excel window, locates the ribbon control that
corresponds to the requested MSO identifier, and clicks it.

Limitations
-----------
- Only a curated set of common MSO identifiers is supported.  Pass an
  unmapped *mso_id* and :exc:`~ezxl.exceptions.GUIOperationError` is raised.
- :meth:`is_enabled`, :meth:`is_pressed`, and :meth:`is_visible` have no
  reliable pywinauto equivalent and raise :exc:`NotImplementedError`.
- This backend has **no COM dependency** and **no thread constraint**.  It
  operates purely at the OS UI level.

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
from typing import Any

# Local imports
from ...exceptions import GUIOperationError
from .._protocols import AbstractRibbonBackend
from ._connect import get_excel_window

# ///////////////////////////////////////////////////////////////
# OPTIONAL DEPENDENCY GUARD
# ///////////////////////////////////////////////////////////////

try:
    from pywinauto.keyboard import (
        send_keys as _pw_send_keys,  # type: ignore[import-untyped]
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

# Mapping from MSO identifiers to pywinauto keyboard shortcut strings.
# Keyboard shortcuts are locale-independent — they work regardless of the
# Excel UI language, unlike AutomationId or button-title searches which
# return localised strings (e.g. "Enregistrer" instead of "Save" in French).
# Keys are MSO identifiers; values are pywinauto send_keys notation.
_MSO_TO_KEYS: dict[str, str] = {
    "FileSave": "^s",
    "Copy": "^c",
    "Paste": "^v",
    "Bold": "^b",
    "Italic": "^i",
    "Underline": "^u",
    "Undo": "^z",
    "Redo": "^y",
}

# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


class PywinautoRibbonBackend(AbstractRibbonBackend):
    """Ribbon command execution via ``pywinauto`` UI Automation clicks.

    Connects to the running Excel window and clicks the ribbon button
    that corresponds to the requested MSO identifier.  Only identifiers
    listed in the internal ``_MSO_TO_KEYS`` mapping are supported.

    This backend is a standalone alternative to the COM-based
    :class:`~ezxl.gui.RibbonProxy`.  It does **not** require an
    :class:`~ezxl.core.ExcelApp` instance and carries no COM STA
    thread constraint.

    Args:
        hwnd: Win32 window handle for the Excel main window.  If ``None``,
            the backend auto-detects the first visible Excel instance whose
            title matches ``".*- Microsoft Excel$"``.

    Example:
        >>> from ezxl.gui.pywinauto import PywinautoRibbonBackend
        >>> ribbon = PywinautoRibbonBackend()
        >>> ribbon.execute("FileSave")   # clicks the Save ribbon button

        >>> # Inject into GUIProxy alongside default COM backends:
        >>> from ezxl import ExcelApp, GUIProxy
        >>> with ExcelApp(mode="attach") as xl:
        ...     gui = GUIProxy(xl, ribbon=PywinautoRibbonBackend())
        ...     gui.ribbon.execute("Bold")
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
        return get_excel_window(self._hwnd)

    # ///////////////////////////////////////////////////////////////
    # PUBLIC METHODS
    # ///////////////////////////////////////////////////////////////

    def execute(self, mso_id: str) -> None:
        """Execute a ribbon command via its keyboard shortcut.

        Maps *mso_id* to a keyboard shortcut via :data:`_MSO_TO_KEYS`, sets
        focus on the Excel window, then sends the keys via
        ``pywinauto.keyboard.send_keys``.

        Using keyboard shortcuts rather than UIA button searches makes this
        backend **locale-independent** — it works identically on French, English,
        and any other Excel locale, unlike title or AutomationId lookups which
        return localised button names.

        Supported MSO identifiers:

        - ``"FileSave"`` → ``Ctrl+S``
        - ``"Copy"``     → ``Ctrl+C``
        - ``"Paste"``    → ``Ctrl+V``
        - ``"Bold"``     → ``Ctrl+B``
        - ``"Italic"``   → ``Ctrl+I``
        - ``"Underline"`` → ``Ctrl+U``
        - ``"Undo"``     → ``Ctrl+Z``
        - ``"Redo"``     → ``Ctrl+Y``

        Args:
            mso_id: MSO control identifier string
                (e.g. ``"FileSave"``, ``"Bold"``).

        Raises:
            GUIOperationError: If *mso_id* is not in the supported mapping,
                the Excel window cannot be found, or the key send fails.

        Example:
            >>> ribbon.execute("FileSave")
            >>> ribbon.execute("Bold")
        """
        keys = _MSO_TO_KEYS.get(mso_id)
        if keys is None:
            supported = ", ".join(sorted(_MSO_TO_KEYS))
            raise GUIOperationError(
                f"MSO ID {mso_id!r} is not mapped in PywinautoRibbonBackend. "
                f"Supported identifiers: {supported}."
            )
        logger.debug("PywinautoRibbonBackend.execute: mso_id=%r, keys=%r", mso_id, keys)
        try:
            window = self._get_window()
            window.set_focus()
            _pw_send_keys(keys)
        except GUIOperationError:
            raise
        except Exception as exc:
            raise GUIOperationError(
                f"Failed to execute ribbon command {mso_id!r} (keys={keys!r}): {exc}",
                cause=exc,
            ) from exc

    def is_enabled(self, mso_id: str) -> bool:  # noqa: ARG002
        """Not supported by this backend.

        pywinauto does not expose MSO command-state queries equivalent to
        ``CommandBars.GetEnabledMso``.  Use the COM-based
        :class:`~ezxl.gui.RibbonProxy` for state queries.

        Args:
            mso_id: MSO control identifier string.

        Raises:
            NotImplementedError: Always.
        """
        raise NotImplementedError(
            "PywinautoRibbonBackend does not support MSO state queries. "
            "Use RibbonProxy (COM backend) for is_enabled()."
        )

    def is_pressed(self, mso_id: str) -> bool:  # noqa: ARG002
        """Not supported by this backend.

        pywinauto does not expose MSO toggle-state queries equivalent to
        ``CommandBars.GetPressedMso``.  Use the COM-based
        :class:`~ezxl.gui.RibbonProxy` for state queries.

        Args:
            mso_id: MSO control identifier string.

        Raises:
            NotImplementedError: Always.
        """
        raise NotImplementedError(
            "PywinautoRibbonBackend does not support MSO state queries. "
            "Use RibbonProxy (COM backend) for is_pressed()."
        )

    def is_visible(self, mso_id: str) -> bool:  # noqa: ARG002
        """Not supported by this backend.

        pywinauto does not expose MSO visibility queries equivalent to
        ``CommandBars.GetVisibleMso``.  Use the COM-based
        :class:`~ezxl.gui.RibbonProxy` for state queries.

        Args:
            mso_id: MSO control identifier string.

        Raises:
            NotImplementedError: Always.
        """
        raise NotImplementedError(
            "PywinautoRibbonBackend does not support MSO state queries. "
            "Use RibbonProxy (COM backend) for is_visible()."
        )
