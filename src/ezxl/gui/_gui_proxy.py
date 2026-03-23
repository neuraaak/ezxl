# ///////////////////////////////////////////////////////////////
# _gui_proxy - GUIProxy: unified GUI interaction facade
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
GUIProxy — lightweight facade that bundles all GUI interaction surfaces.

Exposed via ``ExcelApp.gui``. Provides access to four GUI automation
surfaces without requiring callers to instantiate individual proxies:

- ``gui.ribbon``        — ``AbstractRibbonBackend`` for MSO ribbon commands.
- ``gui.menu``          — ``AbstractMenuBackend`` for legacy CommandBar traversal.
- ``gui.dialog``        — ``AbstractDialogBackend`` for file pickers and alerts.
- ``gui.send_keys(…)``  — Direct ``Application.SendKeys`` pass-through via
  ``AbstractKeysBackend``.

Backend instances are created once during ``__init__`` and stored on the
proxy.  Callers may inject alternative backends (e.g. a pywinauto-based
implementation) via the optional keyword arguments.  Passing ``None`` for
any backend selects the default COM implementation.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
from typing import TYPE_CHECKING

# Third-party imports
from ezplog.lib_mode import get_logger

# Local imports
from ._protocols import (
    AbstractDialogBackend,
    AbstractKeysBackend,
    AbstractMenuBackend,
    AbstractRibbonBackend,
)
from .win32com._dialog import DialogProxy
from .win32com._keys import _COMKeysBackend
from .win32com._menu import MenuProxy
from .win32com._ribbon import RibbonProxy

if TYPE_CHECKING:
    from ..core._excel_app import ExcelApp

# ///////////////////////////////////////////////////////////////
# CONSTANTS
# ///////////////////////////////////////////////////////////////

logger = get_logger(__name__)

# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


class GUIProxy:
    """Unified facade for GUI-level Excel interaction.

    Instantiated by ``ExcelApp.gui`` and bundles all four automation
    surfaces (ribbon, menu, dialog, keys) under a single object.

    Backend injection
    -----------------
    Each surface can be replaced at construction time by passing an
    alternative implementation that satisfies the corresponding abstract
    protocol.  This is the primary extension point for non-COM backends
    such as pywinauto.  Passing ``None`` (the default) selects the
    standard COM implementation.

    Args:
        app: The active ``ExcelApp`` instance that owns this proxy.
        ribbon: Optional ribbon backend. Defaults to
            :class:`~ezxl.gui.RibbonProxy` when ``None``.
        menu: Optional menu backend. Defaults to
            :class:`~ezxl.gui.MenuProxy` when ``None``.
        dialog: Optional dialog backend. Defaults to
            :class:`~ezxl.gui.DialogProxy` when ``None``.
        keys: Optional keys backend. Defaults to
            :class:`~ezxl.gui._keys._COMKeysBackend` when ``None``.

    Security note:
        When injecting pywinauto backends, always pass ``hwnd=app.hwnd``
        to bind the backend to the exact Excel window managed by this
        session.  Omitting ``hwnd`` causes pywinauto to attach to the
        first Excel window it finds, which may not be the correct one
        when multiple Excel instances are running.

        Example::

            gui = GUIProxy(xl, ribbon=PywinautoRibbonBackend(hwnd=xl.hwnd))

    Example:
        >>> with ExcelApp(mode="attach") as xl:
        ...     xl.gui.ribbon.execute("FileSave")
        ...     path = xl.gui.dialog.get_file_open()
        ...     xl.gui.send_keys("^{HOME}")
    """

    # ///////////////////////////////////////////////////////////////
    # INIT
    # ///////////////////////////////////////////////////////////////

    def __init__(
        self,
        app: ExcelApp,
        ribbon: AbstractRibbonBackend | None = None,
        menu: AbstractMenuBackend | None = None,
        dialog: AbstractDialogBackend | None = None,
        keys: AbstractKeysBackend | None = None,
    ) -> None:
        self._app = app
        # Construct default COM implementations when no override is provided.
        self._ribbon: AbstractRibbonBackend = (
            ribbon if ribbon is not None else RibbonProxy(app)
        )
        self._menu: AbstractMenuBackend = menu if menu is not None else MenuProxy(app)
        self._dialog: AbstractDialogBackend = (
            dialog if dialog is not None else DialogProxy(app)
        )
        self._keys: AbstractKeysBackend = (
            keys if keys is not None else _COMKeysBackend(app)
        )

    # ///////////////////////////////////////////////////////////////
    # PROPERTIES
    # ///////////////////////////////////////////////////////////////

    @property
    def ribbon(self) -> AbstractRibbonBackend:
        """Return the ribbon backend for MSO ribbon command interaction.

        Returns:
            AbstractRibbonBackend: The configured ribbon backend
                (default: :class:`~ezxl.gui.RibbonProxy`).

        Example:
            >>> xl.gui.ribbon.execute("FileSave")
            >>> xl.gui.ribbon.is_enabled("Copy")
            True
        """
        return self._ribbon

    @property
    def menu(self) -> AbstractMenuBackend:
        """Return the menu backend for legacy CommandBar interaction.

        Returns:
            AbstractMenuBackend: The configured menu backend
                (default: :class:`~ezxl.gui.MenuProxy`).

        Example:
            >>> xl.gui.menu.list_bars()
            ['Standard', 'Formatting', ...]
        """
        return self._menu

    @property
    def dialog(self) -> AbstractDialogBackend:
        """Return the dialog backend for file picker and alert dialogs.

        Returns:
            AbstractDialogBackend: The configured dialog backend
                (default: :class:`~ezxl.gui.DialogProxy`).

        Example:
            >>> path = xl.gui.dialog.get_file_open(title="Select report")
        """
        return self._dialog

    # ///////////////////////////////////////////////////////////////
    # PUBLIC METHODS
    # ///////////////////////////////////////////////////////////////

    def send_keys(self, keys: str, wait: bool = True) -> None:
        """Send a keystroke sequence to the Excel Application window.

        Delegates to the configured :class:`~ezxl.gui._protocols.AbstractKeysBackend`.
        The ``keys`` string must use standard VBA SendKeys notation
        (e.g. ``"{ENTER}"``, ``"^s"`` for Ctrl+S).

        Args:
            keys: Keystroke string in VBA SendKeys notation.
            wait: If ``True``, block until Excel processes the keystrokes
                before returning. Defaults to ``True``.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            COMOperationError: If the SendKeys call fails.

        Example:
            >>> xl.gui.send_keys("^s")           # Ctrl+S
            >>> xl.gui.send_keys("{ESCAPE}")
        """
        logger.debug("GUIProxy.send_keys: keys=%r, wait=%r", keys, wait)
        self._keys.send_keys(keys, wait)
