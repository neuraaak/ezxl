# ///////////////////////////////////////////////////////////////
# _gui_proxy - GUIProxy: unified GUI interaction facade
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
GUIProxy — lightweight facade that bundles all GUI interaction surfaces.

Exposed via ``ExcelApp.gui``. Provides access to six GUI automation
surfaces without requiring callers to instantiate individual proxies:

- ``gui.ribbon``         — ``AbstractRibbonBackend`` for MSO ribbon commands.
- ``gui.menu``           — ``AbstractMenuBackend`` for legacy CommandBar traversal.
- ``gui.dialog``         — ``AbstractDialogBackend`` for file pickers and alerts.
- ``gui.send_keys(…)``   — Direct ``Application.SendKeys`` pass-through via
  ``AbstractKeysBackend``.
- ``gui.backstage``      — ``AbstractBackstageFileOps`` for file operations
  (save, save_as, open_file, close_workbook).  Defaults to
  :class:`~ezxl.gui.win32com._backstage.COMBackstageBackend`.
- ``gui.backstage_nav``  — ``AbstractBackstageNavigator | None`` for UIA-driven
  Backstage navigation (open_options, open_save_as_panel).  Defaults to
  ``None`` — inject a
  :class:`~ezxl.gui.pywinauto._backstage.PywinautoBackstageBackend`
  explicitly when UIA navigation is needed.

Backend instances are created once during ``__init__`` and stored on the
proxy.  Callers may inject alternative backends (e.g. a pywinauto-based
implementation) via the optional keyword arguments.  Passing ``None`` for
any surface selects the default COM implementation for that surface (or
``None`` for ``backstage_nav``, which has no COM default).
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
from typing import TYPE_CHECKING

# Third-party imports
from ezplog.lib_mode import get_logger

from ..utils._com_utils import assert_main_thread, wrap_com_error

# Local imports
from ._protocols import (
    AbstractBackstageFileOps,
    AbstractBackstageNavigator,
    AbstractDialogBackend,
    AbstractKeysBackend,
    AbstractMenuBackend,
    AbstractRibbonBackend,
)
from .win32com._backstage import COMBackstageBackend
from .win32com._dialog import DialogProxy
from .win32com._menu import MenuProxy
from .win32com._ribbon import RibbonProxy

if TYPE_CHECKING:
    from ..core._excel_app import ExcelApp

# ///////////////////////////////////////////////////////////////
# CONSTANTS
# ///////////////////////////////////////////////////////////////

logger = get_logger(__name__)

# ///////////////////////////////////////////////////////////////
# INTERNAL CLASSES
# ///////////////////////////////////////////////////////////////


class _COMKeysBackend(AbstractKeysBackend):
    """COM-based keystroke injection backend using ``Application.SendKeys``.

    Wraps ``Application.SendKeys(keys, wait)`` as an
    :class:`AbstractKeysBackend` instance so that :class:`GUIProxy`
    can hold it as an injected backend alongside the other proxy backends.

    This class is internal. It is not exported from ``ezxl.gui`` or
    ``ezxl`` and must not be referenced by consumer code.

    Args:
        app: The active ``ExcelApp`` instance that owns this backend.
    """

    # ///////////////////////////////////////////////////////////////
    # INIT
    # ///////////////////////////////////////////////////////////////

    def __init__(self, app: ExcelApp) -> None:
        self._app = app

    # ///////////////////////////////////////////////////////////////
    # PUBLIC METHODS
    # ///////////////////////////////////////////////////////////////

    @wrap_com_error
    def send_keys(self, keys: str, wait: bool = True) -> None:
        """Send a keystroke sequence to the Excel Application window.

        Delegates to ``Application.SendKeys(keys, wait)`` directly. The
        ``keys`` string must use standard VBA SendKeys notation
        (e.g. ``"{ENTER}"``, ``"^s"`` for Ctrl+S, ``"%{F4}"`` for Alt+F4).

        Args:
            keys: Keystroke string in VBA SendKeys notation.
            wait: If ``True``, block until Excel processes the keystrokes
                before returning. Defaults to ``True``.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            COMOperationError: If the SendKeys call fails.

        Example:
            >>> backend = _COMKeysBackend(xl)
            >>> backend.send_keys("^s")
            >>> backend.send_keys("{ESCAPE}", wait=False)
        """
        assert_main_thread(self._app._thread_id)
        logger.debug("_COMKeysBackend.send_keys: keys=%r, wait=%r", keys, wait)
        self._app._get_app().SendKeys(keys, wait)


# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


class GUIProxy:
    """Unified facade for GUI-level Excel interaction.

    Instantiated by ``ExcelApp.gui`` and bundles all six automation
    surfaces (ribbon, menu, dialog, keys, backstage, backstage_nav)
    under a single object.

    Backend injection
    -----------------
    Each surface can be replaced at construction time by passing an
    alternative implementation that satisfies the corresponding abstract
    protocol.  This is the primary extension point for non-COM backends
    such as pywinauto.  Passing ``None`` (the default) selects the
    standard COM implementation for that surface — except for
    ``backstage_nav``, which defaults to ``None`` (no UIA navigator).

    The ``backstage`` and ``backstage_nav`` surfaces are intentionally
    separate:

    - ``backstage`` (``AbstractBackstageFileOps``) — file I/O operations
      via COM.  Focus-independent, locale-independent.  Always present.
    - ``backstage_nav`` (``AbstractBackstageNavigator | None``) — visual
      panel navigation via UIA.  Required for ``open_options`` and
      ``open_save_as_panel``.  Inject explicitly when needed.

    Args:
        app: The active ``ExcelApp`` instance that owns this proxy.
        ribbon: Optional ribbon backend. Defaults to
            :class:`~ezxl.gui.RibbonProxy` when ``None``.
        menu: Optional menu backend. Defaults to
            :class:`~ezxl.gui.MenuProxy` when ``None``.
        dialog: Optional dialog backend. Defaults to
            :class:`~ezxl.gui.DialogProxy` when ``None``.
        keys: Optional keys backend. Defaults to
            :class:`~ezxl.gui._gui_proxy._COMKeysBackend` when ``None``.
        backstage: Optional Backstage file-ops backend. Defaults to
            :class:`~ezxl.gui.win32com._backstage.COMBackstageBackend` when
            ``None``.
        backstage_nav: Optional Backstage UIA navigator. Defaults to
            ``None`` (no UIA navigator).  Pass a
            :class:`~ezxl.gui.pywinauto.PywinautoBackstageBackend` to enable
            ``open_options`` and ``open_save_as_panel``.

    Security note:
        When injecting pywinauto backends, always pass ``hwnd=app.hwnd``
        to bind the backend to the exact Excel window managed by this
        session.  Omitting ``hwnd`` causes pywinauto to attach to the
        first Excel window it finds, which may not be the correct one
        when multiple Excel instances are running.

    Example:
        >>> with ExcelApp(mode="attach") as xl:
        ...     xl.gui.ribbon.execute("FileSave")
        ...     xl.gui.backstage.save()
        ...     path = xl.gui.dialog.get_file_open()
        ...     xl.gui.send_keys("^{HOME}")

    Example — with UIA navigator::

        gui = GUIProxy(
            xl,
            backstage=COMBackstageBackend(xl),
            backstage_nav=PywinautoBackstageBackend(hwnd=xl.hwnd, locale="fr"),
        )
        gui.backstage.save()                        # COM: direct save
        gui.backstage.save_as(path="out.xlsx")      # COM: format-aware save
        gui.backstage_nav.open_options()            # UIA: opens Options panel
        gui.backstage_nav.open_save_as_panel()      # UIA: opens panel only
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
        backstage: AbstractBackstageFileOps | None = None,
        backstage_nav: AbstractBackstageNavigator | None = None,
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
        # COM backend for file operations — always present.
        self._backstage: AbstractBackstageFileOps = (
            backstage if backstage is not None else COMBackstageBackend(app)
        )
        # UIA navigator — optional; no default COM equivalent exists.
        self._backstage_nav: AbstractBackstageNavigator | None = backstage_nav

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

    @property
    def backstage(self) -> AbstractBackstageFileOps:
        """Return the Backstage file-ops backend.

        Handles ``save``, ``save_as``, ``open_file``, and
        ``close_workbook`` via the COM object model — focus-independent
        and locale-independent.

        The default backend is
        :class:`~ezxl.gui.win32com._backstage.COMBackstageBackend`.
        Replace it at construction time to swap the implementation::

            gui = GUIProxy(xl, backstage=MyCustomFileOpsBackend(xl))

        Returns:
            AbstractBackstageFileOps: The configured file-ops backend.

        Example:
            >>> xl.gui.backstage.save()
            >>> xl.gui.backstage.save_as(path="C:\\\\output.xlsx")
        """
        return self._backstage

    @property
    def backstage_nav(self) -> AbstractBackstageNavigator | None:
        """Return the Backstage UIA navigator, or ``None`` if not injected.

        Handles ``open_options``, ``open_save_as_panel``, ``open_file``,
        and ``close_workbook`` via UIA direct click.  This backend must be
        injected explicitly — there is no COM default::

            gui = GUIProxy(
                xl,
                backstage_nav=PywinautoBackstageBackend(
                    hwnd=xl.hwnd, locale="fr"
                ),
            )

        Returns:
            AbstractBackstageNavigator | None: The configured UIA navigator,
                or ``None`` if none was provided.

        Example:
            >>> if xl.gui.backstage_nav is not None:
            ...     xl.gui.backstage_nav.open_options()
        """
        return self._backstage_nav

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
