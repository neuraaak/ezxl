# ///////////////////////////////////////////////////////////////
# _menu - PywinautoMenuBackend: menu navigation via UI Automation
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
PywinautoMenuBackend — Excel menu bar navigation via ``pywinauto``.

Implements :class:`~ezxl.gui._protocols.AbstractMenuBackend` using
``pywinauto`` UI Automation rather than ``Application.CommandBars``.  This
backend locates the Excel window's menu bar, clicks top-level items by
caption, and traverses sub-menus by the provided path.

Limitations
-----------
- :meth:`list_bars` enumerates the children of Excel's top-level menu bar
  control.  In modern Excel (2013+) this is the ribbon tab bar, not a legacy
  CommandBar.  The result may differ from :meth:`~ezxl.gui.MenuProxy.list_bars`
  which iterates the full ``CommandBars`` COM collection.  If the UI
  hierarchy is unexpected, an empty list is returned rather than raising.
- :meth:`list_controls` clicks the named top-level item to open its menu,
  collects child captions, then presses Escape to close.  This is a
  **side-effecting** operation — it briefly activates the menu in the Excel
  window.
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
import logging
from typing import Any

# Local imports
from ...exceptions import GUIOperationError
from .._protocols import AbstractMenuBackend
from ._connect import get_excel_window

# ///////////////////////////////////////////////////////////////
# OPTIONAL DEPENDENCY GUARD
# ///////////////////////////////////////////////////////////////

# pywinauto guard is in _connect.  No direct pywinauto import needed here.

try:
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

# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


class PywinautoMenuBackend(AbstractMenuBackend):
    """Excel menu bar navigation and control execution via ``pywinauto``.

    Traverses the Excel window's menu/ribbon tab bar by caption, clicking
    top-level items and sub-menu items in sequence.

    This backend is a standalone alternative to the COM-based
    :class:`~ezxl.gui.MenuProxy`.  It does **not** require an
    :class:`~ezxl.core.ExcelApp` instance and carries no COM STA
    thread constraint.

    Args:
        hwnd: Win32 window handle for the Excel main window.  If ``None``,
            the backend auto-detects the first visible Excel instance whose
            title matches ``".*- Microsoft Excel$"``.

    Example:
        >>> from ezxl.gui.pywinauto import PywinautoMenuBackend
        >>> menu = PywinautoMenuBackend()
        >>> menu.list_bars()
        ['File', 'Home', 'Insert', ...]
        >>> menu.click("File", "Save")

        >>> # Inject into GUIProxy:
        >>> from ezxl import ExcelApp, GUIProxy
        >>> with ExcelApp(mode="attach") as xl:
        ...     gui = GUIProxy(xl, menu=PywinautoMenuBackend())
        ...     gui.menu.list_bars()
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

    def _get_menu_bar(self, window: Any) -> Any:
        """Return the top-level menu/ribbon tab bar control.

        Attempts to locate the menu bar via the ``MenuBar`` control type.
        Falls back to a ``Tab`` control type which is what Excel 2013+
        exposes for its ribbon tab strip.

        Args:
            window: The pywinauto ``WindowSpecification`` for Excel.

        Returns:
            Any: The pywinauto control representing the menu/tab bar.

        Raises:
            GUIOperationError: If no menu bar control can be found.
        """
        for control_type in ("MenuBar", "Tab"):
            try:
                bar: Any = window.child_window(control_type=control_type)
                # Verify the control is reachable.
                bar.wrapper_object()
                return bar
            except Exception as exc:
                logger.debug(
                    "_get_menu_bar: control_type=%r not found: %s", control_type, exc
                )
                continue
        raise GUIOperationError(
            "Could not locate a menu bar or tab control in the Excel window. "
            "The UI hierarchy may differ from the expected structure."
        )

    # ///////////////////////////////////////////////////////////////
    # PUBLIC METHODS
    # ///////////////////////////////////////////////////////////////

    def click(self, bar_name: str, *item_path: str) -> None:
        """Traverse a menu by caption path and click the final item.

        Locates the top-level menu/tab bar, clicks the item matching
        *bar_name*, then clicks each caption in *item_path* within the
        resulting sub-menu.

        Side effects
        ------------
        This method activates the Excel window and briefly opens the
        named menu before clicking the target item.  Focus is transferred
        to Excel during the operation.

        Args:
            bar_name: Caption of the top-level menu item
                (e.g. ``"File"``, ``"Home"``).
            *item_path: One or more captions forming the path to the
                target item.  At least one caption is required.

        Raises:
            GUIOperationError: If the bar, any intermediate item, or the
                final item cannot be found, or if the click fails.

        Example:
            >>> menu.click("File", "Save")
            >>> menu.click("File", "Save As")
        """
        if not item_path:
            raise GUIOperationError(
                "click() requires at least one item caption in item_path."
            )
        logger.debug("PywinautoMenuBackend.click: bar=%r, path=%r", bar_name, item_path)
        try:
            window = self._get_window()
            window.set_focus()
            menu_bar = self._get_menu_bar(window)

            # Click the top-level bar item.
            top_item: Any = menu_bar.child_window(title=bar_name)
            top_item.click_input()

            # Traverse sub-menu path.
            for caption in item_path:
                sub_item: Any = window.child_window(
                    title=caption, control_type="MenuItem"
                )
                sub_item.click_input()

        except GUIOperationError:
            raise
        except Exception as exc:
            raise GUIOperationError(
                f"Failed to click menu path ({bar_name!r}, {item_path!r}): {exc}",
                cause=exc,
            ) from exc

    def list_bars(self) -> list[str]:
        """Return the captions of all top-level menu/ribbon tab items.

        Enumerates the children of Excel's top-level menu bar (or ribbon
        tab strip).  In modern Excel this returns the ribbon tab names
        (``"File"``, ``"Home"``, ``"Insert"``, …).

        This is a best-effort operation.  If the UI hierarchy is
        unexpected, an empty list is returned rather than raising.

        Returns:
            list[str]: Captions of top-level menu/tab items, in order.
                May be empty if the bar structure cannot be determined.

        Raises:
            GUIOperationError: If the Excel window cannot be found.

        Example:
            >>> menu.list_bars()
            ['File', 'Home', 'Insert', 'Page Layout', ...]
        """
        logger.debug("PywinautoMenuBackend.list_bars")
        try:
            window = self._get_window()
            menu_bar = self._get_menu_bar(window)
            children: list[Any] = menu_bar.children()
            names: list[str] = []
            for child in children:
                try:
                    name: str = child.window_text()
                    if name and name.strip():
                        names.append(name.strip())
                except Exception as exc:
                    logger.debug(
                        "PywinautoMenuBackend.list_bars: skipping inaccessible child: %s",
                        exc,
                    )
                    continue
            return names
        except GUIOperationError:
            raise
        except Exception as exc:
            logger.debug(
                "PywinautoMenuBackend.list_bars: returning empty list due to error: %s",
                exc,
            )
            return []

    def list_controls(self, bar_name: str) -> list[str]:
        """Return the captions of all items in a top-level menu.

        Clicks *bar_name* to open its drop-down, collects the captions of
        the resulting menu items, then presses Escape to close the menu.

        Side effects
        ------------
        This method briefly opens the named menu in the Excel window.
        Focus is transferred to Excel during the operation.

        Args:
            bar_name: Caption of the top-level menu item to inspect
                (e.g. ``"File"``, ``"Insert"``).

        Returns:
            list[str]: Captions of the menu items, in order.  Separators
                and items with no caption are excluded.

        Raises:
            GUIOperationError: If the bar cannot be found, the menu
                cannot be opened, or control enumeration fails.

        Example:
            >>> menu.list_controls("Insert")
            ['Tables', 'Illustrations', 'Charts', ...]
        """
        logger.debug("PywinautoMenuBackend.list_controls: bar=%r", bar_name)
        try:
            window = self._get_window()
            window.set_focus()
            menu_bar = self._get_menu_bar(window)

            # Click the top-level item to open its menu.
            top_item: Any = menu_bar.child_window(title=bar_name)
            top_item.click_input()

            # Collect MenuItem children from the opened menu panel.
            captions: list[str] = []
            try:
                items: list[Any] = window.children(control_type="MenuItem")
                for item in items:
                    try:
                        caption: str = item.window_text()
                        if caption and caption.strip():
                            captions.append(caption.strip())
                    except Exception as exc:
                        logger.debug(
                            "PywinautoMenuBackend.list_controls: "
                            "skipping inaccessible item: %s",
                            exc,
                        )
                        continue
            finally:
                # Always close the menu, even if collection fails.
                try:
                    _pw_send_keys("{ESC}")
                except Exception as exc:
                    logger.debug(
                        "PywinautoMenuBackend.list_controls: "
                        "could not send Escape to close menu: %s",
                        exc,
                    )

            return captions

        except GUIOperationError:
            raise
        except Exception as exc:
            raise GUIOperationError(
                f"Failed to list controls for menu item {bar_name!r}: {exc}",
                cause=exc,
            ) from exc
