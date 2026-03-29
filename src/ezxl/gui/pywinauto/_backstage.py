# ///////////////////////////////////////////////////////////////
# _backstage - PywinautoBackstageBackend: Backstage navigation via UIA
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
PywinautoBackstageBackend — Excel Backstage (File tab) visual navigation via
``pywinauto`` UI Automation with Alt-sequence fallback.

The Backstage is the full-screen overlay opened via the Excel File ribbon
tab.  Excel does not expose ``AutomationId`` on Backstage controls.  This
backend navigates primarily via UIA direct click (locale-dependent but
focus-independent), with an Alt-key sequence fallback for environments
where UIA click is unreliable.

Implements :class:`~ezxl.gui._protocols.AbstractBackstageNavigator`, which
covers UIA-driven operations: opening the Options panel, opening the Save As
panel without confirming, and the UIA variants of open-file and close-workbook.

Strategy order
--------------
1. **Primary — UIA direct click**: locate the ``Button "Onglet Fichier"``
   (or locale equivalent) and click it to open the Backstage, then click
   the target ``ListItem`` inside the resulting ``List "Fichier"`` (or
   locale equivalent).  This strategy requires no keyboard focus and is
   robust against window-focus loss.
2. **Fallback — Alt-sequence**: send ``spec.alt_sequence`` via
   ``pywinauto.keyboard.send_keys``.  Only attempted when the UIA strategy
   fails *and* the spec provides a non-empty sequence.

All actions are flat methods — no object-tree traversal.

Responsibilities
----------------
This backend owns UIA-level navigation only.  File I/O operations
(``save``, ``save_as`` with a path, format-aware SaveAs) belong to
:class:`~ezxl.gui.win32com.COMBackstageBackend` via
:class:`~ezxl.gui._protocols.AbstractBackstageFileOps`.  The two backends
compose inside :class:`~ezxl.gui.GUIProxy` with no cross-dependency.

Extensibility
-------------
Consumer libraries extend the element registry without modifying ``ezxl``::

    from ezxl.gui.pywinauto._backstage import PywinautoBackstageBackend
    from ezxl.gui.pywinauto._registry import UIElementSpec

    _SAP_ELEMENTS = {"sap_logon": UIElementSpec(key="sap_logon", alt_sequence="%XL")}

    class HanaisBackstageBackend(PywinautoBackstageBackend):
        _ELEMENTS = PywinautoBackstageBackend._ELEMENTS | _SAP_ELEMENTS

        def sap_logon(self) -> None:
            self._execute_by_spec(self._get_spec("sap_logon"))

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
import time

# Local imports
from ...exceptions import GUIOperationError
from .._protocols import AbstractBackstageNavigator
from ._connect import _get_excel_window
from ._registry import (
    BACKSTAGE_ELEMENTS,
    FILE_BUTTON_NAMES,
    FILE_LIST_NAMES,
    UIElementSpec,
)

# ///////////////////////////////////////////////////////////////
# OPTIONAL DEPENDENCY GUARD
# ///////////////////////////////////////////////////////////////

try:
    from pywinauto.application import (  # type: ignore[import-untyped]
        WindowSpecification,
    )
    from pywinauto.base_wrapper import BaseWrapper  # type: ignore[import-untyped]
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

# Seconds to wait after clicking/sending a key before assuming Excel has
# processed the navigation.  Backstage animations in Excel can take 200–800 ms.
# This value is intentionally conservative.
_BACKSTAGE_SETTLE: float = 0.4

# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


class PywinautoBackstageBackend(AbstractBackstageNavigator):
    """Excel Backstage visual navigation via UIA direct click with Alt-sequence fallback.

    Implements :class:`~ezxl.gui._protocols.AbstractBackstageNavigator` using
    ``pywinauto``.  The primary strategy for every action is a UIA direct
    click — the backend opens the Backstage via ``Button "Onglet Fichier"``
    (or locale equivalent), then clicks the target ``ListItem`` by its
    localised UIA ``Name``.  This approach is focus-independent.

    An Alt-sequence fallback is attempted only when the UIA click fails
    **and** the element spec carries a non-empty ``alt_sequence``.

    This backend does **not** perform file I/O.  Operations that write to
    disk (``save``, ``save_as`` with an explicit path) belong to
    :class:`~ezxl.gui.win32com.COMBackstageBackend` and are exposed through
    ``GUIProxy.backstage``.  This backend is composed alongside it via
    ``GUIProxy.backstage_nav``.

    Args:
        hwnd: Win32 window handle for the Excel main window.  Pass
            ``None`` to auto-detect the first visible Excel instance.
            Always pass ``xl.hwnd`` in production to avoid targeting the
            wrong window when multiple Excel instances are open.
        locale: Locale code used for UIA ``Name``-based searches.
            Accepted values: ``"en"`` (default), ``"fr"``.

    Example:
        >>> backend = PywinautoBackstageBackend(hwnd=xl.hwnd, locale="fr")
        >>> backend.open_options()           # UIA: opens Options panel
        >>> backend.open_save_as_panel()     # UIA: opens Save As panel, leaves open
        >>> backend.open_file()              # UIA: opens Open panel
        >>> backend.close_workbook()         # UIA: clicks Close in Backstage
    """

    # Class-level element registry.  Consumer libraries override via:
    #   MyBackend._ELEMENTS = PywinautoBackstageBackend._ELEMENTS | MY_ELEMENTS
    _ELEMENTS: dict[str, UIElementSpec] = BACKSTAGE_ELEMENTS

    # ///////////////////////////////////////////////////////////////
    # INIT
    # ///////////////////////////////////////////////////////////////

    def __init__(
        self,
        hwnd: int | None = None,
        locale: str = "en",
    ) -> None:
        self._hwnd = hwnd
        self._locale = locale

    # ///////////////////////////////////////////////////////////////
    # PRIVATE HELPERS
    # ///////////////////////////////////////////////////////////////

    def _get_window(self) -> WindowSpecification:
        """Return the pywinauto WindowSpecification for the Excel window."""
        return _get_excel_window(self._hwnd)

    def _get_spec(self, key: str) -> UIElementSpec:
        """Retrieve a UIElementSpec by key.

        Args:
            key: Stable element identifier (e.g. ``"file_options"``).

        Raises:
            GUIOperationError: If the key is not found in the registry.
        """
        spec = self._ELEMENTS.get(key)
        if spec is None:
            raise GUIOperationError(
                f"Unknown backstage element key {key!r}. "
                f"Available keys: {sorted(self._ELEMENTS)}"
            )
        return spec

    def _is_backstage_open(self, window: WindowSpecification) -> bool:
        """Return whether the Backstage (File list) is currently open.

        Detects the presence of the ``List "Fichier"`` (or locale equivalent)
        control, which only exists in the UIA tree while the Backstage is open.

        Args:
            window: A pywinauto ``WindowSpecification`` for the Excel window.

        Returns:
            bool: ``True`` if the Backstage list is found; ``False`` otherwise.
        """
        list_name = FILE_LIST_NAMES.get(self._locale, FILE_LIST_NAMES["en"])
        try:
            window.child_window(title=list_name, control_type="List").wrapper_object()
            return True
        except Exception:
            return False

    def _ensure_backstage_open(self, window: WindowSpecification) -> BaseWrapper:
        """Open the Backstage if not already open and return its List control.

        Calls ``set_focus()`` on the window once before clicking the File tab
        button (focus is required for the UIA click to register).  Does not
        call ``set_focus()`` again for subsequent actions within the same
        Backstage session.

        Args:
            window: A pywinauto ``WindowSpecification`` for the Excel window.

        Returns:
            BaseWrapper: The pywinauto wrapper for the ``List "Fichier"`` control.

        Raises:
            GUIOperationError: If the Backstage cannot be opened or the list
                control cannot be located after clicking.
        """
        list_name = FILE_LIST_NAMES.get(self._locale, FILE_LIST_NAMES["en"])
        button_name = FILE_BUTTON_NAMES.get(self._locale, FILE_BUTTON_NAMES["en"])

        if not self._is_backstage_open(window):
            logger.debug(
                "_ensure_backstage_open: Backstage closed — clicking %r", button_name
            )
            window.set_focus()
            try:
                file_btn = window.child_window(title=button_name, control_type="Button")
                file_btn.click_input()
                time.sleep(_BACKSTAGE_SETTLE)
            except Exception as exc:
                raise GUIOperationError(
                    f"Could not click File tab button {button_name!r} "
                    f"to open the Backstage: {exc}",
                    cause=exc,
                ) from exc

        # Retrieve the list whether we just opened it or it was already open.
        # Return the wrapper_object() (ListViewWrapper) — children() works on it.
        try:
            spec = window.child_window(title=list_name, control_type="List")
            return spec.wrapper_object()
        except Exception as exc:
            raise GUIOperationError(
                f"Backstage list {list_name!r} not found after opening: {exc}",
                cause=exc,
            ) from exc

    def _click_item_in_list(self, file_list: BaseWrapper, item_name: str) -> None:
        """Click a ``ListItem`` in the Backstage list by its localised title.

        Args:
            file_list: The pywinauto wrapper for the Backstage ``List`` control.
            item_name: Localised UIA ``Name`` of the target ``ListItem``.

        Raises:
            GUIOperationError: If the item cannot be found or clicked.
        """
        # ListViewWrapper (returned by wrapper_object()) has no child_window().
        # Iterate children() and match by window_text() instead.
        try:
            for child in file_list.children():
                try:
                    if child.window_text().strip() == item_name:
                        child.click_input()
                        time.sleep(_BACKSTAGE_SETTLE)
                        return
                except Exception as child_exc:
                    logger.debug("_click_item_in_list: skipping child — %s", child_exc)
                    continue
            raise GUIOperationError(
                f"ListItem {item_name!r} not found in Backstage list. "
                f"Check locale setting and available items."
            )
        except GUIOperationError:
            raise
        except Exception as exc:
            raise GUIOperationError(
                f"Could not click Backstage item {item_name!r}: {exc}",
                cause=exc,
            ) from exc

    def _execute_by_spec(
        self,
        spec: UIElementSpec,
        locale: str | None = None,
    ) -> None:
        """Execute a UI action described by a UIElementSpec.

        Resolution order:

        1. UIA direct click (primary): open the Backstage via the File tab
           button, then click the target ``ListItem`` by its localised name.
        2. Alt-sequence fallback: send ``spec.alt_sequence`` via
           ``pywinauto.keyboard.send_keys``.  Only attempted when the UIA
           strategy fails *and* ``spec.alt_sequence`` is non-empty.

        Args:
            spec: Element descriptor from the registry.
            locale: Locale override for name resolution.
                Defaults to ``self._locale`` when ``None``.

        Raises:
            GUIOperationError: If all resolution strategies fail.
        """
        effective_locale = locale or self._locale
        item_name = spec.names.get(effective_locale)
        logger.debug(
            "PywinautoBackstageBackend._execute_by_spec: key=%r, item=%r, alt=%r",
            spec.key,
            item_name,
            spec.alt_sequence,
        )

        # Strategy 1 — UIA direct click (primary, focus-independent).
        if item_name:
            try:
                window: WindowSpecification = self._get_window()
                file_list: BaseWrapper = self._ensure_backstage_open(window)
                self._click_item_in_list(file_list, item_name)
                return
            except GUIOperationError:
                raise
            except Exception as exc:
                logger.debug(
                    "_execute_by_spec: UIA click for %r failed (%s) "
                    "— trying Alt-sequence fallback",
                    spec.key,
                    exc,
                )

        # Strategy 2 — Alt-sequence fallback (requires focus).
        if spec.alt_sequence:
            try:
                window = self._get_window()
                window.set_focus()
                _pw_send_keys(spec.alt_sequence)
                time.sleep(_BACKSTAGE_SETTLE)
                return
            except Exception as exc:
                logger.debug(
                    "_execute_by_spec: Alt-sequence %r failed: %s",
                    spec.alt_sequence,
                    exc,
                )

        raise GUIOperationError(
            f"Could not execute backstage action {spec.key!r}. "
            f"UIA click failed for locale {effective_locale!r} "
            f"(item name: {item_name!r}) and Alt-sequence fallback "
            f"{spec.alt_sequence!r} also failed or was not available."
        )

    # ///////////////////////////////////////////////////////////////
    # AbstractBackstageNavigator implementation
    # ///////////////////////////////////////////////////////////////

    def open_options(self) -> None:
        """Navigate to the Excel Options panel via the Backstage.

        Raises:
            GUIOperationError: If the Options panel cannot be reached.

        Example:
            >>> backend.open_options()
        """
        logger.debug("PywinautoBackstageBackend.open_options")
        self._execute_by_spec(self._get_spec("file_options"))

    def open_save_as_panel(self) -> None:
        """Open the Save As panel in the Backstage without confirming a save.

        Clicks the ``"Enregistrer sous"`` (or locale equivalent) ListItem
        and leaves the panel open.  Does not write to disk — use
        ``GUIProxy.backstage.save_as(path=...)`` for programmatic saves.

        Raises:
            GUIOperationError: If the panel cannot be reached.

        Example:
            >>> backend.open_save_as_panel()
        """
        logger.debug("PywinautoBackstageBackend.open_save_as_panel")
        self._execute_by_spec(self._get_spec("file_save_as"))

    def open_file(self) -> None:
        """Open the Open panel via the Backstage.

        Raises:
            GUIOperationError: If the panel cannot be reached.

        Example:
            >>> backend.open_file()
        """
        logger.debug("PywinautoBackstageBackend.open_file")
        self._execute_by_spec(self._get_spec("file_open"))

    def close_workbook(self) -> None:
        """Close the active workbook via a Backstage UIA click.

        Raises:
            GUIOperationError: If the action cannot be completed.

        Example:
            >>> backend.close_workbook()
        """
        logger.debug("PywinautoBackstageBackend.close_workbook")
        self._execute_by_spec(self._get_spec("file_close"))
