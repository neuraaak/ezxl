# ///////////////////////////////////////////////////////////////
# _keys - SendKeys wrapper for Application.SendKeys
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
SendKeys thin wrapper — ``Application.SendKeys``.

A single public function, ``send_keys``, that delegates directly to
``Application.SendKeys(keys, wait)``. No key-sequence transformation or
macro-expansion is performed here; callers are responsible for composing
valid SendKeys strings.

This module is intentionally minimal. Complex keyboard automation should
prefer VBA macros (via ``ExcelApp.run_macro``) or ribbon execution
(via ``RibbonProxy.execute``) where possible, since ``SendKeys`` is
inherently fragile when the target application loses focus.
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
from ...utils._com_utils import assert_main_thread, wrap_com_error
from .._protocols import AbstractKeysBackend

if TYPE_CHECKING:
    from ...core._excel_app import ExcelApp

# ///////////////////////////////////////////////////////////////
# CONSTANTS
# ///////////////////////////////////////////////////////////////

logger = get_logger(__name__)

# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


class _COMKeysBackend(AbstractKeysBackend):
    """COM-based keystroke injection backend using ``Application.SendKeys``.

    Wraps the module-level :func:`send_keys` function as an
    :class:`AbstractKeysBackend` instance so that :class:`~ezxl.gui.GUIProxy`
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

    def send_keys(self, keys: str, wait: bool = True) -> None:
        """Send a keystroke sequence to the Excel Application window.

        Delegates to the module-level :func:`send_keys` function.

        Args:
            keys: Keystroke string in VBA SendKeys notation
                (e.g. ``"{ENTER}"``, ``"^s"`` for Ctrl+S).
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
        logger.debug("_COMKeysBackend.send_keys: keys=%r, wait=%r", keys, wait)
        send_keys(self._app, keys, wait)


# ///////////////////////////////////////////////////////////////
# FUNCTIONS
# ///////////////////////////////////////////////////////////////


@wrap_com_error
def send_keys(app: ExcelApp, keys: str, wait: bool = True) -> None:
    """Send a keystroke sequence to the Excel Application window.

    Wraps ``Application.SendKeys(keys, wait)`` directly. The ``keys``
    string must use standard VBA SendKeys notation
    (e.g. ``"{ENTER}"``, ``"^s"`` for Ctrl+S, ``"%{F4}"`` for Alt+F4).

    Args:
        app: The active ``ExcelApp`` instance.
        keys: Keystroke string in VBA SendKeys notation.
        wait: If ``True``, block until Excel processes the keystrokes
            before returning. Defaults to ``True``.

    Raises:
        ExcelThreadViolationError: If called from the wrong thread.
        COMOperationError: If the SendKeys call fails.

    Example:
        >>> send_keys(xl, "^s")          # Ctrl+S
        >>> send_keys(xl, "{ENTER}")
        >>> send_keys(xl, "%{F4}", wait=False)  # Alt+F4, non-blocking
    """
    assert_main_thread(app._thread_id)
    logger.debug("send_keys: keys=%r, wait=%r", keys, wait)
    app._get_app().SendKeys(keys, wait)
