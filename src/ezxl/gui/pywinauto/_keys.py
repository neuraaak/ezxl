# ///////////////////////////////////////////////////////////////
# _keys - PywinautoKeysBackend: keystroke injection via pywinauto
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
PywinautoKeysBackend — keystroke injection via ``pywinauto.keyboard``.

Implements :class:`~ezxl.gui._protocols.AbstractKeysBackend` using
``pywinauto.keyboard.send_keys`` rather than ``Application.SendKeys``.

Key notation translation
------------------------
VBA SendKeys and pywinauto share the same modifier prefixes (``^`` for
Ctrl, ``%`` for Alt, ``+`` for Shift) and most special-key braces.  The
only common divergence is ``{ESCAPE}`` which pywinauto spells ``{ESC}``.
The private :func:`_translate_keys` function normalises the most
frequent VBA patterns to their pywinauto equivalents before injection.

Limitations
-----------
- The *wait* parameter has no direct equivalent in ``pywinauto.keyboard``.
  When ``wait=True`` (default), a brief synchronous pause (50 ms) is
  inserted after the keystroke sequence.  This is a best-effort approximation
  and may not be sufficient for slow Excel responses.
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
import time

# Local imports
from .._protocols import AbstractKeysBackend

# ///////////////////////////////////////////////////////////////
# OPTIONAL DEPENDENCY GUARD
# ///////////////////////////////////////////////////////////////

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

# Pause duration (seconds) used to approximate wait=True behaviour.
# pywinauto.keyboard.send_keys has no built-in "wait for application"
# mechanism, so we sleep briefly after sending.
_WAIT_PAUSE: float = 0.05

# VBA SendKeys → pywinauto translation table.
# Only entries that differ between the two notations are listed.
# Identical entries (^, %, +, {ENTER}, {HOME}, {END}, {TAB}, {F1}…{F12})
# are passed through unchanged.
_VBA_TO_PW: dict[str, str] = {
    "{ESCAPE}": "{ESC}",
}

# ///////////////////////////////////////////////////////////////
# FUNCTIONS
# ///////////////////////////////////////////////////////////////


def _translate_keys(keys: str) -> str:
    """Translate a VBA SendKeys string to pywinauto keyboard notation.

    The two notations are almost identical.  This function corrects the
    known divergences so that callers can use familiar VBA key strings
    with the pywinauto backend.

    Translation table:

    +-------------+-------------+--------------------------------------+
    | VBA token   | pywinauto   | Note                                 |
    +=============+=============+======================================+
    | ``^``       | ``^``       | Ctrl — identical                     |
    +-------------+-------------+--------------------------------------+
    | ``%``       | ``%``       | Alt — identical                      |
    +-------------+-------------+--------------------------------------+
    | ``+``       | ``+``       | Shift — identical                    |
    +-------------+-------------+--------------------------------------+
    | ``{ENTER}`` | ``{ENTER}`` | Enter — identical                    |
    +-------------+-------------+--------------------------------------+
    | ``{ESCAPE}``| ``{ESC}``   | **differs** — VBA uses full word     |
    +-------------+-------------+--------------------------------------+
    | ``{HOME}``  | ``{HOME}``  | identical                            |
    +-------------+-------------+--------------------------------------+
    | ``{END}``   | ``{END}``   | identical                            |
    +-------------+-------------+--------------------------------------+
    | ``{TAB}``   | ``{TAB}``   | identical                            |
    +-------------+-------------+--------------------------------------+
    | ``{F1}``…   | ``{F1}``…   | identical                            |
    | ``{F12}``   | ``{F12}``   |                                      |
    +-------------+-------------+--------------------------------------+

    Args:
        keys: Keystroke string in VBA SendKeys notation.

    Returns:
        str: Equivalent string in pywinauto keyboard notation.

    Example:
        >>> _translate_keys("{ESCAPE}")
        '{ESC}'
        >>> _translate_keys("^{HOME}")
        '^{HOME}'
        >>> _translate_keys("%{F4}")
        '%{F4}'
    """
    result = keys
    for vba_token, pw_token in _VBA_TO_PW.items():
        result = result.replace(vba_token, pw_token)
    return result


# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


class PywinautoKeysBackend(AbstractKeysBackend):
    """Keystroke injection via ``pywinauto.keyboard.send_keys``.

    Translates VBA SendKeys notation to pywinauto notation using
    :func:`_translate_keys`, then delegates to
    ``pywinauto.keyboard.send_keys``.

    This backend is a standalone alternative to the COM-based
    ``_COMKeysBackend``.  It does **not** require an
    :class:`~ezxl.core.ExcelApp` instance and carries no COM STA
    thread constraint.

    Args:
        hwnd: Win32 window handle for the Excel main window.  Currently
            unused — ``pywinauto.keyboard.send_keys`` injects keystrokes
            into the currently focused window.  Ensure the Excel window
            has focus before calling :meth:`send_keys`.  The parameter
            is accepted for API consistency with the other pywinauto
            backends and reserved for future use.

    Example:
        >>> from ezxl.gui.pywinauto import PywinautoKeysBackend
        >>> keys = PywinautoKeysBackend()
        >>> keys.send_keys("^s")          # Ctrl+S
        >>> keys.send_keys("{ESCAPE}")    # maps to {ESC} internally
        >>> keys.send_keys("^{HOME}")

        >>> # Inject into GUIProxy:
        >>> from ezxl import ExcelApp, GUIProxy
        >>> with ExcelApp(mode="attach") as xl:
        ...     gui = GUIProxy(xl, keys=PywinautoKeysBackend())
        ...     gui.send_keys("^{HOME}")
    """

    # ///////////////////////////////////////////////////////////////
    # INIT
    # ///////////////////////////////////////////////////////////////

    def __init__(self, hwnd: int | None = None) -> None:
        # hwnd is reserved for future focus-management logic.
        self._hwnd = hwnd

    # ///////////////////////////////////////////////////////////////
    # PUBLIC METHODS
    # ///////////////////////////////////////////////////////////////

    def send_keys(self, keys: str, wait: bool = True) -> None:
        """Send a keystroke sequence using ``pywinauto.keyboard``.

        Translates *keys* from VBA SendKeys notation to pywinauto notation
        via :func:`_translate_keys`, then calls
        ``pywinauto.keyboard.send_keys``.  If *wait* is ``True``, a brief
        pause of 50 ms is inserted after injection as a best-effort
        approximation of the VBA ``wait=True`` semantics.

        Note:
            ``pywinauto.keyboard.send_keys`` injects keystrokes into
            the **currently focused window**.  Call
            ``window.set_focus()`` on the Excel window before using
            this backend if focus cannot be guaranteed.

        Args:
            keys: Keystroke string in VBA SendKeys notation
                (e.g. ``"{ENTER}"``, ``"^s"`` for Ctrl+S,
                ``"{ESCAPE}"`` for Escape).
            wait: If ``True``, insert a 50 ms pause after sending.
                This is a best-effort approximation; it does not
                guarantee Excel has finished processing.
                Defaults to ``True``.

        Raises:
            GUIOperationError: If the pywinauto keystroke injection
                raises an unexpected error.

        Example:
            >>> backend.send_keys("^s")             # Ctrl+S
            >>> backend.send_keys("{ESCAPE}")       # → {ESC} internally
            >>> backend.send_keys("%{F4}", wait=False)  # Alt+F4, no pause
        """
        translated = _translate_keys(keys)
        logger.debug(
            "PywinautoKeysBackend.send_keys: keys=%r, translated=%r, wait=%r",
            keys,
            translated,
            wait,
        )
        _pw_send_keys(translated)
        if wait:
            time.sleep(_WAIT_PAUSE)
