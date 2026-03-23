# ///////////////////////////////////////////////////////////////
# gui.pywinauto - pywinauto-based GUI backends for EzXl
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
EzXl pywinauto GUI backends — UI Automation alternatives to COM backends.

Provides four backend classes that implement the GUI protocol interfaces
defined in :mod:`ezxl.gui._protocols` using ``pywinauto`` UI Automation
instead of ``win32com``.  These backends operate at the OS UI level and
do **not** require a COM connection.

Backends
--------
- :class:`PywinautoRibbonBackend` — Ribbon command execution by clicking
  UI Automation controls.
- :class:`PywinautoMenuBackend` — Menu bar traversal by caption-path clicks.
- :class:`PywinautoDialogBackend` — File pickers via keyboard shortcuts +
  Windows common dialog interaction.  Alert via Win32 ``MessageBoxW``.
- :class:`PywinautoKeysBackend` — Keystroke injection via
  ``pywinauto.keyboard.send_keys``.

Dependency
----------
All backends in this package require ``pywinauto``::

    pip install pywinauto

``pywinauto`` is an **optional** dependency of ``ezxl`` and is NOT listed
in ``[project.dependencies]``.  Importing any symbol from this package
will raise :exc:`ImportError` with an install hint if pywinauto is absent.

Usage::

    from ezxl import ExcelApp, GUIProxy
    from ezxl.gui.pywinauto import PywinautoRibbonBackend, PywinautoKeysBackend

    with ExcelApp(mode="attach") as xl:
        gui = GUIProxy(
            xl,
            ribbon=PywinautoRibbonBackend(),
            keys=PywinautoKeysBackend(),
        )
        gui.ribbon.execute("FileSave")   # via pywinauto UI click
        gui.send_keys("^{HOME}")         # via pywinauto keyboard
        # COM backends still used for menu and dialog (not overridden):
        gui.menu.list_bars()
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Local imports
from ._dialog import PywinautoDialogBackend
from ._keys import PywinautoKeysBackend
from ._menu import PywinautoMenuBackend
from ._ribbon import PywinautoRibbonBackend

__all__ = [
    "PywinautoRibbonBackend",
    "PywinautoMenuBackend",
    "PywinautoDialogBackend",
    "PywinautoKeysBackend",
]
