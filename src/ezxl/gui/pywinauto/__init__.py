# ///////////////////////////////////////////////////////////////
# gui.pywinauto - pywinauto-based GUI backends for EzXl
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
EzXl pywinauto GUI backends — UI Automation alternatives to COM backends.

Provides two backend classes that implement the GUI protocol interfaces
defined in :mod:`ezxl.gui._protocols` using ``pywinauto`` UI Automation
instead of ``win32com``.  These backends operate at the OS UI level and
do **not** require a COM connection.

Backends
--------
- :class:`PywinautoKeysBackend` — Keystroke injection via
  ``pywinauto.keyboard.send_keys``.
- :class:`PywinautoBackstageBackend` — Excel Backstage (File tab) navigation
  via Alt-sequences and UIA.  Extensible via ``_ELEMENTS`` class attribute.

Dependency
----------
All backends in this package require ``pywinauto``::

    pip install pywinauto

``pywinauto`` is an **optional** dependency of ``ezxl`` and is NOT listed
in ``[project.dependencies]``.  Importing any symbol from this package
will raise :exc:`ImportError` with an install hint if pywinauto is absent.

Usage::

    from ezxl import ExcelApp, GUIProxy
    from ezxl.gui.pywinauto import PywinautoKeysBackend

    with ExcelApp(mode="attach") as xl:
        gui = GUIProxy(
            xl,
            keys=PywinautoKeysBackend(),
        )
        gui.send_keys("^{HOME}")         # via pywinauto keyboard
        # COM backends still used for ribbon, menu and dialog (not overridden):
        gui.menu.list_bars()
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Local imports
from ._backstage import PywinautoBackstageBackend
from ._keys import PywinautoKeysBackend

__all__ = [
    "PywinautoKeysBackend",
    "PywinautoBackstageBackend",
]
