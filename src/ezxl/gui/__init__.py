# ///////////////////////////////////////////////////////////////
# gui - GUI-level Excel interaction subpackage
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
EzXl GUI subpackage — ribbon, menu, dialog, and key automation via COM.

Provides ``GUIProxy`` as the single entry point, accessible through
``ExcelApp.gui``. The proxy bundles four interaction surfaces:

- ``GUIProxy.ribbon`` — MSO ribbon command execution and state queries.
- ``GUIProxy.menu``   — Legacy CommandBar traversal and control execution.
- ``GUIProxy.dialog`` — File-open, file-save, and message-box dialogs.
- ``GUIProxy.send_keys(keys)`` — Direct ``Application.SendKeys`` pass-through.

All COM calls are wrapped via ``wrap_com_error``. Thread identity is asserted
on every public method call using ``assert_main_thread``.

Typical usage::

    with ExcelApp(mode="attach") as xl:
        xl.gui.ribbon.execute("FileSave")
        path = xl.gui.dialog.get_file_open()
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
import sys

# Local imports — cross-platform (pure Python ABCs)
from ._protocols import (
    AbstractDialogBackend,
    AbstractKeysBackend,
    AbstractMenuBackend,
    AbstractRibbonBackend,
)

__all__ = [
    "AbstractRibbonBackend",
    "AbstractMenuBackend",
    "AbstractDialogBackend",
    "AbstractKeysBackend",
]

# Local imports — Windows only (COM / pywinauto)
if sys.platform == "win32":
    from ._gui_proxy import GUIProxy
    from .pywinauto import (
        PywinautoDialogBackend,
        PywinautoKeysBackend,
        PywinautoMenuBackend,
        PywinautoRibbonBackend,
    )
    from .win32com import DialogProxy, MenuProxy, RibbonProxy

    __all__ += [
        "GUIProxy",
        "RibbonProxy",
        "MenuProxy",
        "DialogProxy",
        "PywinautoRibbonBackend",
        "PywinautoMenuBackend",
        "PywinautoDialogBackend",
        "PywinautoKeysBackend",
    ]
