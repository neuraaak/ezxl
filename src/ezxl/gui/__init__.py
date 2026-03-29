# ///////////////////////////////////////////////////////////////
# gui - GUI-level Excel interaction subpackage
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
EzXl GUI subpackage — ribbon, menu, dialog, and key automation via COM.

Provides ``GUIProxy`` as the single entry point, accessible through
``ExcelApp.gui``. The proxy bundles six interaction surfaces:

- ``GUIProxy.ribbon``        — MSO ribbon command execution and state queries.
- ``GUIProxy.menu``          — Legacy CommandBar traversal and control execution.
- ``GUIProxy.dialog``        — File-open, file-save, and message-box dialogs.
- ``GUIProxy.send_keys(keys)`` — Direct ``Application.SendKeys`` pass-through.
- ``GUIProxy.backstage``     — File operations via COM (save, save_as, open_file,
  close_workbook). Default: :class:`~ezxl.gui.win32com.COMBackstageBackend`.
- ``GUIProxy.backstage_nav`` — UIA Backstage navigation (open_options,
  open_save_as_panel). Defaults to ``None``; inject a
  :class:`~ezxl.gui.pywinauto.PywinautoBackstageBackend` explicitly.

All COM calls are wrapped via ``wrap_com_error``. Thread identity is asserted
on every public method call using ``assert_main_thread``.

Typical usage::

    with ExcelApp(mode="attach") as xl:
        xl.gui.ribbon.execute("FileSave")
        path = xl.gui.dialog.get_file_open()
        xl.gui.backstage.save_as(path="C:\\\\output.xlsx")

With UIA navigator::

    gui = GUIProxy(
        xl,
        backstage_nav=PywinautoBackstageBackend(hwnd=xl.hwnd, locale="fr"),
    )
    gui.backstage_nav.open_options()
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
import sys

# Local imports — cross-platform (pure Python ABCs)
from ._protocols import (
    AbstractBackstageBackend,
    AbstractBackstageFileOps,
    AbstractBackstageNavigator,
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
    # Backstage protocols
    "AbstractBackstageFileOps",
    "AbstractBackstageNavigator",
    # Compatibility alias — kept for existing imports
    "AbstractBackstageBackend",
]

# Local imports — Windows only (COM / pywinauto)
if sys.platform == "win32":
    from ._gui_proxy import GUIProxy
    from .pywinauto import (
        PywinautoBackstageBackend,
        PywinautoKeysBackend,
    )
    from .win32com import COMBackstageBackend, DialogProxy, MenuProxy, RibbonProxy

    __all__ += [
        "GUIProxy",
        "COMBackstageBackend",
        "RibbonProxy",
        "MenuProxy",
        "DialogProxy",
        "PywinautoKeysBackend",
        "PywinautoBackstageBackend",
    ]
