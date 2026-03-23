# ///////////////////////////////////////////////////////////////
# gui.win32com - win32com-based GUI backends for EzXl
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
EzXl win32com GUI backends — COM-based implementations of the GUI protocol interfaces.

Provides three backend classes that implement the GUI protocol interfaces
defined in :mod:`ezxl.gui._protocols` using ``win32com`` / Excel's COM
object model.  These backends require an active COM connection to a running
Excel instance.

Backends
--------
- :class:`RibbonProxy` — Ribbon command execution and state queries via
  ``Application.CommandBars.*Mso`` methods.
- :class:`MenuProxy` — Legacy CommandBar traversal and control execution
  via ``Application.CommandBars``.
- :class:`DialogProxy` — File pickers via ``Application.GetOpenFilename``
  / ``Application.GetSaveAsFilename`` and alert via Win32 ``MessageBoxW``.

Note
----
The ``_COMKeysBackend`` class is intentionally **not** exported from this
package.  It is internal to :mod:`ezxl.gui.win32com._keys` and must not
be referenced by consumer code.

Usage::

    from ezxl import ExcelApp, GUIProxy
    from ezxl.gui.win32com import RibbonProxy, MenuProxy, DialogProxy

    with ExcelApp(mode="attach") as xl:
        gui = GUIProxy(
            xl,
            ribbon=RibbonProxy(xl),
            menu=MenuProxy(xl),
            dialog=DialogProxy(xl),
        )
        gui.ribbon.execute("FileSave")
        gui.menu.list_bars()
        path = gui.dialog.get_file_open(title="Select report")
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Local imports
from ._dialog import DialogProxy
from ._menu import MenuProxy
from ._ribbon import RibbonProxy

__all__ = [
    "RibbonProxy",
    "MenuProxy",
    "DialogProxy",
]
