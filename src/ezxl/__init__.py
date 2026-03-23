# ///////////////////////////////////////////////////////////////
# EzXl - Main Module
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""EzXl — Generic Excel automation library via COM and openpyxl.

Provides a clean Python interface for:

- Opening and closing Excel files via COM (``ExcelApp``, ``WorkbookProxy``)
- Attaching to an already-running Excel instance
- Navigating worksheets and manipulating cells/ranges (``SheetProxy``,
  ``CellProxy``, ``RangeProxy``)
- Converting between file formats without a live Excel process
  (``read_excel``, ``read_csv``, ``xlsx_to_csv``, ``csv_to_xlsx``,
  ``read_sheet``)
- Formatting closed workbook files via openpyxl (``ExcelFormatter``)

Requires Python 3.11+ and a 64-bit Excel installation for COM features.
Format conversion and closed-file formatting work without Excel installed.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
import sys

# Local imports
from .version import __version__

# ///////////////////////////////////////////////////////////////
# METADATA INFORMATION
# ///////////////////////////////////////////////////////////////

__author__ = "Neuraaak"
__maintainer__ = "Neuraaak"
__description__ = "EzXl — Generic Excel automation library via COM and openpyxl"
__python_requires__ = ">=3.11"
__keywords__ = [
    "excel",
    "automation",
    "com",
    "openpyxl",
    "xlsx",
    "win32com",
    "spreadsheet",
    "office",
]
__url__ = "https://github.com/neuraaak/EzXl"
__repository__ = "https://github.com/neuraaak/EzXl"

# ///////////////////////////////////////////////////////////////
# PYTHON VERSION CHECK
# ///////////////////////////////////////////////////////////////

if sys.version_info < (3, 11):  # noqa: UP036
    raise RuntimeError(
        f"EzXl {__version__} requires Python 3.11 or higher. "
        f"Current version: {sys.version}"
    )

# ///////////////////////////////////////////////////////////////
# PUBLIC API — cross-platform (pure Python, no COM)
# ///////////////////////////////////////////////////////////////

from .exceptions import (
    COMOperationError,
    ExcelNotAvailableError,
    ExcelSessionLostError,
    ExcelThreadViolationError,
    EzXlError,
    FormatterError,
    GUIOperationError,
    SheetNotFoundError,
    WorkbookNotFoundError,
)
from .gui._protocols import (
    AbstractDialogBackend,
    AbstractKeysBackend,
    AbstractMenuBackend,
    AbstractRibbonBackend,
)
from .io._converters import csv_to_xlsx, read_csv, read_excel, read_sheet, xlsx_to_csv
from .io._formatters import ExcelFormatter

# ///////////////////////////////////////////////////////////////
# PUBLIC API — Windows only (COM / win32com / pywinauto)
# ///////////////////////////////////////////////////////////////

if sys.platform == "win32":
    from .core._excel_app import ExcelApp
    from .core._sheet import CellProxy, RangeProxy, SheetProxy
    from .core._workbook import WorkbookProxy
    from .gui._gui_proxy import GUIProxy
    from .gui.pywinauto import (
        PywinautoDialogBackend,
        PywinautoKeysBackend,
        PywinautoMenuBackend,
        PywinautoRibbonBackend,
    )
    from .gui.win32com._dialog import DialogProxy
    from .gui.win32com._menu import MenuProxy
    from .gui.win32com._ribbon import RibbonProxy

__all__ = [
    # Exceptions
    "EzXlError",
    "ExcelNotAvailableError",
    "ExcelSessionLostError",
    "ExcelThreadViolationError",
    "WorkbookNotFoundError",
    "SheetNotFoundError",
    "COMOperationError",
    "FormatterError",
    "GUIOperationError",
    # GUI protocols (cross-platform ABCs)
    "AbstractRibbonBackend",
    "AbstractMenuBackend",
    "AbstractDialogBackend",
    "AbstractKeysBackend",
    # Closed-file utilities
    "ExcelFormatter",
    "read_excel",
    "read_csv",
    "xlsx_to_csv",
    "csv_to_xlsx",
    "read_sheet",
]

if sys.platform == "win32":
    __all__ += [
        # COM automation
        "ExcelApp",
        "WorkbookProxy",
        "SheetProxy",
        "CellProxy",
        "RangeProxy",
        # GUI interaction
        "GUIProxy",
        "RibbonProxy",
        "MenuProxy",
        "DialogProxy",
        # GUI — pywinauto backends
        "PywinautoRibbonBackend",
        "PywinautoMenuBackend",
        "PywinautoDialogBackend",
        "PywinautoKeysBackend",
    ]
