# ///////////////////////////////////////////////////////////////
# test_package - EzXl public API smoke tests
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
Package-level smoke tests.

Validates that the ``ezxl`` package can be imported cleanly, that
``__version__`` is well-formed, and that every symbol declared in
``__all__`` is reachable as an attribute on the module.

These tests do not instantiate any class or call any function — they
are pure import-time checks that are safe to run in CI without an
Excel installation.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
import importlib
import types

# Third-party imports
import pytest

# ///////////////////////////////////////////////////////////////
# TESTS — package import
# ///////////////////////////////////////////////////////////////


@pytest.mark.unit
def test_should_import_package_without_error() -> None:
    """Verify that ``import ezxl`` succeeds without raising any exception."""
    module = importlib.import_module("ezxl")
    assert isinstance(module, types.ModuleType)


@pytest.mark.unit
def test_should_expose_version_string() -> None:
    """Verify that ``ezxl.__version__`` is a non-empty string."""
    import ezxl

    assert hasattr(ezxl, "__version__")
    assert isinstance(ezxl.__version__, str)
    assert len(ezxl.__version__) > 0


# ///////////////////////////////////////////////////////////////
# TESTS — __all__ symbol groups
# ///////////////////////////////////////////////////////////////


@pytest.mark.unit
def test_should_export_all_com_proxy_symbols() -> None:
    """Verify that the core COM proxy classes are exported from ``ezxl``."""
    import ezxl

    com_proxy_symbols = [
        "ExcelApp",
        "WorkbookProxy",
        "SheetProxy",
        "CellProxy",
        "RangeProxy",
    ]
    for name in com_proxy_symbols:
        assert hasattr(ezxl, name), f"ezxl.{name} not found"


@pytest.mark.unit
def test_should_export_all_exception_symbols() -> None:
    """Verify that all exception classes are exported from ``ezxl``."""
    import ezxl

    exception_symbols = [
        "EzXlError",
        "ExcelNotAvailableError",
        "ExcelSessionLostError",
        "ExcelThreadViolationError",
        "WorkbookNotFoundError",
        "SheetNotFoundError",
        "COMOperationError",
        "FormatterError",
        "GUIOperationError",
    ]
    for name in exception_symbols:
        assert hasattr(ezxl, name), f"ezxl.{name} not found"


@pytest.mark.unit
def test_should_export_all_gui_protocol_symbols() -> None:
    """Verify that the four abstract GUI protocol classes are exported."""
    import ezxl

    protocol_symbols = [
        "AbstractRibbonBackend",
        "AbstractMenuBackend",
        "AbstractDialogBackend",
        "AbstractKeysBackend",
    ]
    for name in protocol_symbols:
        assert hasattr(ezxl, name), f"ezxl.{name} not found"


@pytest.mark.unit
def test_should_export_all_gui_win32com_symbols() -> None:
    """Verify that the win32com GUI backend classes are exported."""
    import ezxl

    win32com_symbols = ["GUIProxy", "RibbonProxy", "MenuProxy", "DialogProxy"]
    for name in win32com_symbols:
        assert hasattr(ezxl, name), f"ezxl.{name} not found"


@pytest.mark.unit
def test_should_export_all_gui_pywinauto_symbols() -> None:
    """Verify that the pywinauto GUI backend classes are exported."""
    import ezxl

    pywinauto_symbols = [
        "PywinautoRibbonBackend",
        "PywinautoMenuBackend",
        "PywinautoDialogBackend",
        "PywinautoKeysBackend",
    ]
    for name in pywinauto_symbols:
        assert hasattr(ezxl, name), f"ezxl.{name} not found"


@pytest.mark.unit
def test_should_export_all_io_symbols() -> None:
    """Verify that the I/O utility functions and formatter are exported."""
    import ezxl

    io_symbols = [
        "ExcelFormatter",
        "read_excel",
        "read_csv",
        "xlsx_to_csv",
        "csv_to_xlsx",
        "read_sheet",
    ]
    for name in io_symbols:
        assert hasattr(ezxl, name), f"ezxl.{name} not found"


# ///////////////////////////////////////////////////////////////
# TESTS — __all__ completeness
# ///////////////////////////////////////////////////////////////


@pytest.mark.unit
def test_should_have_no_missing_all_entries() -> None:
    """Verify every name in ``ezxl.__all__`` is accessible on the module.

    This test catches the case where a symbol is declared in ``__all__``
    but its import statement in ``__init__.py`` is missing or broken.
    """
    import ezxl

    assert hasattr(ezxl, "__all__"), "ezxl.__all__ is not defined"

    missing = [name for name in ezxl.__all__ if not hasattr(ezxl, name)]
    assert missing == [], (
        f"Names declared in ezxl.__all__ but not accessible as attributes: {missing}"
    )
