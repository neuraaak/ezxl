# ///////////////////////////////////////////////////////////////
# test_exceptions - EzXl exception hierarchy tests
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
Unit tests for the EzXl exception hierarchy.

Validates that each exception class can be instantiated, carries the
expected message, stores an optional cause, and participates correctly
in the Python exception chaining mechanism.

No COM calls are made in this module — all tests are pure Python.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Third-party imports
import pytest

# Local imports
from ezxl.exceptions import (
    COMOperationError,
    ExcelNotAvailableError,
    EzXlError,
    FormatterError,
    GUIOperationError,
    SheetNotFoundError,
    WorkbookNotFoundError,
)

# ///////////////////////////////////////////////////////////////
# TESTS — EzXlError base class
# ///////////////////////////////////////////////////////////////


@pytest.mark.unit
def test_should_create_ezxl_error_with_message() -> None:
    """Verify that EzXlError can be raised with a plain string message."""
    error = EzXlError("something went wrong")
    assert str(error) == "something went wrong"


@pytest.mark.unit
def test_should_create_ezxl_error_with_cause() -> None:
    """Verify that the optional ``cause`` argument is stored on the instance."""
    original = ValueError("root cause")
    error = EzXlError("wrapper message", cause=original)
    assert error.cause is original


@pytest.mark.unit
def test_should_inherit_from_exception() -> None:
    """Verify that EzXlError is a subclass of the built-in Exception."""
    assert issubclass(EzXlError, Exception)


# ///////////////////////////////////////////////////////////////
# TESTS — COM availability errors
# ///////////////////////////////////////////////////////////////


@pytest.mark.unit
def test_should_create_excel_not_available_error() -> None:
    """Verify ExcelNotAvailableError is an EzXlError with the expected message."""
    error = ExcelNotAvailableError("no Excel instance found")
    assert isinstance(error, EzXlError)
    assert str(error) == "no Excel instance found"


@pytest.mark.unit
def test_should_create_com_operation_error() -> None:
    """Verify COMOperationError is an EzXlError with the expected message."""
    error = COMOperationError("COM call failed")
    assert isinstance(error, EzXlError)
    assert str(error) == "COM call failed"


# ///////////////////////////////////////////////////////////////
# TESTS — GUI operation errors
# ///////////////////////////////////////////////////////////////


@pytest.mark.unit
def test_should_create_gui_operation_error() -> None:
    """Verify GUIOperationError is an EzXlError with the expected message."""
    error = GUIOperationError("ribbon command failed")
    assert isinstance(error, EzXlError)
    assert str(error) == "ribbon command failed"


# ///////////////////////////////////////////////////////////////
# TESTS — Navigation errors
# ///////////////////////////////////////////////////////////////


@pytest.mark.unit
def test_should_create_workbook_not_found_error() -> None:
    """Verify WorkbookNotFoundError is an EzXlError with the expected message."""
    error = WorkbookNotFoundError("no workbook named 'report.xlsx'")
    assert isinstance(error, EzXlError)
    assert str(error) == "no workbook named 'report.xlsx'"


@pytest.mark.unit
def test_should_create_sheet_not_found_error() -> None:
    """Verify SheetNotFoundError is an EzXlError with the expected message."""
    error = SheetNotFoundError("no sheet named 'Summary'")
    assert isinstance(error, EzXlError)
    assert str(error) == "no sheet named 'Summary'"


# ///////////////////////////////////////////////////////////////
# TESTS — Formatter errors
# ///////////////////////////////////////////////////////////////


@pytest.mark.unit
def test_should_create_formatter_error() -> None:
    """Verify FormatterError is an EzXlError with the expected message."""
    error = FormatterError("invalid cell reference")
    assert isinstance(error, EzXlError)
    assert str(error) == "invalid cell reference"


# ///////////////////////////////////////////////////////////////
# TESTS — Exception chaining
# ///////////////////////////////////////////////////////////////


@pytest.mark.unit
def test_should_preserve_cause_in_from_exc_chain() -> None:
    """Verify that ``cause`` is wired to ``__cause__`` for PEP 3134 chaining.

    When the ``cause`` parameter is supplied, ``__cause__`` must be set
    so that ``raise ... from ...`` displays the full chain in tracebacks.
    """
    root = RuntimeError("low-level failure")
    wrapped = EzXlError("high-level failure", cause=root)
    assert wrapped.__cause__ is root


@pytest.mark.unit
def test_should_not_set_dunder_cause_when_cause_is_none() -> None:
    """Verify that ``__cause__`` is not set when no cause is supplied."""
    error = EzXlError("standalone error")
    # __cause__ defaults to None when no explicit cause is given.
    assert error.__cause__ is None
