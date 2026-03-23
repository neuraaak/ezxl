# ///////////////////////////////////////////////////////////////
# exceptions - EzXl exception hierarchy
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
EzXl exception hierarchy.

All public exceptions raised by the EzXl library. Consumer libraries should
catch ``EzXlError`` as the base type, or specific subclasses for fine-grained
handling.

No exception defined here carries raw COM error objects in its public
interface — callers must not depend on pywin32 internals.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


class EzXlError(Exception):
    """Base exception for all EzXl errors.

    All exceptions raised by the EzXl library inherit from this class.
    Catching ``EzXlError`` is sufficient to handle any EzXl-originated
    failure without importing subclasses.

    Args:
        message: Human-readable description of the error.
        cause: Original exception that triggered this error, if any.

    Example:
        >>> try:
        ...     raise EzXlError("something went wrong")
        ... except EzXlError as e:
        ...     print(e)
        something went wrong
    """

    def __init__(self, message: str, cause: BaseException | None = None) -> None:
        super().__init__(message)
        self.cause = cause
        if cause is not None:
            self.__cause__ = cause


# ///////////////////////////////////////////////////////////////
# COM AVAILABILITY ERRORS
# ///////////////////////////////////////////////////////////////


class ExcelNotAvailableError(EzXlError):
    """Raised when Excel is not open or the COM server is unreachable.

    Typically thrown when ``win32com.client.Dispatch`` or
    ``win32com.client.GetActiveObject`` fails because no Excel process is
    running, or because the COM registration is broken.

    Args:
        message: Human-readable description of the error.
        cause: Original exception that triggered this error, if any.

    Example:
        >>> raise ExcelNotAvailableError(
        ...     "No running Excel instance found", cause=original_err
        ... )
    """

    def __init__(self, message: str, cause: BaseException | None = None) -> None:
        super().__init__(message, cause)


class ExcelSessionLostError(EzXlError):
    """Raised when an established COM connection is lost mid-operation.

    This can happen when the user closes Excel while an automation session
    is active, or when Excel crashes. Unlike ``ExcelNotAvailableError``,
    this implies a previously valid connection existed.

    Args:
        message: Human-readable description of the error.
        cause: Original exception that triggered this error, if any.
    """

    def __init__(self, message: str, cause: BaseException | None = None) -> None:
        super().__init__(message, cause)


class ExcelThreadViolationError(EzXlError):
    """Raised when a COM call is attempted from the wrong thread.

    Excel COM operates under the Single-Threaded Apartment (STA) model.
    All COM calls must originate from the thread that created the
    ``ExcelApp`` instance. This exception is raised proactively before
    the COM call reaches the dispatcher to give a clear diagnostic.

    Args:
        message: Human-readable description of the error.
        cause: Original exception that triggered this error, if any.
    """

    def __init__(self, message: str, cause: BaseException | None = None) -> None:
        super().__init__(message, cause)


# ///////////////////////////////////////////////////////////////
# NAVIGATION ERRORS
# ///////////////////////////////////////////////////////////////


class WorkbookNotFoundError(EzXlError):
    """Raised when a workbook cannot be found by name in the Excel session.

    Args:
        message: Human-readable description of the error.
        cause: Original exception that triggered this error, if any.

    Example:
        >>> raise WorkbookNotFoundError("No workbook named 'report.xlsx'")
    """

    def __init__(self, message: str, cause: BaseException | None = None) -> None:
        super().__init__(message, cause)


class SheetNotFoundError(EzXlError):
    """Raised when a worksheet cannot be found by name in a workbook.

    Args:
        message: Human-readable description of the error.
        cause: Original exception that triggered this error, if any.

    Example:
        >>> raise SheetNotFoundError("No sheet named 'Summary' in 'report.xlsx'")
    """

    def __init__(self, message: str, cause: BaseException | None = None) -> None:
        super().__init__(message, cause)


# ///////////////////////////////////////////////////////////////
# COM OPERATION ERRORS
# ///////////////////////////////////////////////////////////////


class COMOperationError(EzXlError):
    """Raised for unclassified COM errors that do not map to a specific subclass.

    This is the catch-all wrapper for ``pywintypes.com_error`` exceptions.
    If a COM error can be identified as a lost session or unavailability
    issue, the more specific subclass should be used instead.

    Args:
        message: Human-readable description of the error.
        cause: Original exception that triggered this error, if any.
    """

    def __init__(self, message: str, cause: BaseException | None = None) -> None:
        super().__init__(message, cause)


# ///////////////////////////////////////////////////////////////
# GUI OPERATION ERRORS
# ///////////////////////////////////////////////////////////////


class GUIOperationError(EzXlError):
    """Raised when a GUI-level COM operation fails for ribbon, menu, or dialog.

    Distinct from ``COMOperationError`` to allow consumer code to
    differentiate between generic COM failures and failures that occur
    specifically within GUI interaction surfaces (ribbon, CommandBars,
    file dialogs, message boxes).

    Args:
        message: Human-readable description of the error.
        cause: Original exception that triggered this error, if any.

    Example:
        >>> raise GUIOperationError(
        ...     "Failed to execute ribbon command 'FileSave'", cause=exc
        ... )
    """

    def __init__(self, message: str, cause: BaseException | None = None) -> None:
        super().__init__(message, cause)


# ///////////////////////////////////////////////////////////////
# FORMATTER ERRORS
# ///////////////////////////////////////////////////////////////


class FormatterError(EzXlError):
    """Raised when an openpyxl-based formatting operation fails.

    This exception covers errors that occur during closed-file formatting
    via ``ExcelFormatter``, such as invalid cell references, unsupported
    style properties, or file I/O failures.

    Args:
        message: Human-readable description of the error.
        cause: Original exception that triggered this error, if any.
    """

    def __init__(self, message: str, cause: BaseException | None = None) -> None:
        super().__init__(message, cause)
