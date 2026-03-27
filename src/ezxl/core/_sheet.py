# ///////////////////////////////////////////////////////////////
# _sheet - SheetProxy, CellProxy, RangeProxy COM wrappers
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
SheetProxy, CellProxy, RangeProxy — COM proxies for Excel worksheet objects.

Provide a clean Python interface over COM Worksheet, Range, and Cell
objects. COM date values are normalised to ``datetime``. Cell error values
(``#N/A``, ``#VALUE!``, etc.) are mapped to ``None`` with a WARNING log.

All COM calls are wrapped via ``_com_utils.wrap_com_error``.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
from datetime import datetime
from typing import TYPE_CHECKING, Any

# Third-party imports
import pywintypes
from ezplog.lib_mode import get_logger

# Local imports
from ..exceptions import SheetNotFoundError
from ..utils._com_utils import assert_main_thread, wrap_com_error

if TYPE_CHECKING:
    from ._workbook import WorkbookProxy

# ///////////////////////////////////////////////////////////////
# CONSTANTS
# ///////////////////////////////////////////////////////////////

logger = get_logger(__name__)

# pywintypes ships without complete type stubs — cache the runtime types via
# getattr so that type checkers (ty, pyright) don't flag missing attributes.
_COM_TIME_TYPE: type = pywintypes.TimeType  # type: ignore[attr-defined]
_COM_ERROR_TYPE: type = pywintypes.error  # type: ignore[attr-defined]


def _is_com_date(value: Any) -> bool:
    """Return True if ``value`` is a COM pywintypes.datetime instance."""
    return isinstance(value, _COM_TIME_TYPE)


def _normalise_cell_value(value: Any) -> Any:
    """Convert COM-specific types to plain Python types.

    - ``pywintypes.datetime`` → ``datetime``
    - COM error variants → ``None`` (with a WARNING log)
    - All other types pass through unchanged.

    Args:
        value: The raw value from a COM cell.

    Returns:
        Any: The normalised Python value.
    """
    if value is None:
        return None

    # Normalise COM dates to stdlib datetime.
    if _is_com_date(value):
        try:
            return datetime(
                value.year,
                value.month,
                value.day,
                value.hour,
                value.minute,
                value.second,
            )
        except Exception as exc:
            logger.warning("Failed to convert COM date %r to datetime: %s", value, exc)
            return value

    # Detect COM error values (#N/A, #VALUE!, etc.) — pywintypes wraps
    # these as pywintypes.error instances.
    if isinstance(value, _COM_ERROR_TYPE):
        logger.warning("Cell contains a COM error value (%r); returning None.", value)
        return None

    return value


# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


class SheetProxy:
    """COM proxy for a single Excel Worksheet.

    Instances are created by ``WorkbookProxy.sheet()``. Do not instantiate
    directly.

    Args:
        workbook: The parent ``WorkbookProxy``.
        name: The worksheet name (as shown on the tab).

    Example:
        >>> ws = wb.sheet("Data")
        >>> ws.cell("A1").value = "Hello"
        >>> ws.calculate()
    """

    # ///////////////////////////////////////////////////////////////
    # INIT
    # ///////////////////////////////////////////////////////////////

    def __init__(self, workbook: WorkbookProxy, name: str) -> None:
        self._workbook = workbook
        self._name = name

    # ///////////////////////////////////////////////////////////////
    # PRIVATE HELPERS
    # ///////////////////////////////////////////////////////////////

    def _check_thread(self) -> None:
        """Delegate thread assertion to the parent ExcelApp."""
        assert_main_thread(self._workbook._app._thread_id)

    def _get_ws(self) -> Any:
        """Resolve and return the underlying COM Worksheet object.

        Returns:
            The win32com Worksheet COM object.

        Raises:
            SheetNotFoundError: If the sheet no longer exists.
        """
        try:
            wb = self._workbook._get_wb()
            return wb.Sheets(self._name)
        except Exception as exc:
            raise SheetNotFoundError(
                f"Sheet '{self._name}' is no longer available in "
                f"'{self._workbook.name}'.",
                cause=exc,
            ) from exc

    # ///////////////////////////////////////////////////////////////
    # PROPERTIES
    # ///////////////////////////////////////////////////////////////

    @property
    def name(self) -> str:
        """The worksheet name as shown on the tab.

        Returns:
            str: Sheet name.
        """
        return self._name

    @property
    @wrap_com_error
    def used_range(self) -> RangeProxy:
        """The smallest rectangle that contains all used cells.

        Returns:
            RangeProxy: Proxy over the ``Worksheet.UsedRange`` COM range.

        Raises:
            SheetNotFoundError: If the sheet is no longer available.
        """
        self._check_thread()
        ws = self._get_ws()
        address: str = ws.UsedRange.Address
        logger.debug("UsedRange of '%s': %s", self._name, address)
        return RangeProxy(self, address)

    # ///////////////////////////////////////////////////////////////
    # PUBLIC METHODS
    # ///////////////////////////////////////////////////////////////

    @wrap_com_error
    def cell(self, ref: str) -> CellProxy:
        """Return a proxy for a single cell.

        Args:
            ref: Cell address in A1 notation (e.g. ``"B3"``).

        Returns:
            CellProxy: A proxy bound to the cell.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            SheetNotFoundError: If the sheet is no longer available.

        Example:
            >>> ws.cell("A1").value = 42
        """
        self._check_thread()
        # Validate the sheet exists before constructing the proxy.
        self._get_ws()
        return CellProxy(self, ref)

    @wrap_com_error
    def range(self, ref: str) -> RangeProxy:
        """Return a proxy for a cell range.

        Args:
            ref: Range address in A1 notation (e.g. ``"A1:D10"``).

        Returns:
            RangeProxy: A proxy bound to the range.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            SheetNotFoundError: If the sheet is no longer available.

        Example:
            >>> rng = ws.range("A1:C5")
            >>> data = rng.values
        """
        self._check_thread()
        self._get_ws()
        return RangeProxy(self, ref)

    @wrap_com_error
    def calculate(self) -> None:
        """Trigger recalculation of all formulas on this sheet.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            SheetNotFoundError: If the sheet is no longer available.

        Example:
            >>> ws.calculate()
        """
        self._check_thread()
        logger.debug("Calculating sheet '%s'.", self._name)
        self._get_ws().Calculate()


# ///////////////////////////////////////////////////////////////
# CLASSES — CellProxy
# ///////////////////////////////////////////////////////////////


class CellProxy:
    """COM proxy for a single cell in a worksheet.

    Instances are created by ``SheetProxy.cell()``. Do not instantiate
    directly.

    Args:
        sheet: The parent ``SheetProxy``.
        ref: Cell address in A1 notation (e.g. ``"B3"``).

    Example:
        >>> cell = ws.cell("C7")
        >>> cell.value = 100
        >>> print(cell.formula)
        100
    """

    # ///////////////////////////////////////////////////////////////
    # INIT
    # ///////////////////////////////////////////////////////////////

    def __init__(self, sheet: SheetProxy, ref: str) -> None:
        self._sheet = sheet
        self._ref = ref

    # ///////////////////////////////////////////////////////////////
    # PRIVATE HELPERS
    # ///////////////////////////////////////////////////////////////

    def _get_cell(self) -> Any:
        """Resolve the COM Range object for this single cell.

        Returns:
            The win32com Range COM object (single cell).
        """
        ws = self._sheet._get_ws()
        return ws.Range(self._ref)

    # ///////////////////////////////////////////////////////////////
    # PROPERTIES
    # ///////////////////////////////////////////////////////////////

    @property
    def address(self) -> str:
        """The absolute address of this cell (e.g. ``"$B$3"``).

        Returns:
            str: Cell address in absolute notation.
        """
        return self._ref

    @property
    @wrap_com_error
    def value(self) -> Any:
        """The cell's current value, normalised to a Python type.

        COM date objects are converted to ``datetime``. Error values
        (``#N/A`` etc.) are returned as ``None`` with a WARNING log.

        Returns:
            Any: The cell value, or ``None`` for empty/error cells.
        """
        self._sheet._check_thread()
        raw = self._get_cell().Value
        return _normalise_cell_value(raw)

    @value.setter
    @wrap_com_error
    def value(self, val: Any) -> None:
        """Set the cell's value.

        Args:
            val: The value to write. Accepts any type that Excel can store
                (str, int, float, datetime, bool, None).
        """
        self._sheet._check_thread()
        self._get_cell().Value = val

    @property
    @wrap_com_error
    def formula(self) -> str:
        """The formula string in the cell (e.g. ``"=SUM(A1:A5)"``).

        Returns an empty string for cells with no formula.

        Returns:
            str: Formula string, or empty string if none.
        """
        self._sheet._check_thread()
        return self._get_cell().Formula

    @formula.setter
    @wrap_com_error
    def formula(self, expr: str) -> None:
        """Set the cell's formula.

        Args:
            expr: Formula string (e.g. ``"=SUM(A1:A5)"``). Pass an empty
                string to clear the formula without clearing the value.
        """
        self._sheet._check_thread()
        self._get_cell().Formula = expr


# ///////////////////////////////////////////////////////////////
# CLASSES — RangeProxy
# ///////////////////////////////////////////////////////////////


class RangeProxy:
    """COM proxy for a rectangular range of cells in a worksheet.

    Instances are created by ``SheetProxy.range()`` or accessed via
    ``SheetProxy.used_range``. Do not instantiate directly.

    Args:
        sheet: The parent ``SheetProxy``.
        ref: Range address in A1 notation (e.g. ``"A1:D10"``).

    Example:
        >>> rng = ws.range("A1:C3")
        >>> data = rng.values           # list[list[Any]]
        >>> rng.values = [[1, 2, 3],
        ...               [4, 5, 6],
        ...               [7, 8, 9]]
    """

    # ///////////////////////////////////////////////////////////////
    # INIT
    # ///////////////////////////////////////////////////////////////

    def __init__(self, sheet: SheetProxy, ref: str) -> None:
        self._sheet = sheet
        self._ref = ref

    # ///////////////////////////////////////////////////////////////
    # PRIVATE HELPERS
    # ///////////////////////////////////////////////////////////////

    def _get_range(self) -> Any:
        """Resolve the COM Range object for this range.

        Returns:
            The win32com Range COM object.
        """
        ws = self._sheet._get_ws()
        return ws.Range(self._ref)

    # ///////////////////////////////////////////////////////////////
    # PROPERTIES
    # ///////////////////////////////////////////////////////////////

    @property
    def address(self) -> str:
        """The address of this range (e.g. ``"A1:D10"``).

        Returns:
            str: Range address.
        """
        return self._ref

    @property
    @wrap_com_error
    def values(self) -> list[list[Any]]:
        """All cell values in the range as a list of rows.

        Single-row or single-column ranges are normalised to a 2D
        list-of-lists for a consistent return type. COM dates are converted
        to ``datetime``; error cells return ``None``.

        Returns:
            list[list[Any]]: Row-major 2D list of cell values.

        Example:
            >>> data = ws.range("A1:C3").values
            >>> data[0][0]  # row 1, col A
        """
        self._sheet._check_thread()
        rng = self._get_range()

        # COM returns a 2-tuple-of-tuples for multi-cell ranges, a scalar
        # for single cells. Normalise everything to list[list[Any]].
        raw = rng.Value

        if raw is None:
            return [[None]]

        # Single cell: raw is a scalar.
        if not isinstance(raw, (tuple, list)):
            return [[_normalise_cell_value(raw)]]

        # Multi-row range: raw is a tuple of row-tuples.
        if raw and isinstance(raw[0], (tuple, list)):
            return [[_normalise_cell_value(cell) for cell in row] for row in raw]

        # Single-row range: raw is a flat tuple.
        return [[_normalise_cell_value(cell) for cell in raw]]

    @values.setter
    @wrap_com_error
    def values(self, data: list[list[Any]]) -> None:
        """Write a 2D list of values into the range.

        The data dimensions must match the range dimensions; Excel will
        raise a COM error if they do not.

        Args:
            data: Row-major 2D list of values to write.

        Example:
            >>> ws.range("A1:C2").values = [["a", "b", "c"], [1, 2, 3]]
        """
        self._sheet._check_thread()
        self._get_range().Value = data
