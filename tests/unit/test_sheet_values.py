# ///////////////////////////////////////////////////////////////
# EZXL - Cell value normalisation unit tests
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""Unit tests for ``core/_sheet._normalise_cell_value``.

The function is pure: it translates what COM hands back into plain Python
types. No Excel process and no COM call is involved, so these run wherever
pywin32 imports.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
import sys
from datetime import datetime, timedelta, timezone

# Third-party imports
import pytest

pytestmark = pytest.mark.skipif(
    sys.platform != "win32", reason="core._sheet imports pywintypes (Windows-only)"
)

# ///////////////////////////////////////////////////////////////
# TESTS — Excel error values
# ///////////////////////////////////////////////////////////////


@pytest.mark.parametrize(
    ("error_value", "label"),
    [
        (-2146826288, "#NULL!"),
        (-2146826281, "#DIV/0!"),
        (-2146826273, "#VALUE!"),
        (-2146826265, "#REF!"),
        (-2146826259, "#NAME?"),
        (-2146826252, "#NUM!"),
        (-2146826246, "#N/A"),
        (-2146826245, "#GETTING_DATA"),
    ],
)
def test_should_return_none_when_cell_holds_an_excel_error(
    error_value: int, label: str
) -> None:
    """Verify every XlCVError code is recognised.

    Regression guard: the previous check tested ``isinstance(value,
    pywintypes.error)``, an exception class a cell value can never be an
    instance of. ``=1/0`` leaked through as the raw integer -2146826281, so a
    summed column silently came out around minus two billion.
    """
    from ezxl.core._sheet import _normalise_cell_value

    assert _normalise_cell_value(error_value) is None, label


def test_should_pass_through_when_value_is_an_ordinary_negative_number() -> None:
    """A real negative number must not be mistaken for an error code."""
    from ezxl.core._sheet import _normalise_cell_value

    assert _normalise_cell_value(-2146826282) == -2146826282
    assert _normalise_cell_value(-42) == -42
    assert _normalise_cell_value(-1.5) == -1.5


def test_should_pass_through_when_value_is_plain_data() -> None:
    """Strings, numbers, booleans and None are returned unchanged."""
    from ezxl.core._sheet import _normalise_cell_value

    assert _normalise_cell_value(None) is None
    assert _normalise_cell_value("text") == "text"
    assert _normalise_cell_value(42.0) == 42.0
    assert _normalise_cell_value(True) is True


# ///////////////////////////////////////////////////////////////
# TESTS — COM date conversion
# ///////////////////////////////////////////////////////////////


def test_should_preserve_microseconds_when_converting_a_com_date() -> None:
    """Verify the conversion keeps the sub-second part.

    Regression guard: the conversion rebuilt the value from year..second only,
    truncating microseconds on timestamped data.
    """
    from ezxl.core._sheet import _normalise_cell_value
    from ezxl.utils._pywintypes_compat import COM_TIME_TYPE

    result = _normalise_cell_value(COM_TIME_TYPE(2026, 10, 17, 14, 30, 45, 123456))

    assert isinstance(result, datetime)
    assert result.microsecond == 123456
    assert result == datetime(2026, 10, 17, 14, 30, 45, 123456)


def test_should_return_a_naive_datetime_when_com_date_carries_a_timezone() -> None:
    """An Excel date has no zone, so the COM-supplied one must be dropped.

    Real Excel hands back pywintypes.datetime carrying the machine's
    TimeZoneInfo. Keeping it would make the result tz-aware and raise
    "can't compare offset-naive and offset-aware datetimes" the moment a
    caller compared a cell date with datetime.now().
    """
    from ezxl.core._sheet import _normalise_cell_value
    from ezxl.utils._pywintypes_compat import COM_TIME_TYPE

    tz = timezone(timedelta(hours=2))
    result = _normalise_cell_value(COM_TIME_TYPE(2026, 10, 17, 14, 30, 45, tzinfo=tz))

    assert result.tzinfo is None
    # The wall clock the cell displays is preserved, unshifted.
    assert result == datetime(2026, 10, 17, 14, 30, 45)
    # And it still compares against a naive datetime, as it always did.
    assert result > datetime(2020, 1, 1)


def test_should_return_a_plain_datetime_when_given_a_com_date() -> None:
    """The COM subclass must be normalised away, not merely passed through."""
    from ezxl.core._sheet import _normalise_cell_value
    from ezxl.utils._pywintypes_compat import COM_TIME_TYPE

    result = _normalise_cell_value(COM_TIME_TYPE(2026, 1, 2, 3, 4, 5))

    assert type(result) is datetime
    assert result == datetime(2026, 1, 2, 3, 4, 5)
