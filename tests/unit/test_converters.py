# ///////////////////////////////////////////////////////////////
# test_converters - EzXl I/O converter tests
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
Unit tests for the format conversion utilities in ``ezxl.io._converters``.

All tests operate on closed files using the ``sample_xlsx`` and
``sample_csv`` fixtures defined in ``conftest.py``.  No running Excel
process or COM connection is required.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
from pathlib import Path

# Third-party imports
import polars as pl
import pytest

# Local imports
from ezxl.io._converters import csv_to_xlsx, read_csv, read_excel, xlsx_to_csv

# ///////////////////////////////////////////////////////////////
# TESTS — read_excel
# ///////////////////////////////////////////////////////////////


@pytest.mark.unit
def test_should_read_excel_returns_dataframe(sample_xlsx: Path) -> None:
    """Verify that read_excel returns a non-empty polars DataFrame.

    The sample workbook has 2 data rows and 3 columns.

    Args:
        sample_xlsx: Path to the sample ``.xlsx`` fixture.
    """
    df = read_excel(sample_xlsx)
    assert isinstance(df, pl.DataFrame)
    assert df.shape == (2, 3)


@pytest.mark.unit
def test_should_read_excel_raises_on_missing_file(temp_dir: Path) -> None:
    """Verify that read_excel raises FileNotFoundError for a non-existent path.

    Args:
        temp_dir: Temporary directory fixture (used to build a non-existent path).
    """
    missing = temp_dir / "does_not_exist.xlsx"
    with pytest.raises(FileNotFoundError):
        read_excel(missing)


# ///////////////////////////////////////////////////////////////
# TESTS — read_csv
# ///////////////////////////////////////////////////////////////


@pytest.mark.unit
def test_should_read_csv_returns_dataframe(sample_csv: Path) -> None:
    """Verify that read_csv returns a non-empty polars DataFrame.

    The sample CSV has 2 data rows and 3 columns.

    Args:
        sample_csv: Path to the sample ``.csv`` fixture.
    """
    df = read_csv(sample_csv)
    assert isinstance(df, pl.DataFrame)
    assert df.shape == (2, 3)


@pytest.mark.unit
def test_should_read_csv_raises_on_missing_file(temp_dir: Path) -> None:
    """Verify that read_csv raises FileNotFoundError for a non-existent path.

    Args:
        temp_dir: Temporary directory fixture (used to build a non-existent path).
    """
    missing = temp_dir / "does_not_exist.csv"
    with pytest.raises(FileNotFoundError):
        read_csv(missing)


# ///////////////////////////////////////////////////////////////
# TESTS — xlsx_to_csv
# ///////////////////////////////////////////////////////////////


@pytest.mark.unit
def test_should_xlsx_to_csv_creates_output_file(
    sample_xlsx: Path, temp_dir: Path
) -> None:
    """Verify that xlsx_to_csv writes a CSV file to the destination path.

    Checks that the output file exists and is non-empty after the call.

    Args:
        sample_xlsx: Path to the sample ``.xlsx`` fixture.
        temp_dir: Temporary directory fixture (used as the output location).
    """
    dest = temp_dir / "output.csv"
    xlsx_to_csv(sample_xlsx, dest)
    assert dest.exists()
    assert dest.stat().st_size > 0


# ///////////////////////////////////////////////////////////////
# TESTS — csv_to_xlsx
# ///////////////////////////////////////////////////////////////


@pytest.mark.unit
def test_should_csv_to_xlsx_creates_output_file(
    sample_csv: Path, temp_dir: Path
) -> None:
    """Verify that csv_to_xlsx writes an Excel file to the destination path.

    Checks that the output file exists and is non-empty after the call.

    Args:
        sample_csv: Path to the sample ``.csv`` fixture.
        temp_dir: Temporary directory fixture (used as the output location).
    """
    dest = temp_dir / "output.xlsx"
    csv_to_xlsx(sample_csv, dest)
    assert dest.exists()
    assert dest.stat().st_size > 0
