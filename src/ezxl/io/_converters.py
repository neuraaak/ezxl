# ///////////////////////////////////////////////////////////////
# _converters - Format conversion utilities
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""Format conversion utilities for Excel and CSV files.

Provides read and export paths backed by polars for high-throughput
data processing on closed files (no running Excel process required):

- ``read_excel``: read an ``.xlsx`` file into a polars DataFrame.
- ``read_csv``: read a ``.csv`` file into a polars DataFrame.
- ``xlsx_to_csv``: convert an Excel sheet to CSV via polars.
- ``csv_to_xlsx``: convert a CSV file to an ``.xlsx`` file via polars.
- ``read_sheet``: compatibility shim — returns ``list[list[Any]]`` for
  callers that expect the legacy row-major format.

All functions operate on **closed** files and require no running Excel
process.  polars delegates Excel I/O to ``fastexcel`` (a Rust-based
engine bundled with polars extras) which provides performance comparable
to the former ``python-calamine`` path.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
from pathlib import Path
from typing import Any

# Third-party imports
import polars as pl
from ezplog.lib_mode import get_logger, get_printer

# ///////////////////////////////////////////////////////////////
# CONSTANTS
# ///////////////////////////////////////////////////////////////

logger = get_logger(__name__)
printer = get_printer()

# ///////////////////////////////////////////////////////////////
# FUNCTIONS
# ///////////////////////////////////////////////////////////////


def read_excel(
    source: str | Path,
    sheet: str | None = None,
) -> pl.DataFrame:
    """Read an Excel workbook sheet into a polars DataFrame.

    Delegates to ``polars.read_excel`` which uses ``fastexcel`` (Rust)
    under the hood.  No running Excel process is required.

    Args:
        source: Path to the source ``.xlsx`` / ``.xlsm`` file.
        sheet: Worksheet name to read.  Pass ``None`` to read the first
            sheet (polars default when ``sheet_name`` is omitted).

    Returns:
        pl.DataFrame: Contents of the requested sheet as a polars
        DataFrame, with the first row used as column headers.

    Raises:
        FileNotFoundError: If ``source`` does not exist.
        ImportError: If polars (or its ``fastexcel`` extra) is not
            installed.

    Example:
        >>> df = read_excel("report.xlsx", sheet="Data")
        >>> print(df.head())
    """
    source_path = Path(source).resolve()

    if not source_path.exists():
        raise FileNotFoundError(f"Source file not found: {source_path}")

    logger.debug("read_excel: %s (sheet=%r)", source_path, sheet)

    df: pl.DataFrame = pl.read_excel(source_path, sheet_name=sheet)

    logger.debug("read_excel: read %d rows from '%s'.", len(df), source_path)
    return df


def read_csv(
    source: str | Path,
    separator: str = ",",
    encoding: str = "utf-8",
) -> pl.DataFrame:
    """Read a CSV file into a polars DataFrame.

    Args:
        source: Path to the source ``.csv`` file.
        separator: Column delimiter character.  Defaults to ``","``
            (standard CSV).  Use ``"\\t"`` for TSV files.
        encoding: File encoding passed through to polars.  Defaults to
            ``"utf-8"``.

    Returns:
        pl.DataFrame: Parsed contents of the CSV file.

    Raises:
        FileNotFoundError: If ``source`` does not exist.

    Example:
        >>> df = read_csv("transactions.csv", separator=";")
        >>> print(df.schema)
    """
    source_path = Path(source).resolve()

    if not source_path.exists():
        raise FileNotFoundError(f"Source file not found: {source_path}")

    logger.debug("read_csv: %s (sep=%r, enc=%r)", source_path, separator, encoding)

    df: pl.DataFrame = pl.read_csv(
        source_path,
        separator=separator,
        encoding=encoding,
    )

    logger.debug("read_csv: read %d rows from '%s'.", len(df), source_path)
    return df


def xlsx_to_csv(
    source: str | Path,
    dest: str | Path,
    sheet: str | None = None,
    separator: str = ",",
) -> None:
    """Convert an Excel workbook sheet to a CSV file using polars.

    Supersedes both the former ``xlsx_to_csv`` (openpyxl) and
    ``xlsx_to_csv_fast`` (python-calamine) functions.  polars uses
    ``fastexcel`` (Rust) for the read step, providing the same
    high-throughput characteristics as the former fast path.

    Args:
        source: Path to the source ``.xlsx`` / ``.xlsm`` file.
        dest: Destination ``.csv`` file path.  Parent directories must
            exist.
        sheet: Worksheet name to export.  Pass ``None`` to use the
            first sheet.
        separator: Column delimiter for the CSV output.  Defaults to
            ``","`` (standard CSV).

    Raises:
        FileNotFoundError: If ``source`` does not exist.

    Example:
        >>> xlsx_to_csv("data.xlsx", "data.csv", sheet="Transactions")
        >>> xlsx_to_csv("data.xlsx", "data.tsv", separator="\\t")
    """
    dest_path = Path(dest).resolve()

    logger.debug(
        "xlsx_to_csv: %s → %s (sheet=%r, sep=%r)",
        Path(source).resolve(),
        dest_path,
        sheet,
        separator,
    )

    df = read_excel(source, sheet=sheet)
    df.write_csv(dest_path, separator=separator)

    logger.debug("xlsx_to_csv: completed — wrote %s", dest_path)
    printer.success(f"xlsx_to_csv: conversion complete — {dest_path}")


def csv_to_xlsx(
    source: str | Path,
    dest: str | Path,
    sheet_name: str = "Sheet1",
) -> None:
    """Convert a CSV file to an Excel workbook using polars.

    Reads the CSV with polars and writes it as an ``.xlsx`` file.
    polars delegates the Excel write step to ``xlsxwriter`` or
    ``openpyxl`` depending on which is installed; no additional
    configuration is required.

    Args:
        source: Path to the source ``.csv`` file.
        dest: Destination ``.xlsx`` file path.  Parent directories must
            exist.
        sheet_name: Name of the worksheet to create in the output
            workbook.  Defaults to ``"Sheet1"``.

    Raises:
        FileNotFoundError: If ``source`` does not exist.

    Example:
        >>> csv_to_xlsx("transactions.csv", "transactions.xlsx", sheet_name="Data")
    """
    dest_path = Path(dest).resolve()

    logger.debug(
        "csv_to_xlsx: %s → %s (sheet=%r)",
        Path(source).resolve(),
        dest_path,
        sheet_name,
    )

    df = read_csv(source)
    df.write_excel(dest_path, worksheet=sheet_name)

    logger.debug("csv_to_xlsx: completed — wrote %s", dest_path)
    printer.success(f"csv_to_xlsx: conversion complete — {dest_path}")


def read_sheet(
    source: str | Path,
    sheet: str | None = None,
) -> list[list[Any]]:
    """Read a worksheet into a row-major list of lists (compatibility shim).

    Wraps ``read_excel`` and converts the resulting polars DataFrame to
    a ``list[list[Any]]`` via ``DataFrame.rows()``.  The first row
    contains the column headers as extracted by polars.

    This function exists for backwards compatibility with callers that
    pre-date the polars migration.  New code should use ``read_excel``
    directly to benefit from the full polars API.

    Args:
        source: Path to the source ``.xlsx`` / ``.xlsm`` file.
        sheet: Worksheet name to read.  Pass ``None`` to use the first
            sheet.

    Returns:
        list[list[Any]]: Row-major 2D list of cell values.  The first
        row contains column headers; subsequent rows contain data
        values.  Empty cells are represented as ``None``.

    Raises:
        FileNotFoundError: If ``source`` does not exist.

    Example:
        >>> data = read_sheet("report.xlsx", sheet="Data")
        >>> headers = data[0]
        >>> rows = data[1:]
    """
    logger.debug("read_sheet: %s (sheet=%r) — delegating to read_excel", source, sheet)

    df = read_excel(source, sheet=sheet)

    # Prepend column names as the first row to preserve the legacy contract
    # where callers expected headers in row 0.
    header_row: list[Any] = list(df.columns)
    data_rows: list[list[Any]] = [list(row) for row in df.rows()]

    result: list[list[Any]] = [header_row, *data_rows]

    logger.debug("read_sheet: returning %d rows (incl. header).", len(result))
    return result
