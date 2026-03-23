# ///////////////////////////////////////////////////////////////
# EzXl - Pytest Configuration and Fixtures
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
Pytest configuration and shared fixtures for EzXl tests.

This module provides fixtures for:
- Temporary files and directories
- Sample Excel and CSV files for I/O tests
- Custom pytest markers
- Windows COM teardown hooks for test isolation
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
import csv
import gc
import shutil
import sys
import tempfile
from collections.abc import Generator
from pathlib import Path

# Third-party imports
import openpyxl
import pytest

# ///////////////////////////////////////////////////////////////
# FIXTURES
# ///////////////////////////////////////////////////////////////


@pytest.fixture
def temp_dir() -> Generator[Path, None, None]:
    """Create a temporary directory for test files.

    The directory is removed after the test completes. ``gc.collect()``
    is called before removal to release any Python-level file references
    that could prevent deletion on Windows.

    Yields:
        Path: Path to the freshly created temporary directory.
    """
    tmpdir = Path(tempfile.mkdtemp())
    yield tmpdir
    # gc.collect() releases Python-level file handle references before
    # rmtree so that Windows does not raise PermissionError on cleanup.
    gc.collect()
    shutil.rmtree(tmpdir, ignore_errors=True)


@pytest.fixture
def temp_file(temp_dir: Path) -> Path:
    """Return a path to a non-existent file inside the temporary directory.

    The file is not created — callers are responsible for writing to it.
    The parent temporary directory is cleaned up by the ``temp_dir``
    fixture after the test completes.

    Args:
        temp_dir: Temporary directory fixture.

    Returns:
        Path: Path to a temporary file (not yet created on disk).
    """
    return temp_dir / "test_file.tmp"


@pytest.fixture
def sample_xlsx(temp_dir: Path) -> Path:
    """Create a minimal ``.xlsx`` workbook for use in converter tests.

    Writes a workbook with one sheet named ``"Sheet1"`` containing
    three rows and three columns of simple string/integer values.
    The first row acts as a header row.

    Layout::

        | col_a | col_b | col_c |
        |-------|-------|-------|
        |   1   |   2   |   3   |
        |   4   |   5   |   6   |

    Args:
        temp_dir: Temporary directory fixture.

    Returns:
        Path: Absolute path to the created ``.xlsx`` file.
    """
    path = temp_dir / "sample.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None, "openpyxl.Workbook().active must not be None"
    ws.title = "Sheet1"
    ws.append(["col_a", "col_b", "col_c"])
    ws.append([1, 2, 3])
    ws.append([4, 5, 6])
    wb.save(path)
    return path


@pytest.fixture
def sample_csv(temp_dir: Path) -> Path:
    """Create a minimal ``.csv`` file for use in converter tests.

    Writes a CSV file with three columns and three rows (one header row
    and two data rows) using standard comma separation and UTF-8 encoding.

    Layout::

        col_a,col_b,col_c
        1,2,3
        4,5,6

    Args:
        temp_dir: Temporary directory fixture.

    Returns:
        Path: Absolute path to the created ``.csv`` file.
    """
    path = temp_dir / "sample.csv"
    with path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(["col_a", "col_b", "col_c"])
        writer.writerow([1, 2, 3])
        writer.writerow([4, 5, 6])
    return path


# ///////////////////////////////////////////////////////////////
# PYTEST HOOKS
# ///////////////////////////////////////////////////////////////


def pytest_runtest_teardown(item, nextitem) -> None:  # noqa: ARG001
    """Force a GC cycle after every test teardown on Windows.

    ``gc.collect()`` releases Python-level file handle references
    immediately after each test so that temporary directory cleanup
    does not encounter locked files left over by COM or openpyxl.
    """
    if sys.platform == "win32":
        gc.collect()


@pytest.hookimpl(tryfirst=True)
def pytest_runtest_makereport(item, call):  # noqa: ARG001  # type: ignore[no-untyped-def]
    """Suppress Windows file-lock errors recorded during teardown.

    On Windows, COM and openpyxl can hold file handles slightly past
    the end of a test. If pytest attempts to clean up a temporary
    directory while a handle is still open it records a
    ``PermissionError`` or ``NotADirectoryError``. These do not indicate
    test failures and are suppressed here so they do not pollute the
    test report.
    """
    if sys.platform == "win32" and call.when == "teardown" and call.excinfo:
        exc_type = call.excinfo.type
        if exc_type in (NotADirectoryError, PermissionError, OSError):
            exc_value = str(call.excinfo.value).lower()
            if any(
                keyword in exc_value
                for keyword in [
                    "winerror 32",
                    "winerror 267",
                    "utilisé par un autre processus",
                    "used by another process",
                    "nom de répertoire non valide",
                ]
            ):
                call.excinfo = None


def pytest_runtest_logreport(report) -> None:
    """Suppress Windows file-lock teardown errors in the final report.

    Companion to ``pytest_runtest_makereport``: ensures that errors
    already recorded in the report object are also cleared when they
    match known Windows file-locking patterns.
    """
    if sys.platform == "win32" and report.when == "teardown" and report.failed:
        longrepr_str = ""
        if hasattr(report, "longrepr") and report.longrepr:
            longrepr_str = str(report.longrepr).lower()

        if any(
            keyword in longrepr_str
            for keyword in [
                "winerror 32",
                "winerror 267",
                "utilisé par un autre processus",
                "used by another process",
                "nom de répertoire non valide",
                "notadirectoryerror",
                "permissionerror",
            ]
        ):
            report.outcome = "passed"
            report.longrepr = None
            report.sections = []


# ///////////////////////////////////////////////////////////////
# MARKERS
# ///////////////////////////////////////////////////////////////

# Markers are registered in pyproject.toml [tool.pytest.ini_options].
# Available markers:
# - @pytest.mark.excel:       requires a live Excel installation (excluded from CI)
# - @pytest.mark.unit:        pure-Python tests with no external dependencies
# - @pytest.mark.integration: tests that exercise multiple layers together
# - @pytest.mark.slow:        tests that take longer than a few seconds


def pytest_configure(config) -> None:
    """Register custom markers to suppress PytestUnknownMarkWarning.

    Args:
        config: The pytest configuration object.
    """
    config.addinivalue_line(
        "markers",
        "excel: marks tests requiring a live Excel installation "
        "(deselect with '-m not excel')",
    )
    config.addinivalue_line(
        "markers",
        "unit: marks tests as unit tests (no external dependencies)",
    )
    config.addinivalue_line(
        "markers",
        "integration: marks tests as integration tests",
    )
    config.addinivalue_line(
        "markers",
        "slow: marks tests as slow (can be excluded with --fast)",
    )
