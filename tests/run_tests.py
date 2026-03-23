#!/usr/bin/env python
# ///////////////////////////////////////////////////////////////
# RUN_TESTS - Test runner script
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
Test runner script for EzXl.

Provides a convenient CLI wrapper around pytest for executing different
types of tests (unit, integration) with various configurations.

Supports:
    - Running specific test types or all tests
    - Coverage reporting
    - Verbose output
    - Parallel execution via pytest-xdist
    - Marker-based filtering
    - Fast mode (excluding slow tests)
    - Excel mode (including live-Excel tests, excluded by default)

Example:
    python run_tests.py --type unit --verbose --coverage
    python run_tests.py --type all --parallel
    python run_tests.py --marker slow --fast
    python run_tests.py --type all --excel
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
import argparse
import logging
import subprocess
import sys
from pathlib import Path

# ///////////////////////////////////////////////////////////////
# CONSTANTS
# ///////////////////////////////////////////////////////////////

logger = logging.getLogger(__name__)

# ///////////////////////////////////////////////////////////////
# HELPER FUNCTIONS
# ///////////////////////////////////////////////////////////////


def run_command(cmd: list[str], description: str) -> bool:
    """Execute a command and stream its output in real-time.

    Args:
        cmd: Command and its arguments as a list of strings.
        description: Human-readable description of the operation,
            printed as a header before the command runs.

    Returns:
        bool: ``True`` if the command exited with code 0, ``False``
            otherwise.
    """
    logger.info("\n%s", "=" * 60)
    logger.info("%s", description)
    logger.info("%s", "=" * 60)
    try:
        result = subprocess.run(cmd, check=False)  # noqa: S603
        return result.returncode == 0
    except KeyboardInterrupt:
        logger.warning("Interrupted by user (Ctrl+C)")
        return False
    except Exception as e:
        logger.exception("Execution error: %s", e)
        return False


# ///////////////////////////////////////////////////////////////
# MAIN
# ///////////////////////////////////////////////////////////////


def main() -> None:
    """Entry point for the test runner.

    Parses CLI arguments, constructs the appropriate pytest invocation,
    and exits with code 0 on success or 1 on failure.

    The ``--excel`` flag opts in to tests marked
    ``@pytest.mark.excel``.  Without it those tests are excluded via
    ``-m 'not excel'`` so that CI runs never require a live Excel
    installation.

    Exit codes:
        0: All selected tests passed.
        1: One or more tests failed, or ``pyproject.toml`` was not found.
    """
    logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")

    parser = argparse.ArgumentParser(
        description="Test runner for EzXl with flexible configuration"
    )
    parser.add_argument(
        "--type",
        choices=["unit", "integration", "all"],
        default="unit",
        help="Type of tests to run (default: unit)",
    )
    parser.add_argument(
        "--coverage",
        action="store_true",
        help="Generate a coverage report",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable verbose pytest output",
    )
    parser.add_argument(
        "--fast",
        action="store_true",
        help="Exclude tests marked @pytest.mark.slow",
    )
    parser.add_argument(
        "--parallel",
        action="store_true",
        help="Run tests in parallel using pytest-xdist",
    )
    parser.add_argument(
        "--marker",
        type=str,
        help="Run only tests matching this marker expression",
    )
    parser.add_argument(
        "--excel",
        action="store_true",
        help=(
            "Include tests marked @pytest.mark.excel "
            "(requires a live Excel installation). "
            "By default these tests are excluded."
        ),
    )
    args = parser.parse_args()

    if not Path("pyproject.toml").exists():
        logger.error("pyproject.toml not found. Run this script from the project root.")
        sys.exit(1)

    # Build pytest command
    cmd_parts = [sys.executable, "-m", "pytest"]

    if args.verbose:
        cmd_parts.append("-v")

    # Marker expression — precedence: --marker > --fast > --excel default filter.
    if args.marker:
        cmd_parts.extend(["-m", args.marker])
    elif args.fast and not args.excel:
        cmd_parts.extend(["-m", "not slow and not excel"])
    elif args.fast and args.excel:
        cmd_parts.extend(["-m", "not slow"])
    elif not args.excel:
        # Default: exclude excel tests so CI never needs a live Excel.
        cmd_parts.extend(["-m", "not excel"])
    # When --excel is set without --fast: no marker filter — run everything.

    if args.parallel:
        cmd_parts.extend(["-n", "auto"])

    if args.type == "unit":
        cmd_parts.append("tests/unit/")
    elif args.type == "integration":
        cmd_parts.append("tests/integration/")
    else:
        cmd_parts.append("tests/")

    if args.coverage:
        cmd_parts.extend(
            [
                "--cov=src/ezxl",
                "--cov-report=term-missing",
                "--cov-report=html:htmlcov",
            ]
        )

    success = run_command(cmd_parts, f"Running {args.type} tests for EzXl")

    if success:
        logger.info("All tests passed successfully")
        if args.coverage:
            logger.info("Coverage report generated in htmlcov/")
            logger.info("Open htmlcov/index.html in your browser")
    else:
        logger.error("Tests failed")
        sys.exit(1)


if __name__ == "__main__":
    main()
