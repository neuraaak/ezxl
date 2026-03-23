# ///////////////////////////////////////////////////////////////
# EZXL - CLI Info Command
# Project: ezxl
# ///////////////////////////////////////////////////////////////

"""
CLI command for displaying package information.

This module provides the info command for EzXl.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

# Third-party imports
import click
from rich.panel import Panel
from rich.table import Table
from rich.text import Text

# Local imports
import ezxl

from .._console import console

# ///////////////////////////////////////////////////////////////
# COMMANDS
# ///////////////////////////////////////////////////////////////


@click.command(name="info", help="Display package information")
def info_command() -> None:
    """
    Display package information.

    Show detailed information about the EzXl package including
    version, location, and dependencies.
    """
    try:
        # Package info
        ezxl_version = getattr(ezxl, "__version__", "unknown")
        author = getattr(ezxl, "__author__", "unknown")
        maintainer = getattr(ezxl, "__maintainer__", "unknown")
        description = getattr(ezxl, "__description__", "unknown")
        url = getattr(ezxl, "__url__", "unknown")

        # Package location
        try:
            package_path = (
                Path(ezxl.__file__).parent if hasattr(ezxl, "__file__") else None
            )
        except (AttributeError, TypeError, OSError):
            package_path = None

        # Build info text
        text = Text()
        text.append("Package Information\n", style="bold bright_blue")
        text.append("=" * 50 + "\n\n", style="dim")

        # Version
        text.append("Version: ", style="bold")
        text.append(f"{ezxl_version}\n", style="white")

        # Author
        text.append("Author: ", style="bold")
        text.append(f"{author}\n", style="white")

        if maintainer != author:
            text.append("Maintainer: ", style="bold")
            text.append(f"{maintainer}\n", style="white")

        # Description
        text.append("\nDescription:\n", style="bold")
        text.append(f"  {description}\n", style="dim white")

        # URL
        text.append("\nURL: ", style="bold")
        text.append(f"{url}\n", style="cyan")

        # Package location
        if package_path:
            text.append("\nPackage Location: ", style="bold")
            text.append(f"{package_path}\n", style="dim white")

        # Display panel
        panel = Panel(
            text,
            title="[bold bright_blue]EzXl Information[/bold bright_blue]",
            border_style="bright_blue",
            padding=(1, 2),
        )
        console.print(panel)

        # Dependencies table
        try:
            _dep_names = [
                "polars",
                "openpyxl",
                "xlsxwriter",
                "fastexcel",
                "rich",
                "click",
            ]

            deps_table = Table(
                title="Dependencies", show_header=True, header_style="bold blue"
            )
            deps_table.add_column("Package", style="cyan")
            deps_table.add_column("Version", style="green")

            for dep in _dep_names:
                try:
                    dep_version = version(dep)
                except PackageNotFoundError:
                    dep_version = "unknown"
                deps_table.add_row(dep, dep_version)

            console.print("\n")
            console.print(deps_table)
        except (OSError, RuntimeError, ValueError) as e:
            console.print(f"[bold red]Error:[/bold red] {e}")

    except click.ClickException:
        raise
    except (OSError, RuntimeError, ValueError, TypeError, AttributeError) as e:
        raise click.ClickException(str(e)) from e
