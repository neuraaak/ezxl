# ///////////////////////////////////////////////////////////////
# EZXL - CLI Main Entry Point
# Project: ezxl
# ///////////////////////////////////////////////////////////////

"""
Main CLI entry point for EzXl Excel automation library.

This module provides the command-line interface for managing EzXl
configuration and performing various operations.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Third-party imports
import click
from rich.panel import Panel
from rich.text import Text

# Local imports
from .._version import __version__
from ._console import console
from .commands import _docs, _info, _version

# ///////////////////////////////////////////////////////////////
# CLI GROUP
# ///////////////////////////////////////////////////////////////


@click.group(
    name="ezxl",
    invoke_without_command=True,
    context_settings={"help_option_names": ["-h", "--help"]},
)
@click.version_option(
    __version__,
    "-v",
    "--version",
    prog_name="EzXl CLI",
    message="%(prog)s version %(version)s",
)
@click.pass_context
def cli(ctx: click.Context) -> None:
    """
    EzXl CLI - Excel Automation Library

    Command-line interface for managing EzXl Excel automation operations.

    Use 'ezxl <command> --help' for more information on a specific command.
    """
    # If no subcommand is invoked, display the welcome screen and help
    if ctx.invoked_subcommand is None:
        _display_welcome()
        click.echo(ctx.get_help())


def _display_welcome() -> None:
    """Display welcome message with Rich."""
    try:
        welcome_text = Text()
        welcome_text.append("🚀 ", style="bold bright_green")
        welcome_text.append("EzXl CLI", style="bold bright_blue")
        welcome_text.append(" - Excel Automation Library", style="dim white")

        panel = Panel(
            welcome_text,
            title="[bold bright_blue]Welcome[/bold bright_blue]",
            border_style="bright_blue",
            padding=(1, 2),
        )
        console.print(panel)
    except (OSError, RuntimeError, ValueError):
        # Fallback if Rich is not available
        click.echo("EzXl CLI - Excel Automation Library")


# ///////////////////////////////////////////////////////////////
# COMMAND GROUPS
# ///////////////////////////////////////////////////////////////


# Register commands and groups
cli.add_command(_version.version_command)
cli.add_command(_info.info_command)
cli.add_command(_docs.docs_command)


# ///////////////////////////////////////////////////////////////
# MAIN ENTRY POINT
# ///////////////////////////////////////////////////////////////


def main() -> None:
    """
    Main entry point for the CLI.

    This function is called when the CLI is invoked from the command line.
    """
    try:
        cli()
    except click.ClickException as e:
        e.show()
        raise SystemExit(e.exit_code) from e
    except KeyboardInterrupt as e:
        console.print("\n[yellow]Interrupted by user[/yellow]")
        raise SystemExit(1) from e
    except (OSError, RuntimeError, ValueError) as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        raise SystemExit(1) from e


if __name__ == "__main__":
    main()
