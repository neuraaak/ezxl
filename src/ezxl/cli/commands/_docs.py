# ///////////////////////////////////////////////////////////////
# EZXL - CLI Docs Command
# Project: ezxl
# ///////////////////////////////////////////////////////////////

"""CLI command for opening ezxl online documentation."""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
import webbrowser

# Third-party imports
import click

# ///////////////////////////////////////////////////////////////
# CONSTANTS
# ///////////////////////////////////////////////////////////////

DOCS_URL = "https://neuraaak.github.io/ezxl/"


# ///////////////////////////////////////////////////////////////
# COMMANDS
# ///////////////////////////////////////////////////////////////


@click.command(name="docs", help="Open the online documentation")
def docs_command() -> None:
    """Open the ezxl documentation website in the default browser."""
    try:
        opened = webbrowser.open(DOCS_URL, new=2)
    except (OSError, RuntimeError, webbrowser.Error) as e:
        raise click.ClickException(str(e)) from e

    if opened:
        click.echo(f"Opened documentation: {DOCS_URL}")
        return

    # Fallback for environments where a browser cannot be opened.
    click.echo(DOCS_URL)
