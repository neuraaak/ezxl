# ///////////////////////////////////////////////////////////////
# EZXL - CLI unit tests
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""Unit tests for the ``ezxl`` command-line interface.

The CLI is a user-facing entry point and needs no Excel process, so it is
covered here rather than excluded from the coverage gate. Every test drives
the commands through click's own ``CliRunner``.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
import webbrowser

# Third-party imports
import click
import pytest
from click.testing import CliRunner

# Local imports
import ezxl
from ezxl.cli.commands import _docs
from ezxl.cli.main import cli, main

pytestmark = pytest.mark.cli

# ///////////////////////////////////////////////////////////////
# FIXTURES
# ///////////////////////////////////////////////////////////////


@pytest.fixture
def runner() -> CliRunner:
    """Return a click test runner."""
    return CliRunner()


# ///////////////////////////////////////////////////////////////
# TESTS — group behaviour
# ///////////////////////////////////////////////////////////////


def test_should_show_welcome_and_help_when_invoked_without_a_subcommand(
    runner: CliRunner,
) -> None:
    """The bare `ezxl` call is documented to print a welcome plus the help."""
    result = runner.invoke(cli, [])

    assert result.exit_code == 0
    assert "EzXl" in result.output
    assert "Commands" in result.output


@pytest.mark.parametrize("flag", ["-h", "--help"])
def test_should_list_every_command_when_help_is_requested(
    runner: CliRunner, flag: str
) -> None:
    """Both help flags are wired, and all three commands are registered."""
    result = runner.invoke(cli, [flag])

    assert result.exit_code == 0
    for command in ("version", "info", "docs"):
        assert command in result.output


@pytest.mark.parametrize("flag", ["-v", "--version"])
def test_should_print_the_package_version_when_version_flag_is_given(
    runner: CliRunner, flag: str
) -> None:
    """The group-level version flag reports the installed version."""
    result = runner.invoke(cli, [flag])

    assert result.exit_code == 0
    assert ezxl.__version__ in result.output


def test_should_exit_with_an_error_when_the_command_is_unknown(
    runner: CliRunner,
) -> None:
    """An unknown subcommand must fail rather than fall through to the welcome."""
    result = runner.invoke(cli, ["no-such-command"])

    assert result.exit_code != 0


# ///////////////////////////////////////////////////////////////
# TESTS — version command
# ///////////////////////////////////////////////////////////////


def test_should_report_the_version_when_version_command_runs(
    runner: CliRunner,
) -> None:
    """`ezxl version` prints the short form."""
    result = runner.invoke(cli, ["version"])

    assert result.exit_code == 0
    assert ezxl.__version__ in result.output


def test_should_include_the_author_when_version_command_is_full(
    runner: CliRunner,
) -> None:
    """`ezxl version --full` adds the author block."""
    result = runner.invoke(cli, ["version", "--full"])

    assert result.exit_code == 0
    assert ezxl.__version__ in result.output
    assert "Author" in result.output


# ///////////////////////////////////////////////////////////////
# TESTS — info command
# ///////////////////////////////////////////////////////////////


def test_should_report_package_metadata_when_info_command_runs(
    runner: CliRunner,
) -> None:
    """`ezxl info` prints the package name and version without failing."""
    result = runner.invoke(cli, ["info"])

    assert result.exit_code == 0
    assert "ezxl" in result.output.lower()
    assert ezxl.__version__ in result.output


# ///////////////////////////////////////////////////////////////
# TESTS — docs command
# ///////////////////////////////////////////////////////////////


def test_should_report_the_url_when_the_browser_opens(
    runner: CliRunner, monkeypatch: pytest.MonkeyPatch
) -> None:
    """A successful browser launch is confirmed to the user."""
    opened: list[str] = []

    def _fake_open(url: str, **_kwargs: object) -> bool:
        opened.append(url)
        return True

    monkeypatch.setattr(_docs.webbrowser, "open", _fake_open)
    result = runner.invoke(cli, ["docs"])

    assert result.exit_code == 0
    assert opened == [_docs.DOCS_URL]
    assert _docs.DOCS_URL in result.output


def test_should_print_the_bare_url_when_no_browser_is_available(
    runner: CliRunner, monkeypatch: pytest.MonkeyPatch
) -> None:
    """A headless machine still gets the URL, not an error."""
    monkeypatch.setattr(_docs.webbrowser, "open", lambda _url, **_kwargs: False)
    result = runner.invoke(cli, ["docs"])

    assert result.exit_code == 0
    assert result.output.strip() == _docs.DOCS_URL


def test_should_raise_a_click_exception_when_the_browser_errors(
    runner: CliRunner, monkeypatch: pytest.MonkeyPatch
) -> None:
    """A browser failure is surfaced as a CLI error, not a traceback."""

    def _raise(_url: str, **_kwargs: object) -> bool:
        raise webbrowser.Error("no browser")

    monkeypatch.setattr(_docs.webbrowser, "open", _raise)
    result = runner.invoke(cli, ["docs"])

    assert result.exit_code != 0
    assert "no browser" in result.output


# ///////////////////////////////////////////////////////////////
# TESTS — main() entry point
# ///////////////////////////////////////////////////////////////


def test_should_exit_with_the_click_code_when_a_click_exception_escapes(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """`main` translates a ClickException into its own exit code."""

    def _raise() -> None:
        raise click.ClickException("bad usage")

    monkeypatch.setattr("ezxl.cli.main.cli", _raise)

    with pytest.raises(SystemExit) as excinfo:
        main()

    assert excinfo.value.code == 1


def test_should_exit_with_one_when_interrupted_by_the_user(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Ctrl+C exits cleanly instead of printing a traceback."""

    def _raise() -> None:
        raise KeyboardInterrupt

    monkeypatch.setattr("ezxl.cli.main.cli", _raise)

    with pytest.raises(SystemExit) as excinfo:
        main()

    assert excinfo.value.code == 1


def test_should_exit_with_one_when_an_os_error_escapes(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """An unexpected OSError is reported, not raised at the terminal."""

    def _raise() -> None:
        raise OSError("disk gone")

    monkeypatch.setattr("ezxl.cli.main.cli", _raise)

    with pytest.raises(SystemExit) as excinfo:
        main()

    assert excinfo.value.code == 1
