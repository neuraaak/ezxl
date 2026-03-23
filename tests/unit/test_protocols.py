# ///////////////////////////////////////////////////////////////
# test_protocols - EzXl GUI protocol contract tests
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
Unit tests for the abstract GUI backend protocols.

Verifies the four ABC contracts defined in ``ezxl.gui._protocols``:

- Concrete subclasses that omit required abstract methods cannot be
  instantiated (Python raises ``TypeError``).
- Concrete subclasses that implement all abstract methods can be
  instantiated and used without error.

No COM calls are made in this module — all implementations are pure
Python stubs.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Third-party imports
import pytest

# Local imports
from ezxl.gui._protocols import (
    AbstractDialogBackend,
    AbstractKeysBackend,
    AbstractMenuBackend,
    AbstractRibbonBackend,
)

# ///////////////////////////////////////////////////////////////
# PARTIAL STUBS (missing at least one abstract method each)
# ///////////////////////////////////////////////////////////////


class _PartialRibbon(AbstractRibbonBackend):
    """Ribbon stub missing ``is_enabled``, ``is_pressed``, and ``is_visible``."""

    def execute(self, mso_id: str) -> None: ...  # noqa: D102,ARG002


class _PartialMenu(AbstractMenuBackend):
    """Menu stub missing ``list_bars`` and ``list_controls``."""

    def click(self, bar_name: str, *item_path: str) -> None: ...  # noqa: D102,ARG002


class _PartialDialog(AbstractDialogBackend):
    """Dialog stub missing ``get_file_save`` and ``alert``."""

    def get_file_open(
        self,
        title: str = "Open",  # noqa: ARG002
        initial_dir: str | None = None,  # noqa: ARG002
        filter: str = "Excel Files (*.xls*), *.xls*",  # noqa: A002,ARG002
    ) -> str | None:  # noqa: D102
        return None


class _PartialKeys(AbstractKeysBackend):
    """Keys stub with no methods implemented."""


# ///////////////////////////////////////////////////////////////
# FULL STUBS (all abstract methods implemented)
# ///////////////////////////////////////////////////////////////


class _FullRibbon(AbstractRibbonBackend):
    """Minimal but complete ribbon backend stub."""

    def execute(self, mso_id: str) -> None: ...  # noqa: D102,ARG002

    def is_enabled(self, mso_id: str) -> bool:  # noqa: ARG002
        return True

    def is_pressed(self, mso_id: str) -> bool:  # noqa: ARG002
        return False

    def is_visible(self, mso_id: str) -> bool:  # noqa: ARG002
        return True


class _FullMenu(AbstractMenuBackend):
    """Minimal but complete menu backend stub."""

    def click(self, bar_name: str, *item_path: str) -> None: ...  # noqa: D102,ARG002

    def list_bars(self) -> list[str]:  # noqa: D102
        return []

    def list_controls(self, bar_name: str) -> list[str]:  # noqa: ARG002
        return []


class _FullDialog(AbstractDialogBackend):
    """Minimal but complete dialog backend stub."""

    def get_file_open(
        self,
        title: str = "Open",  # noqa: ARG002
        initial_dir: str | None = None,  # noqa: ARG002
        filter: str = "Excel Files (*.xls*), *.xls*",  # noqa: A002,ARG002
    ) -> str | None:  # noqa: D102
        return None

    def get_file_save(
        self,
        title: str = "Save As",  # noqa: ARG002
        initial_dir: str | None = None,  # noqa: ARG002
        filter: str = "Excel Files (*.xlsx), *.xlsx",  # noqa: A002,ARG002
    ) -> str | None:  # noqa: D102
        return None

    def alert(self, message: str, title: str = "EzXl") -> None: ...  # noqa: D102,ARG002


class _FullKeys(AbstractKeysBackend):
    """Minimal but complete keys backend stub."""

    def send_keys(self, keys: str, wait: bool = True) -> None: ...  # noqa: D102,ARG002


# ///////////////////////////////////////////////////////////////
# TESTS — partial implementations are rejected
# ///////////////////////////////////////////////////////////////


@pytest.mark.unit
def test_should_raise_type_error_when_abstractribbon_not_fully_implemented() -> None:
    """Verify that a partial AbstractRibbonBackend cannot be instantiated."""
    with pytest.raises(TypeError):
        _PartialRibbon()  # type: ignore[abstract]


@pytest.mark.unit
def test_should_raise_type_error_when_abstractmenu_not_fully_implemented() -> None:
    """Verify that a partial AbstractMenuBackend cannot be instantiated."""
    with pytest.raises(TypeError):
        _PartialMenu()  # type: ignore[abstract]


@pytest.mark.unit
def test_should_raise_type_error_when_abstractdialog_not_fully_implemented() -> None:
    """Verify that a partial AbstractDialogBackend cannot be instantiated."""
    with pytest.raises(TypeError):
        _PartialDialog()  # type: ignore[abstract]


@pytest.mark.unit
def test_should_raise_type_error_when_abstractkeys_not_fully_implemented() -> None:
    """Verify that a partial AbstractKeysBackend cannot be instantiated."""
    with pytest.raises(TypeError):
        _PartialKeys()  # type: ignore[abstract]


# ///////////////////////////////////////////////////////////////
# TESTS — full implementations are accepted
# ///////////////////////////////////////////////////////////////


@pytest.mark.unit
def test_should_accept_fully_implemented_ribbon_backend() -> None:
    """Verify that a fully implemented AbstractRibbonBackend can be instantiated."""
    ribbon = _FullRibbon()
    assert isinstance(ribbon, AbstractRibbonBackend)


@pytest.mark.unit
def test_should_accept_fully_implemented_menu_backend() -> None:
    """Verify that a fully implemented AbstractMenuBackend can be instantiated."""
    menu = _FullMenu()
    assert isinstance(menu, AbstractMenuBackend)


@pytest.mark.unit
def test_should_accept_fully_implemented_dialog_backend() -> None:
    """Verify that a fully implemented AbstractDialogBackend can be instantiated."""
    dialog = _FullDialog()
    assert isinstance(dialog, AbstractDialogBackend)


@pytest.mark.unit
def test_should_accept_fully_implemented_keys_backend() -> None:
    """Verify that a fully implemented AbstractKeysBackend can be instantiated."""
    keys = _FullKeys()
    assert isinstance(keys, AbstractKeysBackend)
