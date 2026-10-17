# ///////////////////////////////////////////////////////////////
# EZXL - ExcelApp acquisition mode unit tests
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""Unit tests for how ``ExcelApp`` acquires and releases an Excel instance.

The COM layer is substituted, so these assert the ownership rules without
starting Excel: which call is made per mode, and — the part that matters —
that a session never quits an instance it did not start.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
import sys
from typing import Any

# Third-party imports
import pytest

pytestmark = pytest.mark.skipif(
    sys.platform != "win32", reason="core._excel_app imports win32com (Windows-only)"
)

# ///////////////////////////////////////////////////////////////
# DOUBLES
# ///////////////////////////////////////////////////////////////


class FakeApp:
    """Stands in for the COM Application object."""

    def __init__(self) -> None:
        self.Visible: bool = False
        self.DisplayAlerts: bool = True
        self.quit_calls: int = 0

    def Quit(self) -> None:  # noqa: N802 - mirrors the COM method name
        self.quit_calls += 1


class FakeWin32:
    """Records which acquisition call was made."""

    def __init__(self, *, running: FakeApp | None = None) -> None:
        self._running = running
        self.calls: list[str] = []

    def GetActiveObject(self, _progid: str) -> FakeApp:  # noqa: N802
        self.calls.append("GetActiveObject")
        if self._running is None:
            raise OSError("no running instance")
        return self._running

    def DispatchEx(self, _progid: str) -> FakeApp:  # noqa: N802
        self.calls.append("DispatchEx")
        return FakeApp()


@pytest.fixture
def patch_win32(monkeypatch: pytest.MonkeyPatch) -> Any:
    """Return a factory that installs a FakeWin32 into the module."""
    from ezxl.core import _excel_app

    def install(*, running: FakeApp | None = None) -> FakeWin32:
        fake = FakeWin32(running=running)
        monkeypatch.setattr(_excel_app, "win32", fake)
        return fake

    return install


# ///////////////////////////////////////////////////////////////
# TESTS — which call each mode makes
# ///////////////////////////////////////////////////////////////


def test_should_start_a_private_instance_when_mode_is_dispatch(
    patch_win32: Any,
) -> None:
    """dispatch must never adopt a running Excel, even when one exists."""
    from ezxl import ExcelApp

    running = FakeApp()
    fake = patch_win32(running=running)

    with ExcelApp(mode="dispatch", visible=False) as xl:
        assert fake.calls == ["DispatchEx"]
        assert xl._app is not running


def test_should_use_the_running_instance_when_mode_is_attach(patch_win32: Any) -> None:
    """attach binds to what is already open."""
    from ezxl import ExcelApp

    running = FakeApp()
    fake = patch_win32(running=running)

    with ExcelApp(mode="attach") as xl:
        assert fake.calls == ["GetActiveObject"]
        assert xl._app is running


def test_should_raise_when_attach_finds_no_running_instance(patch_win32: Any) -> None:
    """attach is explicit: no fallback, a clear error instead."""
    from ezxl import ExcelApp
    from ezxl.exceptions import ExcelNotAvailableError

    patch_win32(running=None)

    with pytest.raises(ExcelNotAvailableError, match="No running Excel instance"):
        ExcelApp(mode="attach").__enter__()


def test_should_prefer_the_running_instance_when_mode_is_auto(patch_win32: Any) -> None:
    """auto reuses the Excel the user already has open."""
    from ezxl import ExcelApp

    running = FakeApp()
    fake = patch_win32(running=running)

    with ExcelApp(mode="auto") as xl:
        assert fake.calls == ["GetActiveObject"]
        assert xl._app is running


def test_should_fall_back_to_a_private_instance_when_auto_finds_none(
    patch_win32: Any,
) -> None:
    """auto starts Excel when nothing is running — the old Dispatch fallback."""
    from ezxl import ExcelApp

    fake = patch_win32(running=None)

    with ExcelApp(mode="auto", visible=False) as xl:
        assert fake.calls == ["GetActiveObject", "DispatchEx"]
        assert xl._app is not None


# ///////////////////////////////////////////////////////////////
# TESTS — ownership: never quit what we did not start
# ///////////////////////////////////////////////////////////////


def test_should_quit_on_exit_when_the_session_started_excel(patch_win32: Any) -> None:
    """A dispatched instance is ours to close."""
    from ezxl import ExcelApp

    patch_win32(running=None)

    app = ExcelApp(mode="dispatch", visible=False)
    with app:
        com = app._app
    assert com.quit_calls == 1


@pytest.mark.parametrize("mode", ["attach", "auto"])
def test_should_leave_excel_running_on_exit_when_it_was_already_open(
    patch_win32: Any, mode: str
) -> None:
    """The regression that matters: the user's Excel must survive the exit.

    Regression guard: dispatch mode used to adopt a running instance and quit
    it on __exit__, discarding the user's unsaved work.
    """
    from ezxl import ExcelApp

    running = FakeApp()
    patch_win32(running=running)

    with ExcelApp(mode=mode):  # type: ignore[arg-type]
        pass

    assert running.quit_calls == 0


@pytest.mark.parametrize("mode", ["attach", "auto"])
def test_should_only_release_when_quit_is_called_on_a_borrowed_instance(
    patch_win32: Any, mode: str
) -> None:
    """An explicit quit() must not close an Excel the session did not start."""
    from ezxl import ExcelApp

    running = FakeApp()
    patch_win32(running=running)

    app = ExcelApp(mode=mode)  # type: ignore[arg-type]
    app._get_app()
    app.quit()

    assert running.quit_calls == 0
    assert app._app is None


def test_should_not_touch_visible_when_attaching(patch_win32: Any) -> None:
    """Hiding the user's window makes Excel look crashed, so leave it alone."""
    from ezxl import ExcelApp

    running = FakeApp()
    running.Visible = True
    patch_win32(running=running)

    with ExcelApp(mode="attach", visible=False):
        pass

    assert running.Visible is True


def test_should_apply_visible_when_starting_a_private_instance(
    patch_win32: Any,
) -> None:
    """The flag still applies to an instance we own."""
    from ezxl import ExcelApp

    patch_win32(running=None)

    app = ExcelApp(mode="dispatch", visible=True)
    with app:
        assert app._app.Visible is True


def test_should_release_and_not_leak_when_initialisation_fails(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A failure after DispatchEx must quit the orphan and clear the cache.

    Regression guard: _app was assigned before Visible was set, so a failing
    assignment left the process running and every later call returned the
    broken object.
    """
    from ezxl import ExcelApp
    from ezxl.core import _excel_app
    from ezxl.exceptions import ExcelNotAvailableError

    created: list[FakeApp] = []

    class ExplodingApp(FakeApp):
        """Refuses the Visible assignment, as a shutting-down Excel does."""

        def __init__(self) -> None:
            super().__init__()
            # Armed only after construction, so __init__ can set its defaults.
            self._armed = True

        def __setattr__(self, name: str, value: object) -> None:
            if name == "Visible" and getattr(self, "_armed", False):
                raise AttributeError("Property 'Visible' can not be set.")
            super().__setattr__(name, value)

    class ExplodingWin32(FakeWin32):
        def DispatchEx(self, _progid: str) -> FakeApp:  # noqa: N802
            app = ExplodingApp()
            created.append(app)
            return app

    monkeypatch.setattr(_excel_app, "win32", ExplodingWin32(running=None))

    app = ExcelApp(mode="dispatch")
    with pytest.raises(ExcelNotAvailableError):
        app._get_app()

    assert app._app is None, "the broken object must not be cached"
    assert created[0].quit_calls == 1, "the orphaned process must be quit"
