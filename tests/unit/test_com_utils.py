# ///////////////////////////////////////////////////////////////
# EZXL - COM utility unit tests
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""Unit tests for ``utils/_com_utils``.

These exercise the COM error-translation logic only — no Excel process and no
COM call is involved, so they run on every platform where pywin32 imports.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
import sys

# Third-party imports
import pytest

pytestmark = pytest.mark.skipif(
    sys.platform != "win32", reason="pywintypes is Windows-only"
)

# ///////////////////////////////////////////////////////////////
# TESTS — wrap_com_error
# ///////////////////////////////////////////////////////////////


def test_should_raise_session_lost_when_hresult_is_signed_disconnect() -> None:
    """A disconnect HRESULT arrives signed from pywin32 and must still map.

    Regression guard: the lookup set holds unsigned values, so without masking
    the signed value never matched and every lost session was reported as a
    plain COMOperationError.
    """
    import pywintypes

    from ezxl.exceptions import ExcelSessionLostError
    from ezxl.utils._com_utils import wrap_com_error

    @wrap_com_error
    def disconnected() -> None:
        # RPC_E_DISCONNECTED (0x80010108) exactly as pywin32 delivers it.
        raise pywintypes.com_error(
            -2147417848, "The object invoked has disconnected", None, None
        )

    with pytest.raises(ExcelSessionLostError):
        disconnected()


@pytest.mark.parametrize(
    "hresult",
    [
        -2147417848,  # RPC_E_DISCONNECTED, signed
        0x80010108,  # the same code unsigned
        -2147417849,  # RPC_E_SERVER_DIED, signed
        -2147023174,  # RPC_S_SERVER_UNAVAILABLE, signed
    ],
)
def test_should_report_hresult_as_unsigned_hex_when_wrapping(hresult: int) -> None:
    """The message must never contain a negative hex literal like 0x-7FFEFEF8."""
    import pywintypes

    from ezxl.exceptions import EzXlError
    from ezxl.utils._com_utils import wrap_com_error

    @wrap_com_error
    def failing() -> None:
        raise pywintypes.com_error(hresult, "boom", None, None)

    with pytest.raises(EzXlError) as excinfo:
        failing()

    assert "0x-" not in str(excinfo.value)


def test_should_raise_com_operation_error_when_hresult_is_unrelated() -> None:
    """A COM failure that is not a disconnect stays a COMOperationError."""
    import pywintypes

    from ezxl.exceptions import COMOperationError, ExcelSessionLostError
    from ezxl.utils._com_utils import wrap_com_error

    @wrap_com_error
    def failing() -> None:
        # 0x800A03EC — the generic Excel "cannot do that" error.
        raise pywintypes.com_error(-2146827284, "Excel said no", None, None)

    with pytest.raises(COMOperationError) as excinfo:
        failing()

    assert not isinstance(excinfo.value, ExcelSessionLostError)


def test_should_reraise_untouched_when_error_is_not_com() -> None:
    """Non-COM exceptions must pass through without being wrapped."""
    from ezxl.utils._com_utils import wrap_com_error

    @wrap_com_error
    def failing() -> None:
        raise ValueError("not a COM problem")

    with pytest.raises(ValueError, match="not a COM problem"):
        failing()


# ///////////////////////////////////////////////////////////////
# TESTS — assert_main_thread
# ///////////////////////////////////////////////////////////////


def test_should_raise_thread_violation_when_called_from_another_thread() -> None:
    """COM STA: a call from a foreign thread must be refused before COM sees it."""
    import threading

    from ezxl.exceptions import ExcelThreadViolationError
    from ezxl.utils._com_utils import assert_main_thread

    owner = threading.get_ident()
    captured: list[BaseException] = []

    def from_other_thread() -> None:
        try:
            assert_main_thread(owner)
        except BaseException as exc:  # noqa: BLE001 - recorded for the assert
            captured.append(exc)

    worker = threading.Thread(target=from_other_thread)
    worker.start()
    worker.join()

    assert len(captured) == 1
    assert isinstance(captured[0], ExcelThreadViolationError)


def test_should_pass_when_called_from_the_owning_thread() -> None:
    """The happy path must not raise."""
    import threading

    from ezxl.utils._com_utils import assert_main_thread

    assert_main_thread(threading.get_ident())


# ///////////////////////////////////////////////////////////////
# TESTS — wait_until_ready
# ///////////////////////////////////////////////////////////////


def test_should_return_when_app_reports_ready() -> None:
    """A ready app returns immediately."""
    from ezxl.utils._com_utils import wait_until_ready

    class ReadyApp:
        Ready = True

    wait_until_ready(ReadyApp(), timeout=1.0)


def test_should_raise_com_operation_error_when_never_ready() -> None:
    """A never-ready app must time out rather than block forever."""
    from ezxl.exceptions import COMOperationError
    from ezxl.utils._com_utils import wait_until_ready

    class BusyApp:
        Ready = False

    with pytest.raises(COMOperationError, match="did not become ready"):
        wait_until_ready(BusyApp(), timeout=0.3)


def test_should_keep_polling_when_ready_property_raises() -> None:
    """Excel's Ready can throw in a transient state; that must not abort the wait."""
    from ezxl.utils._com_utils import wait_until_ready

    class FlakyApp:
        def __init__(self) -> None:
            self._calls = 0

        @property
        def Ready(self) -> bool:  # noqa: N802 - mirrors the COM property name
            self._calls += 1
            if self._calls < 3:
                raise OSError("transient COM state")
            return True

    app = FlakyApp()
    wait_until_ready(app, timeout=5.0)
    assert app._calls >= 3
