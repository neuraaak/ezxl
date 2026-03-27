# ///////////////////////////////////////////////////////////////
# _com_utils - Internal COM utility functions
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
Internal COM utility functions.

Low-level helpers shared across the COM proxy layer. Not part of the
public API — do not import from outside the ezxl package.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
import threading
import time
from collections.abc import Callable
from functools import wraps
from typing import Any, TypeVar

# Third-party imports
import pywintypes
from ezplog.lib_mode import get_logger

# Local imports
from ..exceptions import (
    COMOperationError,
    ExcelSessionLostError,
    ExcelThreadViolationError,
)

# ///////////////////////////////////////////////////////////////
# CONSTANTS
# ///////////////////////////////////////////////////////////////

logger = get_logger(__name__)

# pywintypes ships without complete type stubs — cache the runtime types via
# getattr so that type checkers (ty, pyright) don't flag missing attributes.
_COM_ERROR_CLASS: type = pywintypes.com_error  # type: ignore[attr-defined]

# COM HRESULT codes that indicate a dead/disconnected server.
# RPC_E_DISCONNECTED (0x80010108) and RPC_E_SERVER_DIED (0x80010007).
_DISCONNECTED_HRESULTS: frozenset[int] = frozenset(
    [
        0x80010108,  # RPC_E_DISCONNECTED
        0x80010007,  # RPC_E_SERVER_DIED
        0x800706BA,  # RPC_S_SERVER_UNAVAILABLE
    ]
)

_FuncT = TypeVar("_FuncT", bound=Callable[..., Any])

# ///////////////////////////////////////////////////////////////
# FUNCTIONS
# ///////////////////////////////////////////////////////////////


def wait_until_ready(xl_app: Any, timeout: float = 30.0) -> None:
    """Poll ``xl_app.Ready`` until Excel reports it is ready or the timeout expires.

    Excel's ``Application.Ready`` property returns ``False`` while Excel is
    performing an operation (e.g., opening a file, calculating). This helper
    blocks the calling thread with short sleeps until Excel is available again.

    Args:
        xl_app: A live COM ``Application`` object (``win32com.client`` dispatch).
        timeout: Maximum number of seconds to wait before raising. Defaults to 30.

    Raises:
        COMOperationError: If ``timeout`` is exceeded before Excel becomes ready.

    Example:
        >>> wait_until_ready(xl_app, timeout=60.0)
    """
    deadline = time.monotonic() + timeout
    poll_interval = 0.25  # seconds between each Ready check

    while time.monotonic() < deadline:
        try:
            if xl_app.Ready:
                return
        except Exception as exc:
            # Ready property itself can throw when Excel is in a transient state.
            # Log at DEBUG and retry until the deadline.
            logger.debug("wait_until_ready: Ready check raised (will retry): %s", exc)
        time.sleep(poll_interval)

    raise COMOperationError(f"Excel did not become ready within {timeout:.1f} seconds.")


def wrap_com_error(func: _FuncT) -> _FuncT:
    """Decorator that intercepts ``pywintypes.com_error`` and re-raises as EzXl exceptions.

    Wraps any COM boundary function so that raw pywin32 errors never escape
    the ``ezxl`` package. Disconnection HRESULTs are mapped to
    ``ExcelSessionLostError``; all other COM errors become ``COMOperationError``.

    Args:
        func: The callable to wrap.

    Returns:
        The wrapped callable with identical signature.

    Example:
        >>> @wrap_com_error
        ... def open_workbook(xl_app, path):
        ...     return xl_app.Workbooks.Open(path)
    """

    @wraps(func)
    def _wrapper(*args: Any, **kwargs: Any) -> Any:
        try:
            return func(*args, **kwargs)
        except Exception as exc:
            if isinstance(exc, _COM_ERROR_CLASS):
                hresult: int = exc.args[0] if exc.args else 0
                if hresult in _DISCONNECTED_HRESULTS:
                    raise ExcelSessionLostError(
                        f"Excel COM session lost (HRESULT 0x{hresult:08X}): {exc}",
                        cause=exc,
                    ) from exc
                raise COMOperationError(
                    f"COM operation failed (HRESULT 0x{hresult:08X}): {exc}",
                    cause=exc,
                ) from exc
            # Re-raise non-COM exceptions untouched.
            raise

    return _wrapper  # type: ignore[return-value]


def assert_main_thread(thread_id: int) -> None:
    """Assert that the current thread matches the expected COM apartment thread.

    COM STA (Single-Threaded Apartment) requires that all calls on a COM object
    are made from the thread that created it. This function raises proactively
    before reaching the COM dispatcher, giving callers a clear diagnostic.

    Args:
        thread_id: The thread identifier recorded at ``ExcelApp`` construction time.

    Raises:
        ExcelThreadViolationError: If the calling thread differs from ``thread_id``.

    Example:
        >>> assert_main_thread(self._thread_id)
    """
    current = threading.get_ident()
    if current != thread_id:
        raise ExcelThreadViolationError(
            f"COM call attempted from thread {current}, but ExcelApp was created "
            f"on thread {thread_id}. Excel COM uses STA — all calls must originate "
            f"from the creating thread."
        )
