# ///////////////////////////////////////////////////////////////
# _ribbon - RibbonProxy: MSO CommandBars ribbon interaction
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
RibbonProxy — MSO ribbon interaction via ``Application.CommandBars``.

Provides a thin, thread-safe wrapper around the four key
``CommandBars.*Mso`` methods that cover the most common automation
needs: executing a command, and querying its enabled / pressed /
visible state.

All COM calls are guarded by ``wrap_com_error`` and
``assert_main_thread``.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
from typing import Any

# Third-party imports
from ezplog.lib_mode import get_logger

# Local imports
from ...exceptions import GUIOperationError
from ...utils._com_utils import assert_main_thread, wrap_com_error
from .._protocols import AbstractRibbonBackend, ExcelAppLike

# ///////////////////////////////////////////////////////////////
# CONSTANTS
# ///////////////////////////////////////////////////////////////

logger = get_logger(__name__)

# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


class RibbonProxy(AbstractRibbonBackend):
    """Ribbon command execution and state queries via MSO identifiers.

    Wraps ``Application.CommandBars`` methods to execute and inspect
    built-in Excel ribbon commands without navigating the UI manually.

    Args:
        app: The active ``ExcelApp`` instance that owns this proxy.

    Example:
        >>> proxy = RibbonProxy(xl)
        >>> proxy.execute("FileSave")
        >>> proxy.is_enabled("FileSave")
        True
    """

    # ///////////////////////////////////////////////////////////////
    # INIT
    # ///////////////////////////////////////////////////////////////

    def __init__(self, app: ExcelAppLike) -> None:
        self._app = app

    # ///////////////////////////////////////////////////////////////
    # PRIVATE HELPERS
    # ///////////////////////////////////////////////////////////////

    def _check_thread(self) -> None:
        """Assert the call originates from the COM apartment thread."""
        assert_main_thread(self._app._thread_id)

    def _command_bars(self) -> Any:
        """Return the underlying ``Application.CommandBars`` COM object."""
        bars: Any = self._app._get_app().CommandBars
        return bars

    # ///////////////////////////////////////////////////////////////
    # PUBLIC METHODS
    # ///////////////////////////////////////////////////////////////

    @wrap_com_error
    def execute(self, mso_id: str) -> None:
        """Execute a built-in ribbon command by its MSO identifier.

        Calls ``Application.CommandBars.ExecuteMso(mso_id)``. Use this
        to trigger any standard Excel ribbon button programmatically.

        Args:
            mso_id: MSO control identifier string
                (e.g. ``"FileSave"``, ``"Copy"``, ``"PasteValues"``).

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the MSO ID is unknown or the command
                cannot be executed in the current application state.

        Example:
            >>> ribbon.execute("FileSave")
        """
        self._check_thread()
        logger.debug("RibbonProxy.execute: mso_id=%r", mso_id)
        try:
            self._command_bars().ExecuteMso(mso_id)
        except Exception as exc:
            # wrap_com_error handles pywintypes.com_error; this re-raise
            # catches any other unexpected error and surfaces it as a
            # GUIOperationError with full context.
            raise GUIOperationError(
                f"Failed to execute ribbon command {mso_id!r}: {exc}", cause=exc
            ) from exc

    @wrap_com_error
    def is_enabled(self, mso_id: str) -> bool:
        """Return whether a ribbon command is currently enabled.

        Calls ``Application.CommandBars.GetEnabledMso(mso_id)``.

        Args:
            mso_id: MSO control identifier string.

        Returns:
            bool: ``True`` if the command is enabled; ``False`` otherwise.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the MSO ID is unknown or the query fails.

        Example:
            >>> ribbon.is_enabled("FileSave")
            True
        """
        self._check_thread()
        logger.debug("RibbonProxy.is_enabled: mso_id=%r", mso_id)
        try:
            return bool(self._command_bars().GetEnabledMso(mso_id))
        except Exception as exc:
            raise GUIOperationError(
                f"Failed to query enabled state for {mso_id!r}: {exc}", cause=exc
            ) from exc

    @wrap_com_error
    def is_pressed(self, mso_id: str) -> bool:
        """Return whether a ribbon toggle command is currently pressed.

        Calls ``Application.CommandBars.GetPressedMso(mso_id)``.

        Args:
            mso_id: MSO control identifier string.

        Returns:
            bool: ``True`` if the toggle command is in the pressed/active
                state; ``False`` if not pressed **or** if the control does not
                support the pressed-state query (e.g. regular buttons such as
                ``"FileSave"`` always return ``False``).

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.

        Example:
            >>> ribbon.is_pressed("Bold")
            False
        """
        self._check_thread()
        logger.debug("RibbonProxy.is_pressed: mso_id=%r", mso_id)
        try:
            return bool(self._command_bars().GetPressedMso(mso_id))
        except Exception:
            # GetPressedMso raises COM E_INVALIDARG for non-toggle controls
            # (regular buttons, split buttons, etc.). Treat as "not pressed"
            # rather than propagating an error — the control simply has no
            # pressed state.
            logger.debug(
                "RibbonProxy.is_pressed: %r does not support pressed-state query"
                " — returning False.",
                mso_id,
            )
            return False

    @wrap_com_error
    def is_visible(self, mso_id: str) -> bool:
        """Return whether a ribbon command is currently visible.

        Calls ``Application.CommandBars.GetVisibleMso(mso_id)``.

        Args:
            mso_id: MSO control identifier string.

        Returns:
            bool: ``True`` if the command is visible in the current ribbon
                state; ``False`` otherwise.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the MSO ID is unknown or the query fails.

        Example:
            >>> ribbon.is_visible("FileSave")
            True
        """
        self._check_thread()
        logger.debug("RibbonProxy.is_visible: mso_id=%r", mso_id)
        try:
            return bool(self._command_bars().GetVisibleMso(mso_id))
        except Exception as exc:
            raise GUIOperationError(
                f"Failed to query visible state for {mso_id!r}: {exc}", cause=exc
            ) from exc
