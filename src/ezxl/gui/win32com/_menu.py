# ///////////////////////////////////////////////////////////////
# _menu - MenuProxy: legacy CommandBar navigation and execution
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
MenuProxy — legacy CommandBar interaction via ``Application.CommandBars``.

Provides access to Excel's classic CommandBar (pre-ribbon) menu system.
Supports traversing named bars by a sequence of control captions and
executing the terminal control.

All COM calls are guarded by ``wrap_com_error`` and
``assert_main_thread``.

Note:
    CommandBars are a legacy API. Many bars are hidden or non-functional
    in modern Excel (2013+). Prefer ``RibbonProxy.execute`` for built-in
    commands. Use ``MenuProxy`` only when a specific add-in or legacy
    toolbar must be driven.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
from typing import TYPE_CHECKING, Any

# Third-party imports
from ezplog.lib_mode import get_logger

# Local imports
from ...exceptions import GUIOperationError
from ...utils._com_utils import assert_main_thread, wrap_com_error
from .._protocols import AbstractMenuBackend

if TYPE_CHECKING:
    from ...core._excel_app import ExcelApp

# ///////////////////////////////////////////////////////////////
# CONSTANTS
# ///////////////////////////////////////////////////////////////

logger = get_logger(__name__)

# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


class MenuProxy(AbstractMenuBackend):
    """Legacy CommandBar menu traversal and control execution.

    Allows navigating an Excel CommandBar by name and executing controls
    by their caption path. Also exposes discovery helpers to list all
    available bars and the top-level controls of any given bar.

    Args:
        app: The active ``ExcelApp`` instance that owns this proxy.

    Example:
        >>> proxy = MenuProxy(xl)
        >>> proxy.list_bars()
        ['Standard', 'Formatting', ...]
        >>> proxy.click("Standard", "Open")
    """

    # ///////////////////////////////////////////////////////////////
    # INIT
    # ///////////////////////////////////////////////////////////////

    def __init__(self, app: ExcelApp) -> None:
        self._app = app

    # ///////////////////////////////////////////////////////////////
    # PRIVATE HELPERS
    # ///////////////////////////////////////////////////////////////

    def _check_thread(self) -> None:
        """Assert the call originates from the COM apartment thread."""
        assert_main_thread(self._app._thread_id)

    def _command_bars(self) -> Any:
        """Return the underlying ``Application.CommandBars`` COM collection."""
        bars: Any = self._app._get_app().CommandBars
        return bars

    # ///////////////////////////////////////////////////////////////
    # PUBLIC METHODS
    # ///////////////////////////////////////////////////////////////

    @wrap_com_error
    def click(self, bar_name: str, *item_path: str) -> None:
        """Traverse a CommandBar by caption path and execute the final control.

        Looks up the CommandBar named ``bar_name``, then iterates through
        ``item_path`` — each element being the caption of a nested control —
        descending into sub-controls at each step. Calls ``.Execute()`` on
        the control identified by the last element.

        Args:
            bar_name: The name of the CommandBar to start from
                (e.g. ``"Standard"``).
            *item_path: One or more control captions forming the path to
                the target control. At least one caption is required.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the bar, any intermediate control, or
                the final control cannot be found, or if ``Execute``
                fails.

        Example:
            >>> menu.click("Standard", "Open")
            >>> menu.click("Tools", "Macros", "Visual Basic Editor")
        """
        self._check_thread()
        if not item_path:
            raise GUIOperationError(
                "click() requires at least one item caption in item_path."
            )
        logger.debug("MenuProxy.click: bar=%r, path=%r", bar_name, item_path)
        try:
            bars = self._command_bars()
            bar = bars(bar_name)
        except Exception as exc:
            raise GUIOperationError(
                f"CommandBar {bar_name!r} not found: {exc}", cause=exc
            ) from exc

        # Traverse the control path, descending into sub-controls.
        current_controls: Any = bar.Controls
        for caption in item_path[:-1]:
            control = _find_control(current_controls, caption)
            if control is None:
                raise GUIOperationError(
                    f"Control {caption!r} not found in bar {bar_name!r} "
                    f"at caption path {item_path!r}."
                )
            # Sub-menu controls expose a .Controls collection.
            ctrl: Any = control
            try:
                current_controls = ctrl.Controls
            except Exception as exc:
                raise GUIOperationError(
                    f"Control {caption!r} has no sub-controls "
                    f"(cannot descend further in path {item_path!r}): {exc}",
                    cause=exc,
                ) from exc

        final_caption = item_path[-1]
        final_control = _find_control(current_controls, final_caption)
        if final_control is None:
            raise GUIOperationError(
                f"Final control {final_caption!r} not found "
                f"in bar {bar_name!r} at path {item_path!r}."
            )
        logger.debug("MenuProxy.click: executing control %r", final_caption)
        final_ctrl: Any = final_control
        try:
            final_ctrl.Execute()
        except Exception as exc:
            raise GUIOperationError(
                f"Execute() failed for control {final_caption!r} "
                f"in bar {bar_name!r}: {exc}",
                cause=exc,
            ) from exc

    @wrap_com_error
    def list_bars(self) -> list[str]:
        """Return the names of all CommandBars registered with Excel.

        Iterates the ``Application.CommandBars`` collection and collects
        each bar's ``.Name`` attribute. Bars with no name are skipped.

        Returns:
            list[str]: Sorted list of CommandBar names.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the COM collection cannot be iterated.

        Example:
            >>> menu.list_bars()
            ['3-D Settings', 'Borders', 'Cell', ...]
        """
        self._check_thread()
        logger.debug("MenuProxy.list_bars")
        try:
            bars = self._command_bars()
            names: list[str] = []
            for i in range(1, bars.Count + 1):
                try:
                    name: str = bars(i).Name
                    if name:
                        names.append(name)
                except Exception as exc:
                    # Individual bars may be inaccessible (e.g., protected
                    # add-in toolbars). Skip and log at DEBUG.
                    logger.debug(
                        "list_bars: skipping bar %d (inaccessible): %s", i, exc
                    )
                    continue
            return sorted(names)
        except GUIOperationError:
            raise
        except Exception as exc:
            raise GUIOperationError(
                f"Failed to list CommandBars: {exc}", cause=exc
            ) from exc

    @wrap_com_error
    def list_controls(self, bar_name: str) -> list[str]:
        """Return the captions of all top-level controls in a CommandBar.

        Iterates the ``.Controls`` collection of the named bar. Controls
        that have no ``Caption`` (e.g. separator lines) are skipped.

        Args:
            bar_name: The name of the CommandBar to inspect.

        Returns:
            list[str]: Captions of the bar's top-level controls, in order.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the bar cannot be found or the controls
                collection cannot be iterated.

        Example:
            >>> menu.list_controls("Standard")
            ['New', 'Open', 'Save', ...]
        """
        self._check_thread()
        logger.debug("MenuProxy.list_controls: bar=%r", bar_name)
        try:
            bars = self._command_bars()
            bar = bars(bar_name)
        except Exception as exc:
            raise GUIOperationError(
                f"CommandBar {bar_name!r} not found: {exc}", cause=exc
            ) from exc
        try:
            controls = bar.Controls
            captions: list[str] = []
            for i in range(1, controls.Count + 1):
                try:
                    caption: str = controls(i).Caption
                    if caption:
                        captions.append(caption)
                except Exception as exc:
                    logger.debug(
                        "list_controls: skipping control %d (inaccessible): %s", i, exc
                    )
                    continue
            return captions
        except GUIOperationError:
            raise
        except Exception as exc:
            raise GUIOperationError(
                f"Failed to list controls for bar {bar_name!r}: {exc}", cause=exc
            ) from exc


# ///////////////////////////////////////////////////////////////
# MODULE-LEVEL HELPERS
# ///////////////////////////////////////////////////////////////


def _find_control(controls: Any, caption: str) -> Any | None:
    """Find the first control whose caption matches ``caption`` (case-insensitive).

    Args:
        controls: A COM ``Controls`` collection object.
        caption: The caption string to search for.

    Returns:
        The matching COM control object, or ``None`` if not found.
    """
    # Strip & (accelerator key markers) from captions before comparing.
    target = caption.replace("&", "").strip().lower()
    try:
        count: int = controls.Count
    except Exception:
        return None
    for i in range(1, count + 1):
        try:
            ctrl = controls(i)
            ctrl_caption: str = ctrl.Caption.replace("&", "").strip().lower()
            if ctrl_caption == target:
                return ctrl
        except Exception as exc:
            logger.debug(
                "_find_control: skipping control %d (inaccessible): %s", i, exc
            )
            continue
    return None
