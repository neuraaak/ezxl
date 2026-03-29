# ///////////////////////////////////////////////////////////////
# _excel_app - ExcelApp COM lifecycle manager
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
ExcelApp — entry point for all COM automation sessions.

Manages the full COM lifecycle of an Excel Application object: dispatch
(create a new Excel instance) or attach (bind to an already-running one).
Implements the context manager protocol for deterministic cleanup.

All COM calls are guarded by thread-identity assertions and wrapped via
``_com_utils.wrap_com_error`` so that ``pywintypes.com_error`` never
escapes this module.

Threading note:
    Excel COM operates under the Single-Threaded Apartment (STA) model.
    ``ExcelApp`` is NOT thread-safe. Create one instance per thread, and
    never share an instance across threads. The thread identity is recorded
    at construction time and enforced on every COM call.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
import threading
from pathlib import Path
from typing import TYPE_CHECKING, Any, Literal

# Third-party imports
import win32com.client as win32
from ezplog.lib_mode import get_logger, get_printer

# Local imports
from ..exceptions import ExcelNotAvailableError
from ..utils._com_utils import assert_main_thread, wait_until_ready, wrap_com_error

if TYPE_CHECKING:
    from ..gui._gui_proxy import GUIProxy
    from ._workbook import WorkbookProxy

# ///////////////////////////////////////////////////////////////
# CONSTANTS
# ///////////////////////////////////////////////////////////////

logger = get_logger(__name__)
printer = get_printer()

# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


class ExcelApp:
    """COM automation session for a single Excel Application instance.

    Provides a unified interface for opening, navigating, and controlling
    Excel via COM, regardless of whether the Application object was
    dispatched (new process) or attached (existing process).

    Threading:
        This class is **not thread-safe**. Excel COM uses the STA model.
        All method calls must originate from the thread that constructed
        the instance. Calls from other threads raise
        ``ExcelThreadViolationError`` immediately.

    Lifecycle
    ---------
    Two usage patterns are supported:

    **Context manager** (recommended for short, bounded sessions) — cleanup
    is automatic on exit::

        with ExcelApp(mode="dispatch", visible=False) as xl:
            wb = xl.open("C:/reports/budget.xlsx")
            wb.save()
        # Excel quit automatically

    **Manual lifecycle** (long-running sessions, e.g. consumer libraries) —
    call :meth:`quit` explicitly when done.  In ``attach`` mode, ``quit``
    is optional; Excel stays running and the Python reference is released::

        xl = ExcelApp(mode="attach")
        wb = xl.open("C:/reports/budget.xlsx")
        # ... many operations spread across multiple calls ...
        xl.quit()   # optional in attach mode — Excel stays running

    The COM object is resolved **lazily** on the first public method call;
    constructing an ``ExcelApp`` instance does not connect to COM.

    Args:
        mode: ``"dispatch"`` to start a new Excel instance, or
            ``"attach"`` to bind to the already-running one.
        visible: Whether to make the Excel window visible. Only relevant
            in ``dispatch`` mode; ignored when attaching.

    Raises:
        ExcelNotAvailableError: If ``mode="attach"`` and no Excel instance
            is currently running.
    """

    # ///////////////////////////////////////////////////////////////
    # INIT
    # ///////////////////////////////////////////////////////////////

    def __init__(
        self,
        mode: Literal["dispatch", "attach"] = "dispatch",
        visible: bool = True,
    ) -> None:
        self._mode: Literal["dispatch", "attach"] = mode
        self._visible: bool = visible
        # COM object resolved lazily on first access via _get_app().
        self._app: Any = None
        # Record the creating thread — enforced on every COM call.
        self._thread_id: int = threading.get_ident()
        logger.debug(
            "ExcelApp created (mode=%s, visible=%s, thread=%d)",
            mode,
            visible,
            self._thread_id,
        )

    # ///////////////////////////////////////////////////////////////
    # CONTEXT MANAGER
    # ///////////////////////////////////////////////////////////////

    def __enter__(self) -> ExcelApp:
        """Resolve the COM object and return self.

        Returns:
            ExcelApp: This instance, ready for use.
        """
        self._get_app()
        return self

    def __exit__(self, *_args: Any) -> None:
        """Clean up the COM session.

        In ``dispatch`` mode, quits Excel without saving. In ``attach``
        mode, leaves Excel running and merely releases the Python reference.
        """
        if self._mode == "dispatch":
            self.quit(save_changes=False)
        else:
            logger.debug("ExcelApp detaching (attach mode — Excel left running).")
            printer.system("ExcelApp detached (attach mode — Excel left running).")
            self._app = None

    # ///////////////////////////////////////////////////////////////
    # PRIVATE HELPERS
    # ///////////////////////////////////////////////////////////////

    def _get_app(self) -> Any:
        """Resolve and return the underlying COM Application object.

        Performs lazy initialisation: the first call dispatches or attaches.
        Subsequent calls return the cached object.

        Returns:
            The win32com Application COM object.

        Raises:
            ExcelNotAvailableError: If pywin32 is not installed or the COM
                object cannot be obtained.
        """
        if self._app is not None:
            return self._app

        if self._mode == "dispatch":
            try:
                self._app = win32.Dispatch("Excel.Application")
                self._app.Visible = self._visible
                logger.debug("Excel dispatched (Visible=%s).", self._visible)
                printer.system(
                    f"Excel.Application dispatched (Visible={self._visible})."
                )
            except Exception as exc:
                raise ExcelNotAvailableError(
                    f"Failed to dispatch Excel.Application: {exc}", cause=exc
                ) from exc
        else:  # attach
            try:
                self._app = win32.GetActiveObject("Excel.Application")
                logger.debug("Attached to running Excel instance.")
                printer.detect("Attached to running Excel instance.")
            except Exception as exc:
                raise ExcelNotAvailableError(
                    "No running Excel instance found. "
                    "Start Excel before using mode='attach'.",
                    cause=exc,
                ) from exc

        return self._app

    def _check_thread(self) -> None:
        """Assert the call originates from the COM apartment thread."""
        assert_main_thread(self._thread_id)

    # ///////////////////////////////////////////////////////////////
    # PUBLIC METHODS
    # ///////////////////////////////////////////////////////////////

    @wrap_com_error
    def open(self, path: str | Path) -> WorkbookProxy:
        """Open a workbook file and return a proxy for it.

        Args:
            path: Absolute path to the workbook file.

        Returns:
            WorkbookProxy: A proxy bound to the opened workbook.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            COMOperationError: If Excel cannot open the file.

        Example:
            >>> wb = xl.open("C:/data/report.xlsx")
        """
        from ._workbook import WorkbookProxy  # local import avoids circular dep

        self._check_thread()
        resolved = Path(path).resolve()
        logger.debug("Opening workbook: %s", resolved)
        xl = self._get_app()
        xl.Workbooks.Open(str(resolved))
        # The newly opened workbook becomes the ActiveWorkbook.
        wb_name: str = xl.ActiveWorkbook.Name
        logger.debug("Workbook opened: %s", wb_name)
        printer.system(f"Workbook opened: {wb_name}")
        return WorkbookProxy(self, wb_name)

    @wrap_com_error
    def workbook(self, name: str | None = None) -> WorkbookProxy:
        """Return a proxy for an already-open workbook.

        Args:
            name: The workbook name as displayed in Excel's title bar
                (e.g. ``"report.xlsx"``). Pass ``None`` to get the active
                workbook.

        Returns:
            WorkbookProxy: A proxy bound to the named workbook.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            WorkbookNotFoundError: If no workbook with that name is open.
            COMOperationError: If the COM call fails.

        Example:
            >>> wb = xl.workbook("report.xlsx")
        """
        from ._workbook import WorkbookProxy

        self._check_thread()
        xl = self._get_app()

        if name is None:
            wb_name: str = xl.ActiveWorkbook.Name
            logger.debug("Using active workbook: %s", wb_name)
            return WorkbookProxy(self, wb_name)

        # Validate the name exists before returning the proxy, so callers
        # get a clean WorkbookNotFoundError rather than a late COM fault.
        from ..exceptions import WorkbookNotFoundError

        for i in range(1, xl.Workbooks.Count + 1):
            if xl.Workbooks(i).Name == name:
                logger.debug("Resolved workbook by name: %s", name)
                return WorkbookProxy(self, name)

        raise WorkbookNotFoundError(
            f"No open workbook named '{name}'. "
            f"Open workbooks: {[xl.Workbooks(i).Name for i in range(1, xl.Workbooks.Count + 1)]}"
        )

    @wrap_com_error
    def run_macro(self, name: str, *args: Any) -> Any:
        """Execute a VBA macro by name with optional arguments.

        Args:
            name: Fully qualified macro name (e.g. ``"Module1.MyMacro"`` or
                ``"'report.xlsm'!Module1.MyMacro"``).
            *args: Positional arguments forwarded to the macro.

        Returns:
            Any: The return value of the macro, if any.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            COMOperationError: If the macro call fails.

        Example:
            >>> xl.run_macro("Module1.FormatSheet", "Sheet1")
        """
        self._check_thread()
        logger.debug("Running macro: %s (args=%s)", name, args)
        xl = self._get_app()
        return xl.Run(name, *args)

    @wrap_com_error
    def execute_ribbon(self, mso_id: str) -> None:
        """Execute a built-in ribbon command by its MSO identifier.

        Uses ``Application.CommandBars.ExecuteMso`` to trigger any built-in
        Excel ribbon button programmatically without navigating a menu tree.

        Args:
            mso_id: The MSO control identifier string
                (e.g. ``"FileSave"``, ``"Copy"``, ``"PasteValues"``).

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            COMOperationError: If the MSO ID is unknown or execution fails.

        Example:
            >>> xl.execute_ribbon("FileSave")
        """
        self._check_thread()
        logger.debug("Executing ribbon command: %s", mso_id)
        xl = self._get_app()
        xl.CommandBars.ExecuteMso(mso_id)

    @wrap_com_error
    def wait_ready(self, timeout: float = 30.0) -> None:
        """Block until Excel reports it is ready.

        Delegates to ``_com_utils.wait_until_ready``. Useful after
        operations that trigger asynchronous recalculation or file I/O.

        Args:
            timeout: Maximum seconds to wait. Defaults to 30.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            COMOperationError: If the timeout is exceeded.

        Example:
            >>> xl.wait_ready(timeout=60.0)
        """
        self._check_thread()
        logger.debug("Waiting for Excel to become ready (timeout=%.1fs).", timeout)
        wait_until_ready(self._get_app(), timeout)

    @property
    def gui(self) -> GUIProxy:
        """Access GUI interaction helpers (ribbon, menus, dialogs, keys).

        Returns a ``GUIProxy`` that bundles all GUI automation surfaces.
        The proxy is created fresh on each access; it holds no state of
        its own beyond a reference to this ``ExcelApp`` instance.

        Surfaces available via the proxy:

        - ``gui.ribbon``         — MSO ribbon execution and state queries.
        - ``gui.menu``           — Legacy CommandBar traversal and execution.
        - ``gui.dialog``         — File-open, file-save, and alert dialogs.
        - ``gui.send_keys(…)``   — ``Application.SendKeys`` pass-through.

        Returns:
            GUIProxy: Facade bound to this ``ExcelApp`` instance.

        Raises:
            ExcelThreadViolationError: If accessed from the wrong thread.

        Example:
            >>> with ExcelApp(mode="attach") as xl:
            ...     xl.gui.ribbon.execute("FileSave")
            ...     path = xl.gui.dialog.get_file_open()
        """
        from ..gui._gui_proxy import GUIProxy  # local import avoids circular dep

        self._check_thread()
        return GUIProxy(self)

    @property
    @wrap_com_error
    def hwnd(self) -> int:
        """Win32 window handle of this Excel Application instance.

        Returns the ``Application.Hwnd`` COM property. Used to bind
        pywinauto backends to the exact same Excel window managed by
        this ``ExcelApp`` session, preventing cross-instance interference
        when multiple Excel processes are running simultaneously.

        Returns:
            int: The Win32 HWND of the Excel main window.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            COMOperationError: If the COM call fails.

        Example:
            >>> hwnd = xl.hwnd
            >>> gui = GUIProxy(xl, keys=PywinautoKeysBackend(hwnd=hwnd))
        """
        self._check_thread()
        return int(self._get_app().Hwnd)

    @wrap_com_error
    def quit(self, save_changes: bool = False) -> None:
        """Quit the Excel Application.

        This method is safe to call multiple times; subsequent calls after
        the first are no-ops.

        Args:
            save_changes: If ``True``, Excel will prompt to save unsaved
                workbooks. If ``False``, all unsaved changes are discarded.
                Defaults to ``False``.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            COMOperationError: If the quit call fails unexpectedly.

        Example:
            >>> xl.quit(save_changes=False)
        """
        if self._app is None:
            return
        self._check_thread()
        logger.debug("Quitting Excel (save_changes=%s).", save_changes)
        try:
            self._app.Quit()
            printer.system("Excel.Application quit.")
        except Exception as exc:
            # Swallow errors on quit — the process may already be gone.
            logger.warning("Error during Excel quit (may be harmless): %s", exc)
        finally:
            self._app = None
