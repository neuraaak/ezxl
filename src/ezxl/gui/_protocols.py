# ///////////////////////////////////////////////////////////////
# _protocols - Abstract backend contracts for the gui layer
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
Abstract backend contracts for the ``ezxl.gui`` layer.

Defines four :class:`abc.ABC` base classes that each GUI backend must
implement. Keeping these contracts in a separate module that imports
**only** from the standard library prevents circular imports — proxy
modules may freely import these ABCs without pulling in any ``ezxl``
internals.

Backends
--------
- :class:`AbstractRibbonBackend` — MSO ribbon command execution and state
  queries.
- :class:`AbstractMenuBackend` — Legacy CommandBar traversal and control
  execution.
- :class:`AbstractDialogBackend` — File-open, file-save, and message-box
  dialogs.
- :class:`AbstractKeysBackend` — Keystroke injection (``SendKeys``).

Usage::

    from ezxl.gui._protocols import AbstractRibbonBackend

    class MyRibbonBackend(AbstractRibbonBackend):
        def execute(self, mso_id: str) -> None: ...
        def is_enabled(self, mso_id: str) -> bool: ...
        def is_pressed(self, mso_id: str) -> bool: ...
        def is_visible(self, mso_id: str) -> bool: ...
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
from abc import ABC, abstractmethod

# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


class AbstractRibbonBackend(ABC):
    """Contract for ribbon command execution and state queries.

    Any class that implements this interface can be injected into
    :class:`~ezxl.gui.GUIProxy` as the ribbon backend, replacing the
    default COM-based implementation.

    Implementations are responsible for thread-safety; the caller
    (:class:`~ezxl.gui.GUIProxy`) does **not** perform additional
    thread checks.
    """

    # ///////////////////////////////////////////////////////////////
    # ABSTRACT METHODS
    # ///////////////////////////////////////////////////////////////

    @abstractmethod
    def execute(self, mso_id: str) -> None:
        """Execute a built-in ribbon command by its MSO identifier.

        Args:
            mso_id: MSO control identifier string
                (e.g. ``"FileSave"``, ``"Copy"``, ``"PasteValues"``).

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the MSO ID is unknown or the command
                cannot be executed in the current application state.
        """
        ...

    @abstractmethod
    def is_enabled(self, mso_id: str) -> bool:
        """Return whether a ribbon command is currently enabled.

        Args:
            mso_id: MSO control identifier string.

        Returns:
            bool: ``True`` if the command is enabled; ``False`` otherwise.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the MSO ID is unknown or the query fails.
        """
        ...

    @abstractmethod
    def is_pressed(self, mso_id: str) -> bool:
        """Return whether a ribbon toggle command is currently pressed.

        Args:
            mso_id: MSO control identifier string.

        Returns:
            bool: ``True`` if the toggle command is in the pressed/active
                state; ``False`` if not pressed or if the control does not
                support the pressed-state query.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
        """
        ...

    @abstractmethod
    def is_visible(self, mso_id: str) -> bool:
        """Return whether a ribbon command is currently visible.

        Args:
            mso_id: MSO control identifier string.

        Returns:
            bool: ``True`` if the command is visible in the current ribbon
                state; ``False`` otherwise.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the MSO ID is unknown or the query fails.
        """
        ...


class AbstractMenuBackend(ABC):
    """Contract for legacy CommandBar traversal and control execution.

    Any class that implements this interface can be injected into
    :class:`~ezxl.gui.GUIProxy` as the menu backend, replacing the
    default COM-based implementation.
    """

    # ///////////////////////////////////////////////////////////////
    # ABSTRACT METHODS
    # ///////////////////////////////////////////////////////////////

    @abstractmethod
    def click(self, bar_name: str, *item_path: str) -> None:
        """Traverse a CommandBar by caption path and execute the final control.

        Args:
            bar_name: The name of the CommandBar to start from.
            *item_path: One or more control captions forming the path to
                the target control. At least one caption is required.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the bar, any intermediate control, or
                the final control cannot be found, or if execution fails.
        """
        ...

    @abstractmethod
    def list_bars(self) -> list[str]:
        """Return the names of all CommandBars registered with Excel.

        Returns:
            list[str]: Sorted list of CommandBar names.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the COM collection cannot be iterated.
        """
        ...

    @abstractmethod
    def list_controls(self, bar_name: str) -> list[str]:
        """Return the captions of all top-level controls in a CommandBar.

        Args:
            bar_name: The name of the CommandBar to inspect.

        Returns:
            list[str]: Captions of the bar's top-level controls, in order.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the bar cannot be found or the controls
                collection cannot be iterated.
        """
        ...


class AbstractDialogBackend(ABC):
    """Contract for file-open, file-save, and message-box dialogs.

    Any class that implements this interface can be injected into
    :class:`~ezxl.gui.GUIProxy` as the dialog backend, replacing the
    default COM/Win32-based implementation.

    Default parameter values in implementing methods must match those
    declared on :class:`~ezxl.gui.DialogProxy`.
    """

    # ///////////////////////////////////////////////////////////////
    # ABSTRACT METHODS
    # ///////////////////////////////////////////////////////////////

    @abstractmethod
    def get_file_open(
        self,
        title: str = "Open",
        initial_dir: str | None = None,
        filter: str = "Excel Files (*.xls*), *.xls*",
    ) -> str | None:
        """Show a file-open picker dialog and return the selected path.

        Args:
            title: Dialog title bar text. Defaults to ``"Open"``.
            initial_dir: Directory to open the dialog in. If ``None``,
                the backend chooses the initial directory.
            filter: File-type filter string. Defaults to Excel files.

        Returns:
            str | None: Absolute path chosen by the user, or ``None``
                if the dialog was cancelled.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the underlying call fails.
        """
        ...

    @abstractmethod
    def get_file_save(
        self,
        title: str = "Save As",
        initial_dir: str | None = None,
        filter: str = "Excel Files (*.xlsx), *.xlsx",
    ) -> str | None:
        """Show a file-save picker dialog and return the selected path.

        Args:
            title: Dialog title bar text. Defaults to ``"Save As"``.
            initial_dir: Directory to open the dialog in. If ``None``,
                the backend chooses the initial directory.
            filter: File-type filter string. Defaults to ``.xlsx``.

        Returns:
            str | None: Absolute path chosen by the user, or ``None``
                if the dialog was cancelled.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the underlying call fails.
        """
        ...

    @abstractmethod
    def alert(self, message: str, title: str = "EzXl") -> None:
        """Display a modal information message box.

        Args:
            message: The body text displayed in the message box.
            title: Caption for the message box title bar.
                Defaults to ``"EzXl"``.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the underlying call fails.
        """
        ...


class AbstractKeysBackend(ABC):
    """Contract for keystroke injection into Excel.

    Any class that implements this interface can be injected into
    :class:`~ezxl.gui.GUIProxy` as the keys backend, replacing the
    default COM-based implementation.
    """

    # ///////////////////////////////////////////////////////////////
    # ABSTRACT METHODS
    # ///////////////////////////////////////////////////////////////

    @abstractmethod
    def send_keys(self, keys: str, wait: bool = True) -> None:
        """Send a keystroke sequence to the Excel Application window.

        Args:
            keys: Keystroke string in VBA SendKeys notation
                (e.g. ``"{ENTER}"``, ``"^s"`` for Ctrl+S).
            wait: If ``True``, block until Excel processes the keystrokes
                before returning. Defaults to ``True``.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            COMOperationError: If the keystroke injection call fails.
        """
        ...
