# ///////////////////////////////////////////////////////////////
# _protocols - Abstract backend contracts for the gui layer
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
Abstract backend contracts for the ``ezxl.gui`` layer.

Defines six :class:`abc.ABC` base classes that each GUI backend must
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
- :class:`AbstractBackstageFileOps` — Excel Backstage file operations
  (save, save_as, open_file, close_workbook) via COM.
- :class:`AbstractBackstageNavigator` — Excel Backstage visual navigation
  (open_options, open_save_as_panel, open_file, close_workbook) via UIA.

Compatibility alias
-------------------
:data:`AbstractBackstageBackend` is kept as a type alias for
:class:`AbstractBackstageFileOps` to preserve existing imports and
public-API declarations.  New code should reference
:class:`AbstractBackstageFileOps` or :class:`AbstractBackstageNavigator`
directly.

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
from typing import Any, Protocol

# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


class ExcelAppLike(Protocol):
    """Structural contract for ExcelApp-like objects used by GUI backends.

    Keeps GUI typing independent from ``ezxl.core`` to preserve layer
    boundaries enforced by import-linter.
    """

    _thread_id: int

    def _get_app(self) -> Any:
        """Return the underlying COM ``Application`` object."""
        ...


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


class AbstractBackstageFileOps(ABC):
    """Contract for Excel Backstage file operations via COM.

    Covers the four operations that the COM object model executes
    reliably, focus-independently, and locale-independently:
    ``save``, ``save_as``, ``open_file``, and ``close_workbook``.

    This contract is implemented by
    :class:`~ezxl.gui.win32com.COMBackstageBackend`.  Inject it into
    :class:`~ezxl.gui.GUIProxy` via the *backstage* parameter::

        gui = GUIProxy(xl, backstage=COMBackstageBackend(xl))
    """

    # ///////////////////////////////////////////////////////////////
    # ABSTRACT METHODS
    # ///////////////////////////////////////////////////////////////

    @abstractmethod
    def save(self) -> None:
        """Save the active workbook.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            WorkbookNotFoundError: If no workbook is currently open.
            GUIOperationError: If the action cannot be completed.
        """
        ...

    @abstractmethod
    def save_as(self, path: str | None = None) -> None:
        """Save the active workbook under a new path, or open the Save As dialog.

        Args:
            path: Absolute path for the new file, including extension.
                If ``None``, the built-in Save As dialog is displayed for
                manual path selection.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            WorkbookNotFoundError: If no workbook is currently open.
            GUIOperationError: If the panel cannot be opened or path entry
                fails.
        """
        ...

    @abstractmethod
    def open_file(self) -> None:
        """Show the built-in Excel Open dialog.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            GUIOperationError: If the dialog cannot be opened.
        """
        ...

    @abstractmethod
    def close_workbook(self) -> None:
        """Close the active workbook without saving.

        Raises:
            ExcelThreadViolationError: If called from the wrong thread.
            WorkbookNotFoundError: If no workbook is currently open.
            GUIOperationError: If the action cannot be completed.
        """
        ...


class AbstractBackstageNavigator(ABC):
    """Contract for Excel Backstage visual navigation via UI Automation.

    Covers operations that require UIA-level interaction with the
    Backstage overlay: navigating to the Options panel, opening the
    Save As panel without confirming, and UIA-driven open/close actions.

    This contract is implemented by
    :class:`~ezxl.gui.pywinauto.PywinautoBackstageBackend`.  Inject it
    into :class:`~ezxl.gui.GUIProxy` via the *backstage_nav* parameter::

        gui = GUIProxy(
            xl,
            backstage=COMBackstageBackend(xl),
            backstage_nav=PywinautoBackstageBackend(hwnd=xl.hwnd, locale="fr"),
        )

    Note:
        ``backstage_nav`` is optional — :class:`~ezxl.gui.GUIProxy` defaults
        it to ``None``.  Access it via ``xl.gui.backstage_nav``; guard with
        an ``is not None`` check before calling if it may be absent.
    """

    # ///////////////////////////////////////////////////////////////
    # ABSTRACT METHODS
    # ///////////////////////////////////////////////////////////////

    @abstractmethod
    def open_options(self) -> None:
        """Navigate to the Excel Options panel via the Backstage.

        Raises:
            GUIOperationError: If the Options panel cannot be reached.
        """
        ...

    @abstractmethod
    def open_save_as_panel(self) -> None:
        """Open the Save As panel in the Backstage without confirming a save.

        Clicks the ``"Enregistrer sous"`` (or locale equivalent) ListItem
        and leaves the panel open for manual interaction.

        Raises:
            GUIOperationError: If the panel cannot be reached.
        """
        ...

    @abstractmethod
    def open_file(self) -> None:
        """Open the Open panel via the Backstage.

        Raises:
            GUIOperationError: If the panel cannot be reached.
        """
        ...

    @abstractmethod
    def close_workbook(self) -> None:
        """Close the active workbook via a Backstage UIA click.

        Raises:
            GUIOperationError: If the action cannot be completed.
        """
        ...


# ///////////////////////////////////////////////////////////////
# COMPATIBILITY ALIAS
# ///////////////////////////////////////////////////////////////

#: Deprecated alias for :class:`AbstractBackstageFileOps`.
#:
#: Kept so that existing code importing ``AbstractBackstageBackend`` from
#: ``ezxl`` or ``ezxl.gui._protocols`` continues to work without change.
#: New code should reference :class:`AbstractBackstageFileOps` directly.
AbstractBackstageBackend = AbstractBackstageFileOps
