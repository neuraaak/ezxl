# ///////////////////////////////////////////////////////////////
# _registry - UIA element descriptor registry for pywinauto backends
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""
UIA element descriptor registry for pywinauto-based backends.

Defines :class:`UIElementSpec`, a frozen dataclass describing a single
navigable Excel UI element, and the canonical ``BACKSTAGE_ELEMENTS`` dict
that maps stable keys to their descriptors.

Design notes
------------
- UIA direct click is the **primary** navigation strategy.  The backend
  locates the ``Button "Onglet Fichier"`` (or locale equivalent) to open
  the Backstage, then clicks the target ``ListItem`` by its UIA ``Name``.
  This approach requires no keyboard focus and is robust against window
  focus loss.
- ``alt_sequence`` is kept as a **fallback** for environments where UIA
  click fails (e.g. remote desktop, accessibility restrictions).  It may
  be empty (``""``) for elements that have no viable Alt-sequence.
- ``names`` maps locale codes to localised UIA ``Name`` attribute values.
  Used by the primary UIA click strategy and as a secondary fallback.
- Consumer libraries extend the registry via dict merge (Python 3.9+ ``|``
  operator) rather than subclassing::

      from ezxl.gui.pywinauto._registry import UIElementSpec
      MY_ELEMENTS = {"my_action": UIElementSpec(key="my_action", alt_sequence="%XA")}
      MyBackend._ELEMENTS = PywinautoBackstageBackend._ELEMENTS | MY_ELEMENTS
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
from dataclasses import dataclass, field

# ///////////////////////////////////////////////////////////////
# LOCALISATION MAPS
# ///////////////////////////////////////////////////////////////

# Localised name of the "Browse" button inside the Save As intermediate panel
# ("Mode Backstage" Pane).  Clicking it opens the native Windows Save As dialog.
SAVE_AS_BROWSE_NAMES: dict[str, str] = {
    "en": "Browse",
    "fr": "Parcourir",
}

# Localised title of the intermediate Save As panel that appears after clicking
# the "Enregistrer sous" ListItem in the Backstage list.
SAVE_AS_PANEL_NAMES: dict[str, str] = {
    "en": "Backstage view",
    "fr": "Mode Backstage",
}

# ///////////////////////////////////////////////////////////////
# SAVE AS DIALOG — native Windows dialog (opened via "Parcourir")
# ///////////////////////////////////////////////////////////////

# Title of the native Windows Save As dialog window (child of Excel's window).
SAVE_AS_DIALOG_TITLES: dict[str, str] = {
    "en": "Save As",
    "fr": "Enregistrer sous",
}

# UIA Name of the filename ComboBox / Edit inside the Save As dialog.
SAVE_AS_FILENAME_LABEL: dict[str, str] = {
    "en": "File name:",
    "fr": "Nom de fichier :",
}

# UIA Name of the file-type ComboBox inside the Save As dialog.
SAVE_AS_TYPE_LABEL: dict[str, str] = {
    "en": "Save as type:",
    "fr": "Type :",
}

# UIA Name of the "Open" button inside the Type ComboBox (expands the list).
SAVE_AS_COMBO_OPEN_BTN: dict[str, str] = {
    "en": "Open",
    "fr": "Ouvrir",
}

# UIA Name of the "Save" / "Enregistrer" button in the Save As dialog.
SAVE_AS_SAVE_BTN: dict[str, str] = {
    "en": "Save",
    "fr": "Enregistrer",
}

# Title of the top-level Desktop Pane that hosts the Type combobox dropdown list.
# This pane is NOT under Excel's window in the UIA tree — it is a child of the
# Windows Desktop.  Its title is locale-dependent (virtual desktop name).
SAVE_AS_DESKTOP_PANE: dict[str, str] = {
    "en": "Desktop 1",
    "fr": "Bureau 1",
}

# Title of the overwrite-confirmation dialog that appears when saving over an
# existing file.  This dialog is a child Window of the "Enregistrer sous" Window.
SAVE_AS_OVERWRITE_DIALOG: dict[str, str] = {
    "en": "Confirm Save As",
    "fr": "Confirmer l'enregistrement",
}

# UIA Name of the "Yes" button in the overwrite-confirmation dialog.
SAVE_AS_OVERWRITE_YES: dict[str, str] = {
    "en": "Yes",
    "fr": "Oui",
}

# Maps file extensions to partial format strings used to match ListItems in the
# Type combobox dropdown.  Only a prefix is stored — matching is done with
# ``str.startswith`` so truncated UIA names still match.
SAVE_AS_FORMAT_BY_EXT: dict[str, dict[str, str]] = {
    ".xlsx": {
        "en": "Excel Workbook (*.xlsx)",
        "fr": "Classeur Excel (*.xlsx)",
    },
    ".xlsb": {
        "en": "Excel Binary Workbook",
        "fr": "Classeur Excel binaire",
    },
    ".xls": {
        "en": "Excel 97-2003 Workbook",
        "fr": "Classeur Excel 97 - 2003",
    },
    ".csv": {
        "en": "CSV UTF-8",
        "fr": "CSV UTF-8",
    },
    ".txt": {
        "en": "Text (Tab delimited)",
        "fr": "Texte (séparateur : tabulation",
    },
    ".pdf": {
        "en": "PDF (*.pdf)",
        "fr": "PDF (*.pdf)",
    },
    ".xml": {
        "en": "XML Data (*.xml)",
        "fr": "Données XML (*.xml)",
    },
}

FILE_BUTTON_NAMES: dict[str, str] = {
    "en": "File Tab",
    "fr": "Onglet Fichier",
}
"""Localised UIA ``Name`` of the ``Button`` that opens the Backstage.

Keys are locale codes (``"en"``, ``"fr"``). Used by
:meth:`~ezxl.gui.pywinauto._backstage.PywinautoBackstageBackend._ensure_backstage_open`
to locate and click the File tab button on the ribbon.
"""

FILE_LIST_NAMES: dict[str, str] = {
    "en": "File",
    "fr": "Fichier",
}
"""Localised UIA ``Name`` of the ``List`` control that represents the open Backstage.

Keys are locale codes (``"en"``, ``"fr"``). Used by
:meth:`~ezxl.gui.pywinauto._backstage.PywinautoBackstageBackend._is_backstage_open`
and :meth:`~ezxl.gui.pywinauto._backstage.PywinautoBackstageBackend._ensure_backstage_open`
to detect and retrieve the Backstage list.
"""

# ///////////////////////////////////////////////////////////////
# CLASSES
# ///////////////////////////////////////////////////////////////


@dataclass(frozen=True)
class UIElementSpec:
    """Locale-independent descriptor for a single UIA-navigable Excel element.

    Instances are immutable (``frozen=True``) and safe to use as dict values
    or set members.

    Args:
        key: Stable locale-independent identifier (e.g. ``"file_save"``).
            Used as the lookup key in backend element registries.
        alt_sequence: Alt-key sequence in pywinauto ``send_keys`` notation
            (e.g. ``"%fs"`` for Alt, F, S).  Used as a fallback when the
            primary UIA click strategy fails.  Defaults to ``""`` (no
            fallback sequence available).
        control_type: Expected UIA ``ControlType`` string used for UIA
            ``child_window`` searches inside the Backstage list.
            Defaults to ``"ListItem"`` — Backstage entries are ``ListItem``
            controls, not ``Button`` controls.
        names: Optional mapping of locale code to localised UIA ``Name``
            attribute value.  Used by the primary UIA click strategy and
            as a secondary fallback.  May be empty if no name-based
            resolution is needed.

    Example:
        >>> spec = UIElementSpec(
        ...     key="file_save",
        ...     alt_sequence="%fs",
        ...     control_type="ListItem",
        ...     names={"en": "Save", "fr": "Enregistrer"},
        ... )
    """

    key: str
    alt_sequence: str = ""
    control_type: str = "ListItem"
    names: dict[str, str] = field(default_factory=dict)


# ///////////////////////////////////////////////////////////////
# REGISTRY
# ///////////////////////////////////////////////////////////////

BACKSTAGE_ELEMENTS: dict[str, UIElementSpec] = {
    "file_save": UIElementSpec(
        key="file_save",
        alt_sequence="%fs",
        control_type="ListItem",
        names={"en": "Save", "fr": "Enregistrer"},
    ),
    "file_save_as": UIElementSpec(
        key="file_save_as",
        alt_sequence="%fa",
        control_type="ListItem",
        names={"en": "Save As", "fr": "Enregistrer sous"},
    ),
    "file_open": UIElementSpec(
        key="file_open",
        alt_sequence="%fo",
        control_type="ListItem",
        names={"en": "Open", "fr": "Ouvrir"},
    ),
    "file_close": UIElementSpec(
        key="file_close",
        alt_sequence="%fw",
        control_type="ListItem",
        names={"en": "Close", "fr": "Fermer"},
    ),
    "file_options": UIElementSpec(
        key="file_options",
        alt_sequence="%ft",
        control_type="ListItem",
        names={"en": "Options", "fr": "Options"},
    ),
}
"""Canonical element registry for the Excel Backstage (File menu).

Keys are stable locale-independent identifiers. Consumer libraries may
extend this dict using the ``|`` merge operator::

    from ezxl.gui.pywinauto._registry import BACKSTAGE_ELEMENTS, UIElementSpec

    MY_ELEMENTS = {"my_action": UIElementSpec(key="my_action", alt_sequence="%XA")}
    MyBackend._ELEMENTS = BACKSTAGE_ELEMENTS | MY_ELEMENTS
"""
