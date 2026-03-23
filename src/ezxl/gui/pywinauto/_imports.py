# ///////////////////////////////////////////////////////////////
# _imports - Centralized pywinauto imports with warning suppression
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""Centralized imports for the optional ``pywinauto`` dependency.

``pywinauto`` currently emits a known ``SyntaxWarning`` from
``pywinauto.keyboard`` on recent Python versions because of an invalid
escape sequence in its own source code. The warning is third-party noise,
not an ``ezxl`` issue, so we suppress that single warning at import time
while leaving all other warnings untouched.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
import warnings

# ///////////////////////////////////////////////////////////////
# OPTIONAL DEPENDENCY GUARD
# ///////////////////////////////////////////////////////////////

try:
    with warnings.catch_warnings():
        warnings.filterwarnings(
            action="ignore",
            message=r"invalid escape sequence '\\;'",
            category=SyntaxWarning,
            module=r"pywinauto\.keyboard",
        )
        from pywinauto.application import (  # type: ignore[import-untyped]
            Application,
            WindowSpecification,
        )
        from pywinauto.base_wrapper import BaseWrapper  # type: ignore[import-untyped]
        from pywinauto.keyboard import (  # type: ignore[import-untyped]
            send_keys as _pw_send_keys,
        )
except ImportError as _pwn_import_error:
    raise ImportError(
        "pywinauto is required for the pywinauto GUI backends but is not installed. "
        "Install it with: pip install pywinauto"
    ) from _pwn_import_error

__all__ = [
    "Application",
    "BaseWrapper",
    "WindowSpecification",
    "_pw_send_keys",
]
