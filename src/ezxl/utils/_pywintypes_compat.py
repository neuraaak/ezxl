# ///////////////////////////////////////////////////////////////
# _pywintypes_compat - Typed pywintypes runtime aliases
# Project: EzXl
# ///////////////////////////////////////////////////////////////

"""Typed compatibility aliases for ``pywintypes`` runtime members.

The ``pywintypes`` module exposes several runtime-only attributes whose
type information is incomplete in static analysis tools. This module
centralises the required casts so callers can use named aliases instead
of sprinkling local workarounds across the codebase.
"""

from __future__ import annotations

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library imports
from typing import Any, cast

# Third-party imports
import pywintypes

# ///////////////////////////////////////////////////////////////
# CONSTANTS
# ///////////////////////////////////////////////////////////////

_PYWINTYPES_ANY = cast(Any, pywintypes)

COM_TIME_TYPE: type = cast(type, _PYWINTYPES_ANY.TimeType)
COM_ERROR_TYPE: type[BaseException] = cast(type[BaseException], _PYWINTYPES_ANY.error)
COM_EXCEPTION_TYPE: type[BaseException] = cast(
    type[BaseException], _PYWINTYPES_ANY.com_error
)

__all__ = ["COM_ERROR_TYPE", "COM_EXCEPTION_TYPE", "COM_TIME_TYPE"]
