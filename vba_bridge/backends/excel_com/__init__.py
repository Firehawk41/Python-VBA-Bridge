"""v2: a real-Excel backend via pywin32/win32com.client COM automation.

Implements the same vba_bridge.backends.base.Backend interface as
LibreOfficeBackend, so VBASession and wrapper.py need no changes to use it.
Windows + a local Excel install + pywin32 only -- see README.md's "v2: real
Excel via pywin32" section for setup, and basic_runtime.py's module
docstring for the Core+Agent workbook design this relies on.

Verified against real Excel. See TESTING_REAL_PROJECTS.md for the recipe
covering common real-world gotchas (external library references,
UserForms, isolating from production config, etc.) when bringing an
existing, already-deployed VBA project under test.
"""

from vba_bridge.backends.excel_com.backend import (
    MICROSOFT_ACTIVEX_DATA_OBJECTS,
    MICROSOFT_SCRIPTING_RUNTIME,
    ExcelComBackend,
)

__all__ = [
    "ExcelComBackend",
    "MICROSOFT_SCRIPTING_RUNTIME",
    "MICROSOFT_ACTIVEX_DATA_OBJECTS",
]
