"""v2: a real-Excel backend via pywin32/win32com.client COM automation.

Implements the same vba_bridge.backends.base.Backend interface as
LibreOfficeBackend, so VBASession and wrapper.py need no changes to use it.
Windows + a local Excel install + pywin32 only -- see README.md's "v2: real
Excel via pywin32" section for setup, and basic_runtime.py's module
docstring for the Core+Agent workbook design this relies on.

Written to mirror the LibreOffice backend closely, but not yet run against
real Excel -- treat it as a first draft to verify on-machine, not a
finished, battle-tested backend the way LibreOfficeBackend is.
"""

from vba_bridge.backends.excel_com.backend import ExcelComBackend

__all__ = ["ExcelComBackend"]
