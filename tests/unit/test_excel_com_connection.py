import sys

import pytest

from vba_bridge.backends.excel_com.connection import import_win32com
from vba_bridge.exceptions import ExcelComNotAvailableError


def test_import_win32com_rejects_non_windows(monkeypatch):
    monkeypatch.setattr(sys, "platform", "linux")
    with pytest.raises(ExcelComNotAvailableError, match="requires Windows"):
        import_win32com()


def test_import_win32com_reports_missing_pywin32(monkeypatch):
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setitem(sys.modules, "win32com.client", None)
    monkeypatch.setitem(sys.modules, "win32com", None)
    monkeypatch.setitem(sys.modules, "pythoncom", None)
    with pytest.raises(ExcelComNotAvailableError, match="pywin32"):
        import_win32com()
