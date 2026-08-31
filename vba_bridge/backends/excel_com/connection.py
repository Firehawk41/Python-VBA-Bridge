"""pywin32 import handling for the Excel/COM backend.

Mirrors vba_bridge.backends.libreoffice.connection.import_uno(): a plain
`import win32com.client` failure gets turned into an actionable
ExcelComNotAvailableError rather than a bare ImportError, since the two real
failure modes here are "pywin32 isn't installed" and "this isn't Windows" --
both worth a clear message rather than a stack trace from deep inside COM
plumbing.
"""

import sys

from vba_bridge.exceptions import ExcelComNotAvailableError


def import_win32com():
    """Import and return (win32com.client, pythoncom), or raise
    ExcelComNotAvailableError with an actionable message."""
    if sys.platform != "win32":
        raise ExcelComNotAvailableError(
            "ExcelComBackend requires Windows -- COM automation of a real "
            "Excel.Application is not available on this platform "
            f"({sys.platform}). Use vba_bridge.backends.libreoffice."
            "LibreOfficeBackend instead, or run this on a Windows machine "
            "with Excel installed."
        )
    try:
        import win32com.client
        import pythoncom

        return win32com.client, pythoncom
    except ImportError as exc:
        raise ExcelComNotAvailableError(
            "Could not import win32com.client / pythoncom. Install pywin32: "
            "`pip install pywin32` (or `pip install vba-bridge[excel]`). If "
            "it's already installed and this still fails, try running "
            "`python -m pywin32_postinstall -install` from an elevated "
            "prompt -- pywin32 needs a one-time registration step that its "
            "own installer sometimes skips."
        ) from exc
