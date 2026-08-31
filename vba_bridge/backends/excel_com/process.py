"""Excel.Application process lifecycle: launch a fresh, isolated instance,
track it by PID, and be able to force-kill it if a COM call wedges.

Mirrors vba_bridge.backends.libreoffice.process.SofficeProcess in spirit:
launch() gives back something the rest of the backend can use immediately,
and terminate() is defensive about the underlying process actually being
gone rather than trusting a handle.

Uses DispatchEx (not Dispatch) deliberately: Dispatch can silently attach to
an Excel instance the user already has open and is working in -- DispatchEx
always creates a brand-new, separate Excel.Application process, which is
what we want (isolated from anything else running on the machine, and safe
to force-kill without touching the user's own Excel session).
"""

import os
import subprocess

from vba_bridge.backends.excel_com.connection import import_win32com
from vba_bridge.exceptions import LaunchTimeoutError, VbomAccessDeniedError


class ExcelProcess:
    def __init__(self, *, visible: bool = True, launch_timeout: float = 30.0):
        self.visible = visible
        self.launch_timeout = launch_timeout
        self.application = None
        self.pid = None
        self._win32com = None
        self._pythoncom = None

    def launch(self):
        """Start a fresh Excel.Application, verify VBA project object model
        access is trusted, and return the Application COM object."""
        win32com, pythoncom = import_win32com()
        self._win32com = win32com
        self._pythoncom = pythoncom

        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = self.visible
        app.DisplayAlerts = False
        app.ScreenUpdating = self.visible

        try:
            self.pid = _get_process_id(app)
        except Exception:
            self.pid = None

        self.application = app
        self._verify_vbom_trusted(app)
        return app

    @staticmethod
    def _verify_vbom_trusted(app) -> None:
        """VBComponents access raises a COM error (not a Python-catchable
        AttributeError) when "Trust access to the VBA project object model"
        is off. Probe it now, on a throwaway workbook, so the failure is a
        clear actionable message instead of a cryptic COM error surfacing
        later from deep inside module injection.
        """
        wb = app.Workbooks.Add()
        try:
            _ = wb.VBProject.VBComponents.Count
        except Exception as exc:
            raise VbomAccessDeniedError(
                "Excel refused programmatic access to the VBA project object "
                "model. Enable it once: File > Options > Trust Center > "
                "Trust Center Settings > Macro Settings > check \"Trust "
                "access to the VBA project object model\", then restart "
                "Excel and try again."
            ) from exc
        finally:
            wb.Close(SaveChanges=False)

    @property
    def is_running(self) -> bool:
        if self.application is None:
            return False
        try:
            # Cheap property read; raises once the COM server is actually gone.
            _ = self.application.Version
            return True
        except Exception:
            return False

    def terminate(self, timeout: float = 5.0) -> None:
        if self.application is not None:
            try:
                for wb in list(self.application.Workbooks):
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                self.application.Quit()
            except Exception:
                pass
        self.application = None

        if self.pid is not None:
            _hard_kill(self.pid, timeout=timeout)
        self.pid = None


def _get_process_id(app) -> int:
    """Excel.Application doesn't expose its own PID directly; Hwnd (the main
    window handle) does, via the Win32 API, resolve to one."""
    import win32process

    _, pid = win32process.GetWindowThreadProcessId(app.Hwnd)
    return pid


def _hard_kill(pid: int, timeout: float) -> None:
    """Force-kill an Excel process by PID -- the fallback for a COM call that
    wedged (no native cancel for a stuck Application.Run, same situation as
    LibreOfficeBackend's UNO invoke()). taskkill /T also takes down any
    child processes (e.g. a spawned repair/recovery dialog host)."""
    try:
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=timeout,
        )
    except Exception:
        # Best-effort: if taskkill itself isn't available or errors, there's
        # no further fallback worth adding here -- the caller already treats
        # the backend as dead and will reconnect (a fresh DispatchEx) next call.
        pass
