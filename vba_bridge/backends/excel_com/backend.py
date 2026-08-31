"""Backend implementing vba_bridge.backends.base.Backend against a real,
locally-installed Excel via pywin32/COM automation. See basic_runtime.py's
module docstring for the Core+Agent workbook design, and README.md's "v2:
real Excel via pywin32" section for setup and the specific assumptions here
that still need on-machine verification -- this has been written to mirror
LibreOfficeBackend closely but has not yet been run against real Excel.
"""

import threading
from typing import Any, Sequence

from vba_bridge.backends.base import Backend, RawRunResult
from vba_bridge.backends.excel_com import basic_runtime
from vba_bridge.backends.excel_com.connection import import_win32com
from vba_bridge.backends.excel_com.process import ExcelProcess
from vba_bridge.exceptions import BridgeDisconnectedError, RunTimeoutError


class ExcelComBackend(Backend):
    def __init__(self, *, visible: bool = True, launch_timeout: float = 30.0):
        """visible=True (the default) leaves the Excel window on-screen while
        vba_bridge drives it -- deliberate: you can watch what it's doing,
        which matters more here than for the headless LibreOffice backend
        since this is your real Excel install, not a disposable sandbox
        process. Set False once you trust it, for less screen noise."""
        self.visible = visible
        self.launch_timeout = launch_timeout
        self._process = None
        self._runtime = None

    def connect(self) -> None:
        self._process = ExcelProcess(visible=self.visible, launch_timeout=self.launch_timeout)
        application = self._process.launch()
        self._runtime = basic_runtime.VBAProjectRuntime(application)
        self._runtime.setup()

    def inject_module(self, module_name: str, source: str, *, is_class: bool = False) -> None:
        self._require_connected()
        self._runtime.inject_module(module_name, source, is_class=is_class)

    def run_macro(
        self,
        module_name: str,
        entry_point: str,
        args: Sequence[Any],
        *,
        timeout: float,
        run_token: str = None,
    ) -> RawRunResult:
        self._require_connected()
        # import_win32com() returns (win32com.client, pythoncom) -- the
        # *client* submodule itself, not the top-level win32com package --
        # so win32com_client.Dispatch(...) below, not win32com_client.client.
        win32com_client, pythoncom = import_win32com()

        agent_name = self._runtime.agent_workbook_name
        core_name = self._runtime.core_workbook_name

        # The Application COM object was created on (and belongs to) this
        # thread's apartment. Running the actual call on a worker thread --
        # so a wedged VBA loop can be timed out and the process force-killed,
        # same reasoning as LibreOfficeBackend's UNO invoke() -- means that
        # thread needs its own thread-safe proxy to it, not the raw object:
        # marshal it through a stream (the standard pywin32 pattern for
        # crossing apartments) and unmarshal on the other side.
        stream = pythoncom.CoMarshalInterThreadInterfaceInStream(
            pythoncom.IID_IDispatch, self._runtime.application._oleobj_
        )

        result_box = {}

        def target():
            pythoncom.CoInitialize()
            try:
                dispatch = pythoncom.CoGetInterfaceAndReleaseStream(
                    stream, pythoncom.IID_IDispatch
                )
                app = win32com_client.Dispatch(dispatch)
                basic_runtime.run_on_application(app, agent_name, core_name, module_name)
                packed = basic_runtime.get_result_packed(app, core_name)
                result_box["result"] = basic_runtime.unpack_result(
                    module_name, packed, run_token
                )
            except Exception as exc:  # noqa: BLE001
                result_box["exception"] = exc
            finally:
                pythoncom.CoUninitialize()

        worker = threading.Thread(target=target, daemon=True)
        worker.start()
        worker.join(timeout=timeout)

        if worker.is_alive():
            # No native cancel for a stuck Application.Run; force-kill the
            # Excel process out from under the wedged worker thread (daemon,
            # safe to abandon) and mark disconnected so the next call
            # reconnects with a fresh Excel instance rather than reusing a
            # wedged one.
            stuck_process = self._process
            self._process = None
            self._runtime = None
            if stuck_process is not None:
                stuck_process.terminate()
            raise RunTimeoutError(
                f"run_macro('{module_name}', '{entry_point}') exceeded {timeout}s timeout"
            )

        if "exception" in result_box:
            raise result_box["exception"]
        return result_box["result"]

    def reset(self) -> None:
        self._require_connected()
        # Same reasoning as LibreOfficeBackend.reset(): PyBridge_Reset()
        # already fully re-initializes Core's state on every run(), and the
        # Agent workbook's modules are always fully replaced before a call,
        # not incrementally patched -- nothing else persists across calls
        # that reset() needs to separately clear.

    def shutdown(self) -> None:
        if self._runtime is not None:
            self._runtime.close()
            self._runtime = None
        if self._process is not None:
            self._process.terminate()
            self._process = None

    @property
    def is_alive(self) -> bool:
        return (
            self._process is not None
            and self._process.is_running
            and self._runtime is not None
            and self._runtime.is_alive()
        )

    def _require_connected(self) -> None:
        if self._runtime is None:
            raise BridgeDisconnectedError(
                "ExcelComBackend is not connected; call connect() first (or "
                "the backend was recycled after a timeout -- create/reconnect)."
            )
