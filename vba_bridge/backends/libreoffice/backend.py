import threading
from typing import Any, Sequence

from vba_bridge.backends.base import Backend, RawRunResult
from vba_bridge.backends.libreoffice import connection
from vba_bridge.backends.libreoffice.basic_runtime import BasicRuntime
from vba_bridge.backends.libreoffice.process import SofficeProcess
from vba_bridge.exceptions import BridgeDisconnectedError, RunTimeoutError


class LibreOfficeBackend(Backend):
    def __init__(self, *, soffice_binary: str = "soffice", launch_timeout: float = 30.0):
        self.soffice_binary = soffice_binary
        self.launch_timeout = launch_timeout
        self._process = None
        self._runtime = None

    def connect(self) -> None:
        self._process = SofficeProcess(self.soffice_binary)
        port = self._process.launch()
        ctx = connection.resolve_context(port, timeout=self.launch_timeout)
        self._runtime = BasicRuntime(ctx)
        self._runtime.setup()

    def inject_module(self, module_name: str, source: str) -> None:
        self._require_connected()
        self._runtime.inject_module(module_name, source)

    def run_macro(
        self,
        module_name: str,
        entry_point: str,
        args: Sequence[Any],
        *,
        timeout: float,
    ) -> RawRunResult:
        self._require_connected()

        result_box = {}

        def target():
            try:
                result_box["result"] = self._runtime.run(module_name)
            except Exception as exc:  # noqa: BLE001
                result_box["exception"] = exc

        worker = threading.Thread(target=target, daemon=True)
        worker.start()
        worker.join(timeout=timeout)

        if worker.is_alive():
            # The underlying UNO invoke() has no native timeout / cancel; kill
            # the process out from under the stuck worker thread (it's a
            # daemon thread, safe to abandon) and mark disconnected so the
            # next call reconnects rather than reusing a wedged runtime.
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
        # State lives entirely in PyBridge.Core globals + the AgentCode.Main
        # module, both fully replaced/reset on every run() already; nothing
        # persists across calls that reset() needs to separately clear.

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
                "LibreOfficeBackend is not connected; call connect() first "
                "(or the backend was recycled after a timeout -- create/reconnect)."
            )
