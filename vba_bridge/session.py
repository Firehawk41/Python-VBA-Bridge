import time
from typing import Any, Optional, Sequence

from vba_bridge import wrapper
from vba_bridge.backends.base import Backend
from vba_bridge.results import VBAError, VBAResult

_MODULE_NAME = "Main"


class VBASession:
    """Public, backend-agnostic entry point: write VBA source, run it, get a
    structured VBAResult back, fix, and re-run -- a REPL-style loop for VBA.
    """

    def __init__(
        self,
        backend: Optional[Backend] = None,
        *,
        auto_start: bool = True,
        auto_reconnect: bool = True,
        run_timeout: float = 30.0,
    ):
        if backend is None:
            from vba_bridge.backends.libreoffice import LibreOfficeBackend

            backend = LibreOfficeBackend()
        self._backend = backend
        self._auto_reconnect = auto_reconnect
        self._run_timeout = run_timeout
        self._started = False
        if auto_start:
            self._ensure_started()

    def _ensure_started(self) -> None:
        if not self._started:
            self._backend.connect()
            self._started = True

    def run(
        self,
        vba_source: str,
        *,
        entry_point: Optional[str] = None,
        args: Sequence[Any] = (),
        timeout: Optional[float] = None,
    ) -> VBAResult:
        if not self._started or (self._auto_reconnect and not self._backend.is_alive):
            self._backend.connect()
            self._started = True

        start = time.monotonic()
        try:
            wrapped_source, resolved_entry, _ = wrapper.wrap_module(
                vba_source, entry_point=entry_point, args=args
            )
        except wrapper.EntryPointNotFoundError as exc:
            return VBAResult(
                success=False,
                output=[],
                return_value=None,
                error=None,
                raw_exception=exc,
                duration_ms=(time.monotonic() - start) * 1000,
                entry_point=entry_point or "",
            )

        try:
            self._backend.inject_module(_MODULE_NAME, wrapped_source)
            # `args` is passed through for backends (e.g. a future Excel/COM
            # backend) that can invoke a macro with real arguments directly;
            # LibreOfficeBackend ignores it here since wrap_module() already
            # baked the args into the generated call inside wrapped_source.
            raw = self._backend.run_macro(
                _MODULE_NAME,
                resolved_entry,
                args,
                timeout=timeout if timeout is not None else self._run_timeout,
            )
        except Exception as exc:  # noqa: BLE001 - transport failure, not a VBA error
            return VBAResult(
                success=False,
                output=[],
                return_value=None,
                error=None,
                raw_exception=exc,
                duration_ms=(time.monotonic() - start) * 1000,
                entry_point=resolved_entry,
            )

        duration_ms = (time.monotonic() - start) * 1000
        error = None
        if not raw.success:
            error = VBAError(
                number=raw.err_number,
                description=raw.err_description,
                source=raw.err_source,
            )
        return VBAResult(
            success=raw.success,
            output=raw.output,
            return_value=raw.return_value,
            error=error,
            raw_exception=None,
            duration_ms=duration_ms,
            entry_point=resolved_entry,
        )

    def reset(self) -> None:
        """Clear shared run-state without tearing down the underlying process."""
        if self._started:
            self._backend.reset()

    def restart(self) -> None:
        """Recycle the underlying backend process defensively."""
        self.close()
        self._backend.connect()
        self._started = True

    def close(self) -> None:
        if self._started:
            self._backend.shutdown()
            self._started = False

    def __enter__(self) -> "VBASession":
        self._ensure_started()
        return self

    def __exit__(self, *exc_info) -> None:
        self.close()
