import time
from typing import Any, Mapping, Optional, Sequence, Union

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
        vba_source: Union[str, Mapping[str, str]],
        *,
        entry_point: Optional[str] = None,
        class_modules: Sequence[str] = (),
        args: Sequence[Any] = (),
        timeout: Optional[float] = None,
    ) -> VBAResult:
        """Run VBA/Basic code and return a structured VBAResult.

        vba_source is either a single code snippet (a str -- today's
        single-module behavior, unchanged) or a multi-module/class program
        (a dict of {module_name: source}). For the dict form, class_modules
        names which of those module_name keys are VBA class modules
        (supporting `New module_name`); entry_point may then be
        "ModuleName.ProcName" or a bare name searched across all non-class
        modules.
        """
        if not self._started or (self._auto_reconnect and not self._backend.is_alive):
            self._backend.connect()
            self._started = True

        start = time.monotonic()
        try:
            if isinstance(vba_source, str):
                modules_to_inject, resolved_entry, is_class_by_name, run_token = (
                    self._prepare_single(vba_source, entry_point=entry_point, args=args)
                )
            else:
                modules_to_inject, resolved_entry, is_class_by_name, run_token = (
                    self._prepare_program(
                        vba_source, entry_point=entry_point, class_modules=class_modules, args=args
                    )
                )
        except wrapper.EntryPointNotFoundError as exc:
            return self._exception_result(exc, start, entry_point or "")

        try:
            for module_name, source in modules_to_inject.items():
                self._backend.inject_module(
                    module_name, source, is_class=is_class_by_name.get(module_name, False)
                )
            # `args` is passed through for backends (e.g. a future Excel/COM
            # backend) that can invoke a macro with real arguments directly;
            # LibreOfficeBackend ignores it here since the source generated
            # above already baked the args into the wrapped call. run_token
            # lets a backend detect a run that silently didn't execute (see
            # StaleRunError) instead of returning a previous call's leftover
            # result.
            raw = self._backend.run_macro(
                _MODULE_NAME if isinstance(vba_source, str) else wrapper.ORCHESTRATOR_MODULE_NAME,
                resolved_entry,
                args,
                timeout=timeout if timeout is not None else self._run_timeout,
                run_token=run_token,
            )
        except Exception as exc:  # noqa: BLE001 - transport failure, not a VBA error
            return self._exception_result(exc, start, resolved_entry)

        return self._success_or_error_result(raw, start, resolved_entry)

    @staticmethod
    def _prepare_single(vba_source, *, entry_point, args):
        wrapped_source, resolved_entry, _, run_token = wrapper.wrap_module(
            vba_source, entry_point=entry_point, args=args
        )
        return {_MODULE_NAME: wrapped_source}, resolved_entry, {}, run_token

    @staticmethod
    def _prepare_program(vba_source, *, entry_point, class_modules, args):
        modules_to_inject, entry_module, entry_proc, _, run_token = wrapper.wrap_program(
            vba_source, class_modules=class_modules, entry_point=entry_point, args=args
        )
        is_class_by_name = {name: True for name in class_modules}
        return modules_to_inject, f"{entry_module}.{entry_proc}", is_class_by_name, run_token

    @staticmethod
    def _exception_result(exc, start, entry_point) -> VBAResult:
        return VBAResult(
            success=False,
            output=[],
            return_value=None,
            error=None,
            raw_exception=exc,
            duration_ms=(time.monotonic() - start) * 1000,
            entry_point=entry_point,
        )

    @staticmethod
    def _success_or_error_result(raw, start, resolved_entry) -> VBAResult:
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
