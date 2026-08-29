import abc
from dataclasses import dataclass
from typing import Any, Sequence


@dataclass
class RawRunResult:
    """What a Backend.run_macro() call reports back, before VBASession wraps it
    into a VBAResult. Backend-agnostic shape shared by LibreOffice and (in
    future) Excel/COM backends.
    """

    success: bool
    output: list
    return_value: Any = None
    err_number: int = 0
    err_description: str = ""
    err_source: str = ""


class Backend(abc.ABC):
    """The interface VBASession drives. A concrete backend owns getting a real
    VBA-compatible runtime up, injecting a module's source, running an entry
    point in it, and reporting back a RawRunResult -- all backend-specific
    mechanics (COM vs UNO, Excel vs LibreOffice) stay behind this boundary.
    """

    @abc.abstractmethod
    def connect(self) -> None:
        """Start/attach the underlying runtime and get it ready to run code."""

    @abc.abstractmethod
    def inject_module(self, module_name: str, source: str, *, is_class: bool = False) -> None:
        """Replace (or create) a code module's full source. is_class marks it
        as a VBA class module (supporting `New module_name` instantiation)
        rather than a plain standard module."""

    @abc.abstractmethod
    def run_macro(
        self,
        module_name: str,
        entry_point: str,
        args: Sequence[Any],
        *,
        timeout: float,
        run_token: str = None,
    ) -> RawRunResult:
        """Run `entry_point` (already wrapped with error handling by the
        caller) inside `module_name` and report the outcome. run_token, when
        given, is a value the caller expects to see echoed back by a
        successful run; a backend that can silently report on a run that
        never actually executed (LibreOffice/UNO can -- see StaleRunError)
        should verify it and raise rather than return stale data. Backends
        without that failure mode may ignore it."""

    @abc.abstractmethod
    def reset(self) -> None:
        """Clear any shared run-state without tearing down the process."""

    @abc.abstractmethod
    def shutdown(self) -> None:
        """Tear down the underlying runtime and release all resources."""

    @property
    @abc.abstractmethod
    def is_alive(self) -> bool:
        """Whether the backend is connected and able to run code right now."""
