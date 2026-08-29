"""v2 placeholder: a real-Excel backend via pywin32/win32com.client COM automation.

Not implemented yet. When built, ExcelComBackend will implement the same
vba_bridge.backends.base.Backend interface as LibreOfficeBackend, so
VBASession and wrapper.py need no changes to use it -- only module injection
(VBComponents.Add + CodeModule.AddFromString) and invocation (Application.Run)
are backend-specific.
"""

from vba_bridge.backends.base import Backend


class ExcelComBackend(Backend):
    def __init__(self, *args, **kwargs):
        raise NotImplementedError(
            "ExcelComBackend (real Excel via pywin32) is not implemented yet. "
            "Use vba_bridge.backends.libreoffice.LibreOfficeBackend for now."
        )

    def connect(self) -> None:
        raise NotImplementedError

    def inject_module(self, module_name: str, source: str, *, is_class: bool = False) -> None:
        raise NotImplementedError

    def run_macro(self, module_name, entry_point, args, *, timeout, run_token=None):
        raise NotImplementedError

    def reset(self) -> None:
        raise NotImplementedError

    def shutdown(self) -> None:
        raise NotImplementedError

    @property
    def is_alive(self) -> bool:
        return False
