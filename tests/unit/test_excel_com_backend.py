"""Tests ExcelComBackend's own logic (connect/inject/reset/shutdown/is_alive,
and run_macro's timeout + result-unpacking) against fakes for ExcelProcess
and the win32com/pythoncom modules it imports lazily -- not against real
Excel/COM, which isn't available on this platform. See
test_excel_com_basic_runtime.py for VBAProjectRuntime's own logic, which
these tests reuse via monkeypatching rather than re-testing.
"""

import sys
import types

import pytest

from tests.unit.fakes_excel_com import FakeApplication
from vba_bridge.backends.excel_com import backend as backend_module
from vba_bridge.backends.excel_com import basic_runtime
from vba_bridge.exceptions import BridgeDisconnectedError, RunTimeoutError


class FakeExcelProcess:
    """Stands in for ExcelProcess: hands back a FakeApplication instead of
    launching real Excel, and tracks terminate() calls."""

    def __init__(self, *, visible=True, launch_timeout=30.0):
        self.visible = visible
        self.launch_timeout = launch_timeout
        self.application = FakeApplication()
        self.terminated = False
        self._running = True

    def launch(self):
        return self.application

    @property
    def is_running(self):
        return self._running and not self.terminated

    def terminate(self, timeout=5.0):
        self.terminated = True
        self._running = False


@pytest.fixture
def patched_process(monkeypatch):
    monkeypatch.setattr(backend_module, "ExcelProcess", FakeExcelProcess)
    return FakeExcelProcess


@pytest.fixture
def connected_backend(patched_process):
    b = backend_module.ExcelComBackend()
    b.connect()
    yield b


def test_connect_sets_up_runtime_over_launched_application(connected_backend):
    assert connected_backend.is_alive is True
    assert connected_backend._runtime.core_workbook is not None
    assert connected_backend._runtime.agent_workbook is not None


def test_inject_module_requires_connect_first():
    b = backend_module.ExcelComBackend()
    with pytest.raises(BridgeDisconnectedError):
        b.inject_module("Main", "Sub X()\nEnd Sub\n")


def test_run_macro_requires_connect_first():
    b = backend_module.ExcelComBackend()
    with pytest.raises(BridgeDisconnectedError):
        b.run_macro("Main", "X", (), timeout=5.0)


def test_reset_requires_connect_first():
    b = backend_module.ExcelComBackend()
    with pytest.raises(BridgeDisconnectedError):
        b.reset()


def test_inject_module_delegates_to_runtime(connected_backend):
    connected_backend.inject_module("Main", "Sub X()\nEnd Sub\n", is_class=False)
    components = connected_backend._runtime.agent_workbook.VBProject.VBComponents.items
    assert components[0].Name == "Main"


def test_shutdown_closes_runtime_and_terminates_process(connected_backend, patched_process):
    process = connected_backend._process
    connected_backend.shutdown()
    assert process.terminated is True
    assert connected_backend._runtime is None
    assert connected_backend._process is None
    assert connected_backend.is_alive is False


def test_is_alive_false_before_connect():
    b = backend_module.ExcelComBackend()
    assert b.is_alive is False


def test_is_alive_false_when_process_reports_not_running(connected_backend):
    connected_backend._process._running = False
    assert connected_backend.is_alive is False


def test_is_alive_false_when_runtime_workbook_closed(connected_backend):
    connected_backend._runtime.agent_workbook.Close(SaveChanges=False)
    assert connected_backend.is_alive is False


class _FakeStream:
    pass


def _install_fake_win32com(monkeypatch, application):
    """run_macro() marshals the Application COM pointer across a thread
    boundary via pythoncom.CoMarshalInterThreadInterfaceInStream /
    CoGetInterfaceAndReleaseStream, then re-wraps it with
    win32com.client.Dispatch(). Faked here as an identity round-trip (hand
    back the same FakeApplication) since there's no real COM apartment to
    cross in-process -- this exercises run_macro()'s call sequence and
    timeout/error handling, not pywin32's actual marshaling machinery.
    """
    fake_pythoncom = types.SimpleNamespace(
        IID_IDispatch=object(),
        CoMarshalInterThreadInterfaceInStream=lambda iid, oleobj: _FakeStream(),
        CoGetInterfaceAndReleaseStream=lambda stream, iid: application,
        CoInitialize=lambda: None,
        CoUninitialize=lambda: None,
    )
    fake_win32com_client = types.SimpleNamespace(Dispatch=lambda obj: obj)
    fake_win32com = types.SimpleNamespace(client=fake_win32com_client)

    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setitem(sys.modules, "win32com", fake_win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_win32com_client)
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)


def test_run_macro_happy_path_unpacks_result(connected_backend, monkeypatch):
    application = connected_backend._runtime.application
    # oleobj_ is what real pywin32 Dispatch objects expose for marshaling;
    # our FakeApplication doesn't have one, so give it a placeholder.
    application._oleobj_ = object()
    application.run_results = [
        None,
        (True, 0, "", "", 42.0, 1, ["hi"], "tok-abc"),
    ]
    _install_fake_win32com(monkeypatch, application)

    result = connected_backend.run_macro("Main", "F", (), timeout=5.0, run_token="tok-abc")

    assert result.success is True
    assert result.return_value == 42.0
    assert result.output == ["hi"]


def test_run_macro_stale_token_raises(connected_backend, monkeypatch):
    application = connected_backend._runtime.application
    application._oleobj_ = object()
    application.run_results = [None, (True, 0, "", "", 1.0, 0, [], "old")]
    _install_fake_win32com(monkeypatch, application)

    with pytest.raises(basic_runtime.StaleRunError):
        connected_backend.run_macro("Main", "F", (), timeout=5.0, run_token="new")


def test_run_macro_timeout_terminates_process_and_disconnects(connected_backend, monkeypatch):
    application = connected_backend._runtime.application
    application._oleobj_ = object()

    def blocking_get_interface(stream, iid):
        import time

        time.sleep(10)
        return application

    fake_pythoncom = types.SimpleNamespace(
        IID_IDispatch=object(),
        CoMarshalInterThreadInterfaceInStream=lambda iid, oleobj: _FakeStream(),
        CoGetInterfaceAndReleaseStream=blocking_get_interface,
        CoInitialize=lambda: None,
        CoUninitialize=lambda: None,
    )
    fake_win32com_client = types.SimpleNamespace(Dispatch=lambda obj: obj)
    fake_win32com = types.SimpleNamespace(client=fake_win32com_client)
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setitem(sys.modules, "win32com", fake_win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_win32com_client)
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)

    process = connected_backend._process
    with pytest.raises(RunTimeoutError):
        connected_backend.run_macro("Main", "F", (), timeout=0.1, run_token="tok")

    assert process.terminated is True
    assert connected_backend._runtime is None
    assert connected_backend._process is None


def test_vbasession_end_to_end_against_fake_excel(patched_process, monkeypatch):
    """Runs a real VBASession -> wrapper.wrap_module() -> ExcelComBackend
    chain against the fake COM layer: confirms wrapper.py's generated
    `PyBridge.Core.PyBridge_*` source is exactly what basic_runtime.py's
    Core module (and this backend's Application.Run call strings) expect --
    not just each piece individually, which the other tests above already
    cover in isolation.
    """
    from vba_bridge.session import VBASession

    backend = backend_module.ExcelComBackend()
    session = VBASession(backend=backend)

    application = backend._runtime.application
    application._oleobj_ = object()
    # __PyBridgeRun's own Application.Run call has no useful return value;
    # the second call (PyBridge_GetResultPacked) is what session.run() reads.
    application.run_results = [None, (True, 0, "", "", 42.0, 0, [], "placeholder")]

    # The real run_token is generated fresh inside wrap_module() and isn't
    # known ahead of time, so intercept it off the injected Main module's
    # source rather than hardcoding a value the fake would have to guess.
    original_inject = backend.inject_module

    def capturing_inject(module_name, source, *, is_class=False):
        original_inject(module_name, source, is_class=is_class)
        if module_name == "Main":
            import re

            token = re.search(r'PyBridge_Reset\("([\w]+)"\)', source).group(1)
            application.run_results[-1] = (True, 0, "", "", 42.0, 0, [], token)

    monkeypatch.setattr(backend, "inject_module", capturing_inject)

    _install_fake_win32com(monkeypatch, application)
    result = session.run("Function F() As Long\n    F = 42\nEnd Function\n")

    assert result.success is True
    assert result.return_value == 42.0
    assert application.run_calls == [
        f"'{backend._runtime.agent_workbook.Name}'!Main.__PyBridgeRun",
        f"'{backend._runtime.core_workbook.Name}'!Core.PyBridge_GetResultPacked",
    ]
