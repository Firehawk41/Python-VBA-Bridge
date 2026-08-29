from vba_bridge.backends.base import Backend, RawRunResult
from vba_bridge.session import VBASession


class FakeBackend(Backend):
    """In-memory stand-in for a real backend, to test VBASession's control
    flow (connect timing, reset/restart/close, error surfacing) without
    starting a real soffice process."""

    def __init__(self, run_result=None, run_exception=None):
        self.connect_count = 0
        self.reset_count = 0
        self.shutdown_count = 0
        self.injected = []
        self.run_calls = []
        self._alive = False
        self._run_result = run_result or RawRunResult(success=True, output=[], return_value=None)
        self._run_exception = run_exception

    def connect(self):
        self.connect_count += 1
        self._alive = True

    def inject_module(self, module_name, source):
        self.injected.append((module_name, source))

    def run_macro(self, module_name, entry_point, args, *, timeout):
        self.run_calls.append((module_name, entry_point, tuple(args), timeout))
        if self._run_exception is not None:
            raise self._run_exception
        return self._run_result

    def reset(self):
        self.reset_count += 1

    def shutdown(self):
        self.shutdown_count += 1
        self._alive = False

    @property
    def is_alive(self):
        return self._alive


def test_auto_start_connects_backend():
    backend = FakeBackend()
    VBASession(backend=backend)
    assert backend.connect_count == 1


def test_auto_start_false_defers_connect_until_run():
    backend = FakeBackend()
    session = VBASession(backend=backend, auto_start=False)
    assert backend.connect_count == 0
    session.run("Sub X()\nEnd Sub\n")
    assert backend.connect_count == 1


def test_run_success_returns_result_with_output_and_return_value():
    backend = FakeBackend(
        run_result=RawRunResult(success=True, output=["hi"], return_value=7.0)
    )
    session = VBASession(backend=backend)
    result = session.run("Function F() As Double\nEnd Function\n")
    assert result.success is True
    assert result.output == ["hi"]
    assert result.return_value == 7.0
    assert result.error is None
    assert result.entry_point == "F"


def test_run_failure_builds_structured_error():
    backend = FakeBackend(
        run_result=RawRunResult(
            success=False, output=[], err_number=11, err_description="Division by zero.", err_source="Main"
        )
    )
    session = VBASession(backend=backend)
    result = session.run("Function F() As Double\nEnd Function\n")
    assert result.success is False
    assert result.error.number == 11
    assert result.error.description == "Division by zero."
    assert result.raw_exception is None


def test_run_transport_exception_surfaces_as_raw_exception_not_crash():
    backend = FakeBackend(run_exception=RuntimeError("bridge died"))
    session = VBASession(backend=backend)
    result = session.run("Sub X()\nEnd Sub\n")
    assert result.success is False
    assert result.error is None
    assert isinstance(result.raw_exception, RuntimeError)


def test_run_missing_entry_point_does_not_raise():
    backend = FakeBackend()
    session = VBASession(backend=backend)
    result = session.run("Dim x As Integer\n")
    assert result.success is False
    assert result.raw_exception is not None
    assert backend.run_calls == []  # never reached the backend


def test_run_reconnects_when_backend_reports_dead():
    backend = FakeBackend()
    session = VBASession(backend=backend)
    assert backend.connect_count == 1
    backend._alive = False
    session.run("Sub X()\nEnd Sub\n")
    assert backend.connect_count == 2


def test_reset_delegates_to_backend():
    backend = FakeBackend()
    session = VBASession(backend=backend)
    session.reset()
    assert backend.reset_count == 1


def test_close_delegates_to_backend_and_is_idempotent():
    backend = FakeBackend()
    session = VBASession(backend=backend)
    session.close()
    session.close()
    assert backend.shutdown_count == 1


def test_context_manager_closes_on_exit():
    backend = FakeBackend()
    with VBASession(backend=backend) as session:
        session.run("Sub X()\nEnd Sub\n")
    assert backend.shutdown_count == 1


def test_run_passes_timeout_through_to_backend():
    backend = FakeBackend()
    session = VBASession(backend=backend, run_timeout=30.0)
    session.run("Sub X()\nEnd Sub\n", timeout=5.0)
    assert backend.run_calls[0][3] == 5.0


def test_run_uses_default_timeout_when_not_specified():
    backend = FakeBackend()
    session = VBASession(backend=backend, run_timeout=42.0)
    session.run("Sub X()\nEnd Sub\n")
    assert backend.run_calls[0][3] == 42.0
