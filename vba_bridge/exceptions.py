class BridgeError(Exception):
    """Base class for all vba_bridge errors."""


class UnoNotAvailableError(BridgeError):
    """The `uno`/`pyuno` Python bindings could not be imported.

    This usually means either LibreOffice isn't installed, or the Python
    interpreter running this code has a different ABI than the one LibreOffice's
    pyuno.so was built against. See README.md for supported setups.
    """


class LaunchTimeoutError(BridgeError):
    """soffice did not come up and accept a UNO connection within the timeout."""


class BridgeDisconnectedError(BridgeError):
    """The backend process died or the UNO connection was lost."""


class RunTimeoutError(BridgeError):
    """A run_macro() call exceeded its timeout and the backend was restarted."""


class StaleRunError(BridgeError):
    """A run_macro() call returned without actually executing the requested code.

    Detected via a per-call token written at the start of every run and
    checked on read-back: if the token doesn't match, PyBridgeRun never
    ran, almost always because a module injected for this run failed to
    compile (invalid Basic syntax -- e.g. a VBA class-export's VERSION/BEGIN
    header block injected verbatim, or a genuine typo). Without this check
    the caller would silently get back whatever result the *previous*
    successful run left behind, misread as this run's real result.
    """


class ExcelComNotAvailableError(BridgeError):
    """The `pywin32` COM bindings could not be imported, or this isn't Windows.

    ExcelComBackend automates a real Excel.Application over COM and only
    works on Windows with pywin32 and Excel installed.
    """


class VbomAccessDeniedError(BridgeError):
    """Excel refused programmatic access to a VBA project's object model.

    ExcelComBackend injects code by manipulating VBComponents directly, which
    Excel blocks unless "Trust access to the VBA project object model" is
    enabled: File > Options > Trust Center > Trust Center Settings > Macro
    Settings. This is a one-time, per-machine setting -- vba_bridge cannot
    enable it programmatically (doing so silently would itself be a security
    hole), so it must be turned on by hand before ExcelComBackend can run.
    """
