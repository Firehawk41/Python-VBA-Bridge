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
