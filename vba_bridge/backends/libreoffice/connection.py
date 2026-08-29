"""uno import handling and UNO socket connection."""

import glob
import sys
import time

from vba_bridge.exceptions import LaunchTimeoutError, UnoNotAvailableError

_CANDIDATE_DIST_PACKAGES = [
    "/usr/lib/python3/dist-packages",
    "/usr/lib/libreoffice/program",
    "/opt/libreoffice*/program",
]


def import_uno():
    """Import the `uno` module, trying common install locations if a plain
    `import uno` fails under the current interpreter. Raises UnoNotAvailableError
    with an actionable message rather than a bare ImportError, since the real
    failure mode users hit is a Python-ABI mismatch with pyuno.so, which is
    otherwise a confusing error.
    """
    try:
        import uno  # noqa: F401

        return uno
    except ImportError:
        pass

    tried = []
    for pattern in _CANDIDATE_DIST_PACKAGES:
        for path in glob.glob(pattern):
            tried.append(path)
            if path not in sys.path:
                sys.path.insert(0, path)
            try:
                import uno  # noqa: F401

                return uno
            except ImportError:
                continue

    raise UnoNotAvailableError(
        "Could not import the `uno` Python module. This usually means either "
        "LibreOffice isn't installed, or this Python interpreter's ABI doesn't "
        "match the one LibreOffice's pyuno.so was built against (a venv with a "
        "different Python minor version than the system Python, for example). "
        f"Tried adding these paths to sys.path: {tried or '(none found)'}. "
        "See README.md for supported setups."
    )


def resolve_context(port: int, *, timeout: float = 30.0, poll_interval: float = 0.25):
    """Connect to a running soffice's UNO socket and return its component context."""
    uno = import_uno()
    from com.sun.star.connection import NoConnectException

    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    conn_str = f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext"

    deadline = time.monotonic() + timeout
    last_error = None
    while time.monotonic() < deadline:
        try:
            return resolver.resolve(conn_str)
        except NoConnectException as exc:
            last_error = exc
            time.sleep(poll_interval)

    raise LaunchTimeoutError(
        f"Could not connect to soffice on port {port} within {timeout}s: {last_error}"
    )
