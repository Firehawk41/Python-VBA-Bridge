import shutil

import pytest

from vba_bridge import VBASession

_SOFFICE_AVAILABLE = shutil.which("soffice") is not None


def pytest_collection_modifyitems(config, items):
    if _SOFFICE_AVAILABLE:
        return
    skip_marker = pytest.mark.skip(reason="soffice not found on PATH")
    for item in items:
        if "integration" in str(item.fspath):
            item.add_marker(skip_marker)


@pytest.fixture(scope="module")
def vba_session():
    """One real soffice process shared across a test module, to avoid paying
    the multi-second startup cost per test."""
    session = VBASession(run_timeout=20.0)
    yield session
    session.close()
