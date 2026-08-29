from dataclasses import dataclass
from typing import Any, Optional


@dataclass(frozen=True)
class VBAError:
    """Structured VBA/Basic runtime error, as caught via On Error GoTo / Err."""

    number: int
    description: str
    source: str = ""
    # Precise line tracking (Erl()-based) is not implemented in v1 -- see README
    # "Known limitations". Kept here for forward-compat with a future Excel/COM
    # backend, where it may be more tractable.
    line: Optional[int] = None
    line_is_approximate: bool = False


@dataclass(frozen=True)
class VBAResult:
    """Outcome of a single VBASession.run() call."""

    success: bool
    output: list
    return_value: Any = None
    error: Optional[VBAError] = None
    # Transport-level failure (timeout, dead process) distinct from a VBAError,
    # which represents an ordinary caught VBA runtime error.
    raw_exception: Optional[Exception] = None
    duration_ms: float = 0.0
    entry_point: str = ""
