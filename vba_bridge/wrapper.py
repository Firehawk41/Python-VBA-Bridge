"""Backend-agnostic source wrapping: turn a user's raw VBA/Basic snippet into a
module that (a) exposes an unqualified PyPrint() the user's code can call,
(b) invokes the detected entry point inside On Error GoTo, and (c) reports
success/failure through the persistent PyBridge.Core library rather than
raising.

Note: automatic Erl()-based line numbering was attempted and dropped -- see
the plan doc / README "Known limitations" for why. Error number/description/
source are the reliable signal in v1.
"""

import re
from typing import Any, Sequence, Tuple

_ENTRY_POINT_RE = re.compile(
    r"^\s*(?:Public\s+|Private\s+)?(Sub|Function)\s+(\w+)\s*\(",
    re.IGNORECASE | re.MULTILINE,
)

PYPRINT_FORWARDER = (
    "Sub PyPrint(ByVal Msg As Variant)\n"
    "    PyBridge.Core.PyBridge_Print(CStr(Msg))\n"
    "End Sub\n"
)


class EntryPointNotFoundError(ValueError):
    pass


def detect_entry_point(source: str) -> Tuple[str, bool]:
    """Return (name, is_function) for the first top-level Sub/Function in source."""
    match = _ENTRY_POINT_RE.search(source)
    if not match:
        raise EntryPointNotFoundError(
            "No 'Sub Name(...)' or 'Function Name(...)' found in the supplied source."
        )
    kind, name = match.group(1), match.group(2)
    return name, kind.lower() == "function"


def _basic_string_literal(value: str) -> str:
    escaped = value.replace('"', '""')
    if "\n" not in escaped:
        return f'"{escaped}"'
    parts = escaped.split("\n")
    return " & Chr(10) & ".join(f'"{p}"' for p in parts)


def _basic_scalar_literal(value: Any) -> str:
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        text = repr(value)
        return text if ("." in text or "e" in text or "E" in text) else f"{text}.0"
    if isinstance(value, str):
        return _basic_string_literal(value)
    if value is None:
        return "Empty"
    raise TypeError(f"Unsupported scalar arg type for VBA literal: {type(value)!r}")


def _basic_array_type_for(values: Sequence[Any]) -> str:
    if not values:
        return "Variant"
    if all(isinstance(v, bool) for v in values):
        return "Boolean"
    if all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in values):
        return "Double"
    if all(isinstance(v, str) for v in values):
        return "String"
    return "Variant"


def serialize_args(args: Sequence[Any]) -> Tuple[str, list]:
    """Build (preamble_basic_code, list_of_call_expressions) for the given
    Python args. Lists become typed local arrays built via explicit per-index
    assignment (NOT Basic's Array(), which produces an untyped Variant that
    breaks on typed array parameters -- see README known limitations).
    """
    preamble_lines = []
    call_exprs = []
    for i, value in enumerate(args):
        if isinstance(value, (list, tuple)):
            var_name = f"__pbArg{i}"
            elem_type = _basic_array_type_for(value)
            if value:
                preamble_lines.append(f"Dim {var_name}({len(value) - 1}) As {elem_type}")
                for idx, elem in enumerate(value):
                    literal = _basic_scalar_literal(elem)
                    preamble_lines.append(f"{var_name}({idx}) = {literal}")
            else:
                preamble_lines.append(f"Dim {var_name}(-1) As {elem_type}")
            call_exprs.append(var_name)
        else:
            call_exprs.append(_basic_scalar_literal(value))
    return "\n".join(preamble_lines), call_exprs


def wrap_module(
    user_source: str,
    *,
    entry_point: str = None,
    is_function: bool = None,
    args: Sequence[Any] = (),
) -> Tuple[str, str, bool]:
    """Build the full Main module source for one run() call.

    Returns (wrapped_source, entry_point_name, is_function). If entry_point is
    not given, it is auto-detected from user_source.
    """
    if entry_point is None:
        entry_point, is_function = detect_entry_point(user_source)
    elif is_function is None:
        # explicit entry_point but caller didn't say Sub/Function -- detect it
        _, is_function = detect_entry_point(user_source)

    preamble, call_exprs = serialize_args(args)
    args_joined = ", ".join(call_exprs)
    call_expr = f"{entry_point}({args_joined})"

    if preamble:
        preamble = "    " + preamble.replace("\n", "\n    ") + "\n"

    if is_function:
        call_line = f"{preamble}    PyBridge.Core.PyBridge_SetSuccess({call_expr})"
    else:
        call_line = f"{preamble}    Call {call_expr}\n    PyBridge.Core.PyBridge_SetSuccess(Empty)"

    wrapped = (
        "Option VBASupport 1\n"
        "Option Explicit\n"
        "\n"
        f"{PYPRINT_FORWARDER}\n"
        f"{user_source}\n"
        "\n"
        "Sub __PyBridgeRun()\n"
        "    PyBridge.Core.PyBridge_Reset()\n"
        "    On Error GoTo ErrHandler\n"
        f"{call_line}\n"
        "    Exit Sub\n"
        "ErrHandler:\n"
        "    PyBridge.Core.PyBridge_SetError(Err.Number, Err.Description, Err.Source)\n"
        "End Sub\n"
    )
    return wrapped, entry_point, is_function
