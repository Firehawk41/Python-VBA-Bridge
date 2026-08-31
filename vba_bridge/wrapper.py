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
import uuid
from typing import Any, Dict, Mapping, Sequence, Tuple

_ENTRY_POINT_RE = re.compile(
    r"^\s*(?:Public\s+|Private\s+)?(Sub|Function)\s+(\w+)\s*\(",
    re.IGNORECASE | re.MULTILINE,
)

# A VBA class export (.cls file, e.g. from the VBA editor's "Export File...")
# starts with a header block that's only valid inside the .cls FILE FORMAT
# itself (parsed by VBA's own project importer) -- not valid Basic source
# when injected directly as a module string, and silently breaks compilation.
_CLASS_EXPORT_HEADER_RE = re.compile(
    r"\A[ \t]*VERSION\s+[\d.]+\s+CLASS[ \t]*\r?\n"
    r"[ \t]*BEGIN[ \t]*\r?\n"
    r"(?:.*\r?\n)*?"
    r"[ \t]*END[ \t]*\r?\n",
    re.IGNORECASE,
)

# vba_bridge always adds its own "Option VBASupport 1"/"Option Explicit" at
# the true top of every module it injects; a pasted/exported module that
# already has its own copy (anywhere but the first line) would otherwise
# produce a duplicate-Option compile error, so any pre-existing occurrence is
# stripped first.
_OPTION_PRAGMA_RE = re.compile(
    r"^[ \t]*Option\s+(?:VBASupport\s+1|Explicit)[ \t]*\r?\n?",
    re.IGNORECASE | re.MULTILINE,
)

# A real VBA export (.bas/.cls from the VBA editor's "Export File...") always
# starts with "Attribute VB_Name = "..."" (and a .cls export follows it with
# VB_GlobalNameSpace/VB_Creatable/VB_PredeclaredId/VB_Exposed) -- statements
# only VBE's own project-file importer understands, never valid as literal
# source text handed to CodeModule.AddFromString(). Confirmed against real
# Excel: leaving even just the first one in place is a genuine "Compile
# error: Syntax error" that breaks the whole module (LibreOffice Basic
# tolerates them, which is why this was previously assumed harmless).
_ATTRIBUTE_STATEMENT_RE = re.compile(
    r"^[ \t]*Attribute\s+[\w.]+\s*=.*\r?\n?", re.IGNORECASE | re.MULTILINE
)

PYPRINT_FORWARDER = (
    "Sub PyPrint(ByVal Msg As Variant)\n"
    "    PyBridge.Core.PyBridge_Print(CStr(Msg))\n"
    "End Sub\n"
)

# Multi-module programs (wrap_program) inject an always-regenerated
# orchestrator module under this name; a user module of the same name would
# collide with it and is rejected.
ORCHESTRATOR_MODULE_NAME = "PyBridgeMain"


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
            var_name = f"pbArg{i}"
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


def normalize_user_module_source(source: str) -> str:
    """Make arbitrary pasted/exported VBA module text safe to inject directly
    as a Basic module's full source: strip a class-export's VERSION/BEGIN/END
    header block (see _CLASS_EXPORT_HEADER_RE), any pre-existing "Option
    VBASupport 1"/"Option Explicit" line (vba_bridge adds its own at the true
    top of the module -- a duplicate anywhere else is a compile error), and
    any "Attribute VB_..." statements (VBE-import-only metadata; a genuine
    compile error as literal source under real Excel VBA -- see
    _ATTRIBUTE_STATEMENT_RE).
    """
    source = _CLASS_EXPORT_HEADER_RE.sub("", source, count=1)
    source = _ATTRIBUTE_STATEMENT_RE.sub("", source)
    source = _OPTION_PRAGMA_RE.sub("", source)
    return source


def _new_run_token() -> str:
    return uuid.uuid4().hex


def wrap_module(
    user_source: str,
    *,
    entry_point: str = None,
    is_function: bool = None,
    args: Sequence[Any] = (),
) -> Tuple[str, str, bool, str]:
    """Build the full Main module source for one run() call.

    Returns (wrapped_source, entry_point_name, is_function, run_token). If
    entry_point is not given, it is auto-detected from user_source. run_token
    is a fresh per-call value written by the generated PyBridge_Reset() call
    and must be checked against the read-back result -- see StaleRunError.
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

    run_token = _new_run_token()
    cleaned_user_source = normalize_user_module_source(user_source)
    wrapped = (
        "Option VBASupport 1\n"
        "Option Explicit\n"
        "\n"
        f"{PYPRINT_FORWARDER}\n"
        f"{cleaned_user_source}\n"
        "\n"
        "Sub PyBridgeRun()\n"
        f'    PyBridge.Core.PyBridge_Reset("{run_token}")\n'
        "    On Error GoTo ErrHandler\n"
        f"{call_line}\n"
        "    Exit Sub\n"
        "ErrHandler:\n"
        "    Call PyBridge.Core.PyBridge_SetError(Err.Number, Err.Description, Err.Source)\n"
        "End Sub\n"
    )
    return wrapped, entry_point, is_function, run_token


def _find_proc_in_source(source: str, proc_name: str):
    """Return is_function for `proc_name` if declared as a top-level Sub/Function
    in source, else None."""
    pattern = re.compile(
        rf"^\s*(?:Public\s+|Private\s+)?(Sub|Function)\s+{re.escape(proc_name)}\s*\(",
        re.IGNORECASE | re.MULTILINE,
    )
    match = pattern.search(source)
    if not match:
        return None
    return match.group(1).lower() == "function"


def resolve_program_entry_point(
    modules: Mapping[str, str],
    *,
    class_modules: Sequence[str] = (),
    entry_point: str = None,
) -> Tuple[str, str, bool]:
    """Return (module_name, proc_name, is_function) for a multi-module program.

    entry_point may be "ModuleName.ProcName" (qualified) or a bare "ProcName"
    (searched across all non-class modules, in dict order). If omitted,
    auto-detects the first Sub/Function found across non-class modules, in
    dict order -- same semantics as detect_entry_point(), extended across
    modules. Class modules are skipped: an entry point must be a standalone
    Sub/Function, not an instance method (there's nothing to instantiate the
    class with) -- write a small driver Sub elsewhere that does `Dim c As New
    ClassName` and calls into it instead.
    """
    class_modules = set(class_modules)

    if entry_point:
        if "." in entry_point:
            module_name, proc_name = entry_point.split(".", 1)
            if module_name not in modules:
                raise EntryPointNotFoundError(
                    f"Entry point module '{module_name}' not found in supplied modules."
                )
            if module_name in class_modules:
                raise EntryPointNotFoundError(
                    f"'{module_name}' is a class module; an entry point must be a "
                    "standalone Sub/Function in a non-class module. Write a driver "
                    "Sub elsewhere that instantiates the class and calls into it."
                )
            is_function = _find_proc_in_source(modules[module_name], proc_name)
            if is_function is None:
                raise EntryPointNotFoundError(
                    f"'{proc_name}' not found as a Sub/Function in module '{module_name}'."
                )
            return module_name, proc_name, is_function

        proc_name = entry_point
        for module_name, source in modules.items():
            if module_name in class_modules:
                continue
            is_function = _find_proc_in_source(source, proc_name)
            if is_function is not None:
                return module_name, proc_name, is_function
        raise EntryPointNotFoundError(
            f"'{proc_name}' not found as a Sub/Function in any supplied non-class module."
        )

    for module_name, source in modules.items():
        if module_name in class_modules:
            continue
        try:
            proc_name, is_function = detect_entry_point(source)
            return module_name, proc_name, is_function
        except EntryPointNotFoundError:
            continue
    raise EntryPointNotFoundError(
        "No 'Sub Name(...)' or 'Function Name(...)' found in any supplied non-class module."
    )


def wrap_program(
    modules: Mapping[str, str],
    *,
    class_modules: Sequence[str] = (),
    entry_point: str = None,
    args: Sequence[Any] = (),
) -> Tuple[Dict[str, str], str, str, bool, str]:
    """Build the module set to inject for a multi-module program.

    Returns (modules_to_inject, entry_module, entry_proc, is_function,
    run_token). modules_to_inject is `modules`, each normalized via
    normalize_user_module_source() and given its own "Option VBASupport
    1"/"Option Explicit" (so a real VBA export -- .bas/.cls text pasted
    as-is, headers included -- can be handed in directly), plus a generated
    orchestrator module (ORCHESTRATOR_MODULE_NAME) holding the PyPrint
    forwarder and PyBridgeRun. No PyPrint/error-handling text is spliced
    into user modules themselves (unlike wrap_module()'s single-string path,
    which combines everything into one module). run_token is a fresh
    per-call value written by the generated PyBridge_Reset() call and must
    be checked against the read-back result -- see StaleRunError.

    The call inside PyBridgeRun is always unqualified (bare proc name),
    even when entry_point was given qualified -- cross-module unqualified
    calls within the same library are the confirmed-working mechanism;
    qualifying "Module.Proc" was only used above to locate the right module's
    source for is_function detection.
    """
    if ORCHESTRATOR_MODULE_NAME in modules:
        raise ValueError(
            f"'{ORCHESTRATOR_MODULE_NAME}' is reserved for vba_bridge's generated "
            "orchestrator module; rename your module."
        )

    entry_module, entry_proc, is_function = resolve_program_entry_point(
        modules, class_modules=class_modules, entry_point=entry_point
    )

    preamble, call_exprs = serialize_args(args)
    args_joined = ", ".join(call_exprs)
    call_expr = f"{entry_proc}({args_joined})"

    if preamble:
        preamble = "    " + preamble.replace("\n", "\n    ") + "\n"

    if is_function:
        call_line = f"{preamble}    PyBridge.Core.PyBridge_SetSuccess({call_expr})"
    else:
        call_line = f"{preamble}    Call {call_expr}\n    PyBridge.Core.PyBridge_SetSuccess(Empty)"

    run_token = _new_run_token()
    orchestrator_source = (
        "Option VBASupport 1\n"
        "Option Explicit\n"
        "\n"
        f"{PYPRINT_FORWARDER}\n"
        "Sub PyBridgeRun()\n"
        f'    PyBridge.Core.PyBridge_Reset("{run_token}")\n'
        "    On Error GoTo ErrHandler\n"
        f"{call_line}\n"
        "    Exit Sub\n"
        "ErrHandler:\n"
        "    Call PyBridge.Core.PyBridge_SetError(Err.Number, Err.Description, Err.Source)\n"
        "End Sub\n"
    )

    modules_to_inject = {
        name: "Option VBASupport 1\nOption Explicit\n\n" + normalize_user_module_source(source)
        for name, source in modules.items()
    }
    modules_to_inject[ORCHESTRATOR_MODULE_NAME] = orchestrator_source
    return modules_to_inject, entry_module, entry_proc, is_function, run_token
