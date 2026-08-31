"""Owns two hidden(-ish) workbooks and the VBA-project reference between them:
a persistent "PyBridge" project holding the Core module (mirrors the
LibreOffice backend's PyBridge.Core Basic library), and an "Agent" project
that references it and holds the modules injected/replaced on every run().

Verified against real Excel (Windows, Excel 16.0) -- see README "v2: real
Excel via pywin32" for the confirmed real-VBA-vs-LibreOffice-Basic
differences this uncovered (identifiers can't start with a leading
underscore, no `Option VBASupport 1`, fixed-size arrays can't be `ReDim`'d,
multi-arg Sub calls need `Call`, etc.).

Why two workbooks instead of one: wrapper.py's generated source calls
`PyBridge.Core.PyBridge_Reset(...)` etc. -- a Project.Module.Proc qualified
call -- unchanged from the LibreOffice backend (BasicLibraries there are
addressed the same way). Real VBA has no in-project "library" concept, but
it does support exactly this qualification across a VBA-project reference
(the same mechanism an .xlam add-in uses to expose Public Subs to a
workbook that references it), so two projects reproduces the LibreOffice
addressing scheme without touching wrapper.py at all.
"""

import os
import re
import tempfile

from vba_bridge.backends.base import RawRunResult
from vba_bridge.exceptions import StaleRunError

# wrapper.py (shared with LibreOfficeBackend) unconditionally prefixes every
# module it builds with "Option VBASupport 1" -- a LibreOffice Basic pragma
# that enables its VBA-compatibility mode. Real Excel VBA has no such
# statement at all: leaving it in place is a compile error that silently
# breaks the *entire* module (every Application.Run into it then fails with
# a generic "macro may not be available" COM error, confirmed against real
# Excel). "Option Explicit" is real, valid VBA and is kept.
_VBASUPPORT_PRAGMA_RE = re.compile(
    r"^[ \t]*Option\s+VBASupport\s+1[ \t]*\r?\n?", re.IGNORECASE | re.MULTILINE
)


def _strip_vbasupport_pragma(source: str) -> str:
    return _VBASUPPORT_PRAGMA_RE.sub("", source)

CORE_PROJECT_NAME = "PyBridge"
CORE_MODULE_NAME = "Core"
AGENT_WORKBOOK_BASENAME = "PyBridgeAgent"
CORE_WORKBOOK_BASENAME = "PyBridgeCore"

# vbext_ComponentType enum (VBA extensibility library) -- stable, documented
# constants; used as literals rather than through win32com.client.constants
# to avoid depending on makepy/gen_py cache generation being present.
_VBEXT_CT_STDMODULE = 1
_VBEXT_CT_CLASSMODULE = 2

# xlOpenXMLWorkbookMacroEnabled -- required so VBComponents/CodeModule survive
# a save; a plain .xlsx silently drops all VBA content.
_XL_MACRO_ENABLED_WORKBOOK = 52

CORE_MODULE_SOURCE = """Option Explicit

Public PyBridge_Output() As String
Public PyBridge_OutputCount As Long
Public PyBridge_Success As Boolean
Public PyBridge_ErrNumber As Long
Public PyBridge_ErrDescription As String
Public PyBridge_ErrSource As String
Public PyBridge_ReturnValue As Variant
Public PyBridge_RunToken As String

Sub PyBridge_Reset(ByVal Token As String)
    ReDim PyBridge_Output(63)
    PyBridge_OutputCount = 0
    PyBridge_Success = False
    PyBridge_ErrNumber = 0
    PyBridge_ErrDescription = ""
    PyBridge_ErrSource = ""
    PyBridge_ReturnValue = Empty
    PyBridge_RunToken = Token
End Sub

Sub PyBridge_Print(ByVal Msg As String)
    If PyBridge_OutputCount > UBound(PyBridge_Output) Then
        ReDim Preserve PyBridge_Output(UBound(PyBridge_Output) * 2 + 1)
    End If
    PyBridge_Output(PyBridge_OutputCount) = Msg
    PyBridge_OutputCount = PyBridge_OutputCount + 1
End Sub

Sub PyBridge_SetSuccess(ByVal RetVal As Variant)
    PyBridge_Success = True
    PyBridge_ReturnValue = RetVal
End Sub

Sub PyBridge_SetError(ByVal Num As Long, ByVal Desc As String, ByVal Src As String)
    PyBridge_Success = False
    PyBridge_ErrNumber = Num
    PyBridge_ErrDescription = Desc
    PyBridge_ErrSource = Src
End Sub

Function PyBridge_GetResultPacked() As Variant
    Dim outArr() As String
    Dim i As Long
    Dim upperBound As Long
    upperBound = PyBridge_OutputCount - 1
    If upperBound < 0 Then upperBound = 0
    ReDim outArr(upperBound)
    For i = 0 To PyBridge_OutputCount - 1
        outArr(i) = PyBridge_Output(i)
    Next i
    PyBridge_GetResultPacked = Array(PyBridge_Success, PyBridge_ErrNumber, _
        PyBridge_ErrDescription, PyBridge_ErrSource, PyBridge_ReturnValue, _
        PyBridge_OutputCount, outArr, PyBridge_RunToken)
End Function
"""


class VBAProjectRuntime:
    """Wraps the Core + Agent workbook pair over an already-launched
    Excel.Application."""

    def __init__(self, application):
        self.application = application
        self.core_workbook = None
        self.agent_workbook = None
        self._tmp_dir = None

    def setup(self) -> None:
        self._tmp_dir = tempfile.mkdtemp(prefix="vba_bridge_excel_")

        self.core_workbook = self._new_macro_workbook(
            os.path.join(self._tmp_dir, f"{CORE_WORKBOOK_BASENAME}.xlsm")
        )
        self.core_workbook.VBProject.Name = CORE_PROJECT_NAME
        self._add_module(
            self.core_workbook, CORE_MODULE_NAME, CORE_MODULE_SOURCE, is_class=False
        )
        # References.AddFromFile resolves a path on disk -- the Core
        # workbook's project has to actually be saved before the Agent
        # workbook can reference it.
        self.core_workbook.Save()

        self.agent_workbook = self._new_macro_workbook(
            os.path.join(self._tmp_dir, f"{AGENT_WORKBOOK_BASENAME}.xlsm")
        )
        self.agent_workbook.VBProject.References.AddFromFile(
            self.core_workbook.FullName
        )

    def _new_macro_workbook(self, path: str):
        wb = self.application.Workbooks.Add()
        wb.SaveAs(path, FileFormat=_XL_MACRO_ENABLED_WORKBOOK)
        return wb

    @staticmethod
    def _add_module(workbook, module_name: str, source: str, *, is_class: bool) -> None:
        components = workbook.VBProject.VBComponents
        existing = _find_component(components, module_name)
        if existing is not None:
            components.Remove(existing)

        component_type = _VBEXT_CT_CLASSMODULE if is_class else _VBEXT_CT_STDMODULE
        component = components.Add(component_type)
        component.Name = module_name
        code_module = component.CodeModule
        # A brand-new component isn't always a blank slate: when the VBE
        # option "Require Variable Declaration" is on (Tools > Options >
        # Editor -- off by default, but a per-machine/per-user setting we
        # can't rely on), Excel pre-seeds every new module with its own
        # "Option Explicit" line before we ever touch it. Our own source
        # below adds another one, and two "Option Explicit" statements in
        # one module is a compile error that silently breaks every macro in
        # it (confirmed against real Excel -- Application.Run then fails
        # with a generic "macro may not be available" COM error). Clear
        # whatever the component started with so injection is idempotent
        # regardless of that setting.
        if code_module.CountOfLines > 0:
            code_module.DeleteLines(1, code_module.CountOfLines)
        code_module.AddFromString(_strip_vbasupport_pragma(source))

    def inject_module(self, module_name: str, source: str, *, is_class: bool = False) -> None:
        self._add_module(self.agent_workbook, module_name, source, is_class=is_class)

    @property
    def agent_workbook_name(self) -> str:
        return self.agent_workbook.Name

    @property
    def core_workbook_name(self) -> str:
        return self.core_workbook.Name

    def run(self, module_name: str, expected_token: str = None) -> RawRunResult:
        run_on_application(
            self.application, self.agent_workbook_name, self.core_workbook_name, module_name
        )
        packed = get_result_packed(self.application, self.core_workbook_name)
        return unpack_result(module_name, packed, expected_token)

    def is_alive(self) -> bool:
        if self.agent_workbook is None:
            return False
        try:
            _ = self.agent_workbook.Name
            return True
        except Exception:
            return False

    def close(self) -> None:
        for wb in (self.agent_workbook, self.core_workbook):
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:
                pass
        self.agent_workbook = None
        self.core_workbook = None

        if self._tmp_dir is not None:
            import shutil

            shutil.rmtree(self._tmp_dir, ignore_errors=True)
            self._tmp_dir = None


def _find_component(components, name: str):
    for component in components:
        if component.Name == name:
            return component
    return None


# --- Module-level helpers, callable against a bare `Application` COM object -
# (i.e. without a full VBAProjectRuntime instance). Used directly by
# VBAProjectRuntime.run() above, and separately by backend.py's timeout path,
# which runs on a worker thread against a marshaled Application proxy rather
# than the VBAProjectRuntime that created it -- see backend.py's run_macro()
# for why the two paths need to share this logic without sharing an instance.


def run_on_application(
    application, agent_workbook_name: str, core_workbook_name: str, module_name: str
) -> None:
    application.Run(f"'{agent_workbook_name}'!{module_name}.PyBridgeRun")


def get_result_packed(application, core_workbook_name: str):
    return application.Run(
        f"'{core_workbook_name}'!{CORE_MODULE_NAME}.PyBridge_GetResultPacked"
    )


def unpack_result(module_name: str, packed, expected_token: str = None) -> RawRunResult:
    (
        success,
        err_number,
        err_description,
        err_source,
        return_value,
        output_count,
        output_raw,
        run_token,
    ) = packed
    if expected_token is not None and str(run_token) != expected_token:
        raise StaleRunError(
            f"'{module_name}.PyBridgeRun' did not actually execute -- the "
            "reported result would have been left over from a previous run. "
            "This almost always means a module injected for this call "
            "failed to compile (a genuine syntax error, or invalid Basic "
            "text like a VBA class-export's 'VERSION ... CLASS' / "
            "'BEGIN...END' header injected verbatim). Check the injected "
            "module sources for syntax problems."
        )
    return RawRunResult(
        success=bool(success),
        output=list(output_raw)[: int(output_count)],
        return_value=None if not success else return_value,
        err_number=int(err_number),
        err_description=str(err_description),
        err_source=str(err_source),
    )
