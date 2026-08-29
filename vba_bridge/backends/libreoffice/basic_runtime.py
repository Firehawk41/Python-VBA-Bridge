"""Owns the hidden Calc document, the persistent PyBridge.Core library, and the
AgentCode.Main module that gets replaced on every run(). See the plan doc's
"Corrections found via empirical testing" section for why each piece here is
shaped the way it is (macro security, explicit loadLibrary, no Erl() numbering).
"""

from vba_bridge.backends.base import RawRunResult

CORE_LIBRARY_NAME = "PyBridge"
CORE_MODULE_NAME = "Core"
AGENT_LIBRARY_NAME = "AgentCode"
AGENT_MODULE_NAME = "Main"

CORE_MODULE_SOURCE = """Option VBASupport 1
Option Explicit

Public PyBridge_Output(63) As String
Public PyBridge_OutputCount As Long
Public PyBridge_Success As Boolean
Public PyBridge_ErrNumber As Long
Public PyBridge_ErrDescription As String
Public PyBridge_ErrSource As String
Public PyBridge_ReturnValue As Variant

Sub PyBridge_Reset()
    ReDim PyBridge_Output(63)
    PyBridge_OutputCount = 0
    PyBridge_Success = False
    PyBridge_ErrNumber = 0
    PyBridge_ErrDescription = ""
    PyBridge_ErrSource = ""
    PyBridge_ReturnValue = Empty
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
    ' ReDim with a negative bound (when OutputCount is 0) silently breaks
    ' module compilation in LO Basic -- always ReDim to at least 0 and report
    ' the real count separately so Python can trim the array correctly.
    upperBound = PyBridge_OutputCount - 1
    If upperBound < 0 Then upperBound = 0
    ReDim outArr(upperBound)
    For i = 0 To PyBridge_OutputCount - 1
        outArr(i) = PyBridge_Output(i)
    Next i
    PyBridge_GetResultPacked = Array(PyBridge_Success, PyBridge_ErrNumber, _
        PyBridge_ErrDescription, PyBridge_ErrSource, PyBridge_ReturnValue, _
        PyBridge_OutputCount, outArr)
End Function
"""


class BasicRuntime:
    """Wraps one hidden Calc document + its Basic libraries over an already
    resolved UNO context."""

    def __init__(self, ctx):
        self.ctx = ctx
        self.smgr = ctx.ServiceManager
        self.desktop = None
        self.doc = None
        self._provider = None

    def setup(self) -> None:
        from com.sun.star.beans import PropertyValue

        self._lower_macro_security()

        self.desktop = self.smgr.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx
        )
        hidden = PropertyValue()
        hidden.Name = "Hidden"
        hidden.Value = True
        self.doc = self.desktop.loadComponentFromURL(
            "private:factory/scalc", "_blank", 0, (hidden,)
        )

        libs = self.doc.BasicLibraries
        libs.VBACompatibilityMode = True

        self._ensure_library(CORE_LIBRARY_NAME)
        core_lib = libs.getByName(CORE_LIBRARY_NAME)
        self._replace_or_insert(core_lib, CORE_MODULE_NAME, CORE_MODULE_SOURCE)

        self._ensure_library(AGENT_LIBRARY_NAME)

        factory = self.smgr.createInstanceWithContext(
            "com.sun.star.script.provider.MasterScriptProviderFactory", self.ctx
        )
        self._provider = factory.createScriptProvider(self.doc)

    def _lower_macro_security(self) -> None:
        from com.sun.star.beans import PropertyValue

        cp = self.smgr.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", self.ctx
        )
        node = PropertyValue()
        node.Name = "nodepath"
        node.Value = "/org.openoffice.Office.Common/Security/Scripting"
        cu = cp.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationUpdateAccess", (node,)
        )
        cu.setPropertyValue("MacroSecurityLevel", 0)
        cu.commitChanges()

    def _ensure_library(self, name: str) -> None:
        libs = self.doc.BasicLibraries
        if not libs.hasByName(name):
            libs.createLibrary(name)
        if not libs.isLibraryLoaded(name):
            libs.loadLibrary(name)

    @staticmethod
    def _replace_or_insert(lib, module_name: str, source: str) -> None:
        if lib.hasByName(module_name):
            lib.replaceByName(module_name, source)
        else:
            lib.insertByName(module_name, source)

    @staticmethod
    def _mark_as_class(lib, module_name: str) -> None:
        """Mark module_name as a VBA class module (New module_name support).
        Must be called before the FIRST insertByName() for a given module --
        marking it after insert silently fails (confirmed empirically: the
        metadata is stored correctly per getModuleInfo(), but `New X` still
        doesn't compile). Re-affirming it (remove+re-insert) after every
        replaceByName() is also required, or a later source update loses the
        class typing the same way.
        """
        import uno
        from com.sun.star.script import ModuleType

        info = uno.createUnoStruct("com.sun.star.script.ModuleInfo")
        info.ModuleType = ModuleType.CLASS
        if lib.hasModuleInfo(module_name):
            lib.removeModuleInfo(module_name)
        lib.insertModuleInfo(module_name, info)

    def inject_module(self, module_name: str, source: str, *, is_class: bool = False) -> None:
        libs = self.doc.BasicLibraries
        lib = libs.getByName(AGENT_LIBRARY_NAME)
        is_new = not lib.hasByName(module_name)
        if is_class and is_new:
            self._mark_as_class(lib, module_name)
        self._replace_or_insert(lib, module_name, source)
        if is_class:
            self._mark_as_class(lib, module_name)

    def _invoke(self, library: str, module: str, macro: str, args: tuple = ()):
        uri = f"vnd.sun.star.script:{library}.{module}.{macro}?language=Basic&location=document"
        script = self._provider.getScript(uri)
        return script.invoke(args, (), ())[0]

    def run(self, module_name: str) -> RawRunResult:
        self._invoke(AGENT_LIBRARY_NAME, module_name, "__PyBridgeRun")
        packed = self._invoke(CORE_LIBRARY_NAME, CORE_MODULE_NAME, "PyBridge_GetResultPacked")
        success, err_number, err_description, err_source, return_value, output_count, output_raw = packed
        return RawRunResult(
            success=bool(success),
            output=list(output_raw)[: int(output_count)],
            return_value=None if not success else return_value,
            err_number=int(err_number),
            err_description=str(err_description),
            err_source=str(err_source),
        )

    def is_alive(self) -> bool:
        if self.doc is None:
            return False
        try:
            # XComponent has no isDisposed(); getImplementationName() (from
            # XServiceInfo, always present on a document) is a cheap call
            # that fails once the doc/bridge is actually gone.
            self.doc.getImplementationName()
            return True
        except Exception:
            return False

    def close(self) -> None:
        try:
            if self.doc is not None:
                self.doc.close(False)
        except Exception:
            pass
        try:
            if self.desktop is not None:
                self.desktop.terminate()
        except Exception:
            pass
