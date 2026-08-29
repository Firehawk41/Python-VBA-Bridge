import pytest

from vba_bridge.exceptions import StaleRunError

pytestmark = pytest.mark.integration

_STANDARD_MODULE_EXPORT = '''Attribute VB_Name = "Module1"
Function AddTwo(ByVal a As Double, ByVal b As Double) As Double
    AddTwo = a + b
End Function
'''

_CLASS_MODULE_EXPORT = '''VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mValue As Double

Public Property Get Value() As Double
    Value = mValue
End Property

Public Property Let Value(ByVal v As Double)
    mValue = v
End Property

Public Function DoubleIt() As Double
    DoubleIt = mValue * 2
End Function
'''

_DRIVER_EXPORT = '''Attribute VB_Name = "Driver"
Function UseClass() As Double
    Dim c As New Calculator
    c.Value = 21
    UseClass = c.DoubleIt()
End Function
'''


def test_standard_module_export_pasted_verbatim(vba_session):
    result = vba_session.run(_STANDARD_MODULE_EXPORT, args=[3.0, 4.0])
    assert result.success is True
    assert result.return_value == 7.0


def test_class_module_export_pasted_verbatim(vba_session):
    # regression test: without header-stripping + forced Option VBASupport 1,
    # this used to silently succeed with a leftover value from a PREVIOUS
    # run instead of the real 42.0 -- see StaleRunError.
    result = vba_session.run(
        {"Calculator": _CLASS_MODULE_EXPORT, "Driver": _DRIVER_EXPORT},
        class_modules=["Calculator"],
        entry_point="UseClass",
    )
    assert result.success is True
    assert result.return_value == 42.0


def test_genuine_syntax_error_raises_stale_run_error_not_silent_wrong_result(vba_session):
    setup = vba_session.run("Function F() As Long\n    F = 999\nEnd Function\n")
    assert setup.success is True
    assert setup.return_value == 999

    broken = "Function Broken() As Long\n    Broken = (((1 + 2)\nEnd Function\n"  # unbalanced parens
    result = vba_session.run(broken)
    assert result.success is False
    assert result.return_value is None
    assert isinstance(result.raw_exception, StaleRunError)
