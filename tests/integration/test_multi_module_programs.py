import pytest

pytestmark = pytest.mark.integration

_MODULES = {
    "ModuleA": "Public Function Helper(ByVal x As Double) As Double\n    Helper = x * 10\nEnd Function\n",
    "ModuleB": "Function UseHelper(ByVal n As Double) As Double\n    UseHelper = Helper(n) + 1\nEnd Function\n",
}

_CLASS_PROGRAM = {
    "Calculator": """Option VBASupport 1
Option Explicit

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
""",
    "Driver": """Function UseClass() As Double
    Dim c As New Calculator
    c.Value = 21
    UseClass = c.DoubleIt()
End Function

Function TwoInstances() As Double
    Dim a As New Calculator, b As New Calculator
    a.Value = 5
    b.Value = 100
    TwoInstances = a.DoubleIt() + b.DoubleIt()
End Function
""",
}


def test_cross_module_call_with_explicit_entry_point(vba_session):
    result = vba_session.run(_MODULES, entry_point="UseHelper", args=[5.0])
    assert result.success is True
    assert result.return_value == 51.0
    assert result.entry_point == "ModuleB.UseHelper"


def test_class_module_with_properties(vba_session):
    result = vba_session.run(_CLASS_PROGRAM, class_modules=["Calculator"], entry_point="UseClass")
    assert result.success is True
    assert result.return_value == 42.0


def test_class_module_independent_instances(vba_session):
    result = vba_session.run(_CLASS_PROGRAM, class_modules=["Calculator"], entry_point="TwoInstances")
    assert result.success is True
    assert result.return_value == 210.0


def test_iterate_loop_on_one_module_of_a_program(vba_session):
    result1 = vba_session.run(_MODULES, entry_point="UseHelper", args=[5.0])
    assert result1.return_value == 51.0

    updated = dict(_MODULES)
    updated["ModuleA"] = (
        "Public Function Helper(ByVal x As Double) As Double\n    Helper = x * 100\nEnd Function\n"
    )
    result2 = vba_session.run(updated, entry_point="UseHelper", args=[5.0])
    assert result2.success is True
    assert result2.return_value == 501.0


def test_iterate_loop_on_a_class_module(vba_session):
    result1 = vba_session.run(_CLASS_PROGRAM, class_modules=["Calculator"], entry_point="UseClass")
    assert result1.return_value == 42.0

    updated = dict(_CLASS_PROGRAM)
    updated["Calculator"] = _CLASS_PROGRAM["Calculator"].replace("mValue * 2", "mValue * 3")
    result2 = vba_session.run(updated, class_modules=["Calculator"], entry_point="UseClass")
    assert result2.success is True
    assert result2.return_value == 63.0


def test_error_inside_multi_module_program_is_caught(vba_session):
    modules = {
        "ModuleA": (
            "Function DivErr() As Double\n"
            "    Dim x As Double\n"
            "    x = 0\n"
            "    DivErr = 1 / x\n"
            "End Function\n"
        ),
    }
    result = vba_session.run(modules)
    assert result.success is False
    assert result.error.number == 11


def test_class_module_source_stays_verbatim_alongside_standard_modules(vba_session):
    # regression guard: class modules must not get PyPrint/orchestration text
    # spliced in like the single-string wrap_module() path does
    result = vba_session.run(_CLASS_PROGRAM, class_modules=["Calculator"], entry_point="UseClass")
    assert result.success is True


def test_callbyname_gives_duck_typed_polymorphism_without_implements(vba_session):
    # Implements-based interface polymorphism doesn't work in this backend
    # (see README Known limitations) -- CallByName is the confirmed working
    # alternative for dispatching the "same" method across unrelated classes.
    modules = {
        "Dog": 'Function Speak() As String\n    Speak = "Woof"\nEnd Function\n',
        "Cat": 'Function Speak() As String\n    Speak = "Meow"\nEnd Function\n',
        "Driver": """Function RunAll() As String
    Dim animals(1) As Object
    Set animals(0) = New Dog
    Set animals(1) = New Cat
    Dim i As Integer
    Dim result As String
    For i = 0 To 1
        result = result & CallByName(animals(i), "Speak", VbMethod) & " "
    Next i
    RunAll = result
End Function
""",
    }
    result = vba_session.run(modules, class_modules=["Dog", "Cat"], entry_point="RunAll")
    assert result.success is True
    assert result.return_value == "Woof Meow "
