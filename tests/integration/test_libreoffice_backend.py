import pytest

pytestmark = pytest.mark.integration


def test_success_with_args_pyprint_and_return_value(vba_session):
    src = """Function Average(ByVal nums() As Double) As Double
    Dim total As Double, i As Long
    PyPrint "starting"
    For i = LBound(nums) To UBound(nums)
        total = total + nums(i)
    Next i
    Average = total / (UBound(nums) - LBound(nums) + 1)
    PyPrint "done"
End Function
"""
    result = vba_session.run(src, args=[[1.0, 2.0, 3.0, 4.0]])
    assert result.success is True
    assert result.return_value == 2.5
    assert result.output == ["starting", "done"]
    assert result.error is None


def test_zero_output_returns_empty_list_not_a_crash(vba_session):
    src = "Function Zero() As Long\n    Zero = 0\nEnd Function\n"
    result = vba_session.run(src)
    assert result.success is True
    assert result.output == []
    assert result.return_value == 0


def test_division_by_zero_is_caught_with_correct_error_number(vba_session):
    src = """Function DivErr() As Double
    Dim x As Double
    x = 0
    DivErr = 1 / x
End Function
"""
    result = vba_session.run(src)
    assert result.success is False
    assert result.error.number == 11
    assert "zero" in result.error.description.lower()
    assert result.raw_exception is None


def test_subscript_out_of_range_is_caught_with_correct_error_number(vba_session):
    src = """Sub BadSub()
    Dim arr(2) As Integer
    Dim idx As Integer
    idx = 5
    arr(idx) = 10
End Sub
"""
    result = vba_session.run(src)
    assert result.success is False
    assert result.error.number == 9


def test_iterate_to_fix_workflow(vba_session):
    """The core scenario the whole feature exists for: run buggy code, get a
    structured failure back, fix it, re-run, succeed."""
    buggy = """Function Divide(ByVal a As Double, ByVal b As Double) As Double
    Divide = a / b
End Function
"""
    broken = vba_session.run(buggy, args=[10.0, 0.0])
    assert broken.success is False
    assert broken.error.number == 11

    fixed = """Function Divide(ByVal a As Double, ByVal b As Double) As Double
    If b = 0 Then
        Divide = 0
    Else
        Divide = a / b
    End If
End Function
"""
    result = vba_session.run(fixed, args=[10.0, 0.0])
    assert result.success is True
    assert result.return_value == 0

    result2 = vba_session.run(fixed, args=[10.0, 4.0])
    assert result2.success is True
    assert result2.return_value == 2.5


def test_no_state_leak_between_runs(vba_session):
    src_a = """Function A() As Long
    PyPrint "from A"
    A = 1
End Function
"""
    src_b = """Function B() As Long
    B = 2
End Function
"""
    result_a = vba_session.run(src_a)
    assert result_a.output == ["from A"]

    result_b = vba_session.run(src_b)
    assert result_b.success is True
    assert result_b.return_value == 2
    assert result_b.output == []  # not leaked from src_a's PyPrint


@pytest.mark.slow
def test_soak_many_rapid_runs(vba_session):
    src = "Function F(ByVal n As Double) As Double\n    F = n * 2\nEnd Function\n"
    for i in range(100):
        result = vba_session.run(src, args=[float(i)])
        assert result.success is True
        assert result.return_value == i * 2
