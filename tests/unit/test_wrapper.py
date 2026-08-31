import pytest

from vba_bridge import wrapper


def test_detect_entry_point_function():
    src = "Function Average(ByVal nums() As Double) As Double\nEnd Function\n"
    name, is_function = wrapper.detect_entry_point(src)
    assert name == "Average"
    assert is_function is True


def test_detect_entry_point_sub():
    src = "Sub DoThing()\nEnd Sub\n"
    name, is_function = wrapper.detect_entry_point(src)
    assert name == "DoThing"
    assert is_function is False


def test_detect_entry_point_public_private_prefix():
    src = "Private Function Helper() As Long\nEnd Function\n"
    name, is_function = wrapper.detect_entry_point(src)
    assert name == "Helper"
    assert is_function is True


def test_detect_entry_point_picks_first_of_several():
    src = (
        "Function First() As Long\nEnd Function\n"
        "Function Second() As Long\nEnd Function\n"
    )
    name, _ = wrapper.detect_entry_point(src)
    assert name == "First"


def test_detect_entry_point_none_found_raises():
    with pytest.raises(wrapper.EntryPointNotFoundError):
        wrapper.detect_entry_point("Dim x As Integer\nx = 1\n")


def test_serialize_args_scalars():
    preamble, exprs = wrapper.serialize_args([3, 4.5, "hi", True, None])
    assert preamble == ""
    assert exprs == ["3", "4.5", '"hi"', "True", "Empty"]


def test_serialize_args_float_gets_decimal_point():
    _, exprs = wrapper.serialize_args([3.0])
    assert exprs == ["3.0"]


def test_serialize_args_string_escapes_quotes():
    _, exprs = wrapper.serialize_args(['he said "hi"'])
    assert exprs == ['"he said ""hi"""']


def test_serialize_args_numeric_list_builds_typed_double_array():
    preamble, exprs = wrapper.serialize_args([[1.0, 2.0, 3.0]])
    assert "Dim pbArg0(2) As Double" in preamble
    assert "pbArg0(0) = 1.0" in preamble
    assert "pbArg0(1) = 2.0" in preamble
    assert "pbArg0(2) = 3.0" in preamble
    assert exprs == ["pbArg0"]


def test_serialize_args_string_list_builds_typed_string_array():
    preamble, exprs = wrapper.serialize_args([["a", "b"]])
    assert "Dim pbArg0(1) As String" in preamble
    assert exprs == ["pbArg0"]


def test_serialize_args_empty_list_builds_empty_variant_array():
    preamble, exprs = wrapper.serialize_args([[]])
    assert "Dim pbArg0(-1) As Variant" in preamble
    assert exprs == ["pbArg0"]


def test_serialize_args_mixed_list_falls_back_to_variant():
    preamble, exprs = wrapper.serialize_args([[1, "a"]])
    assert "As Variant" in preamble


def test_wrap_module_auto_detects_entry_point():
    src = "Function AddTwo(ByVal a As Double, ByVal b As Double) As Double\n    AddTwo = a + b\nEnd Function\n"
    wrapped, entry, is_function, run_token = wrapper.wrap_module(src, args=[3.0, 4.0])
    assert entry == "AddTwo"
    assert is_function is True
    assert run_token
    assert "Option VBASupport 1" in wrapped
    assert "Sub PyPrint(ByVal Msg As Variant)" in wrapped
    assert "PyBridge.Core.PyBridge_SetSuccess(AddTwo(3.0, 4.0))" in wrapped
    assert "On Error GoTo ErrHandler" in wrapped
    assert "PyBridge.Core.PyBridge_SetError(Err.Number, Err.Description, Err.Source)" in wrapped
    assert f'PyBridge.Core.PyBridge_Reset("{run_token}")' in wrapped


def test_wrap_module_returns_fresh_token_each_call():
    src = "Sub X()\nEnd Sub\n"
    _, _, _, token1 = wrapper.wrap_module(src)
    _, _, _, token2 = wrapper.wrap_module(src)
    assert token1 != token2


def test_wrap_module_sub_entry_point_uses_call_and_empty_success():
    src = "Sub DoThing()\nEnd Sub\n"
    wrapped, entry, is_function, _ = wrapper.wrap_module(src)
    assert is_function is False
    assert "Call DoThing()" in wrapped
    assert "PyBridge.Core.PyBridge_SetSuccess(Empty)" in wrapped


def test_wrap_module_with_array_arg_uses_preamble_not_array_literal():
    src = "Function Average(ByVal nums() As Double) As Double\nEnd Function\n"
    wrapped, _, _, _ = wrapper.wrap_module(src, args=[[1.0, 2.0, 3.0, 4.0]])
    assert "Dim pbArg0(3) As Double" in wrapped
    assert "Average(pbArg0)" in wrapped
    assert "Array(" not in wrapped


def test_wrap_module_strips_class_export_header_from_pasted_source():
    pasted = (
        "VERSION 1.0 CLASS\n"
        "BEGIN\n"
        "  MultiUse = -1  'True\n"
        "END\n"
        'Attribute VB_Name = "Something"\n'
        "Function F() As Long\n    F = 1\nEnd Function\n"
    )
    wrapped, *_ = wrapper.wrap_module(pasted)
    assert "VERSION 1.0 CLASS" not in wrapped
    assert wrapped.count("Option VBASupport 1") == 1


def test_wrap_module_strips_duplicate_option_pragmas_from_pasted_source():
    pasted = (
        "Option VBASupport 1\nOption Explicit\n"
        "Function F() As Long\n    F = 1\nEnd Function\n"
    )
    wrapped, *_ = wrapper.wrap_module(pasted)
    assert wrapped.count("Option VBASupport 1") == 1
    assert wrapped.count("Option Explicit") == 1
