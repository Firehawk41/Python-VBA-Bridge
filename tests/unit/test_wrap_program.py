import pytest

from vba_bridge import wrapper


def test_resolve_entry_point_auto_detects_first_across_modules():
    modules = {
        "ModuleA": "Dim x As Integer\n",  # no proc here
        "ModuleB": "Function DoThing() As Long\nEnd Function\n",
    }
    module_name, proc_name, is_function = wrapper.resolve_program_entry_point(modules)
    assert module_name == "ModuleB"
    assert proc_name == "DoThing"
    assert is_function is True


def test_resolve_entry_point_qualified():
    modules = {
        "ModuleA": "Function First() As Long\nEnd Function\n",
        "ModuleB": "Sub Second()\nEnd Sub\n",
    }
    module_name, proc_name, is_function = wrapper.resolve_program_entry_point(
        modules, entry_point="ModuleB.Second"
    )
    assert module_name == "ModuleB"
    assert proc_name == "Second"
    assert is_function is False


def test_resolve_entry_point_bare_name_searched_across_modules():
    modules = {
        "ModuleA": "Function Helper() As Long\nEnd Function\n",
        "ModuleB": "Function Target() As Long\nEnd Function\n",
    }
    module_name, proc_name, is_function = wrapper.resolve_program_entry_point(
        modules, entry_point="Target"
    )
    assert module_name == "ModuleB"


def test_resolve_entry_point_skips_class_modules_when_auto_detecting():
    modules = {
        "MyClass": "Public Function Method() As Long\nEnd Function\n",
        "ModuleA": "Function RealEntry() As Long\nEnd Function\n",
    }
    module_name, proc_name, _ = wrapper.resolve_program_entry_point(
        modules, class_modules=["MyClass"]
    )
    assert module_name == "ModuleA"
    assert proc_name == "RealEntry"


def test_resolve_entry_point_qualified_class_module_rejected():
    modules = {"MyClass": "Public Function Method() As Long\nEnd Function\n"}
    with pytest.raises(wrapper.EntryPointNotFoundError):
        wrapper.resolve_program_entry_point(
            modules, class_modules=["MyClass"], entry_point="MyClass.Method"
        )


def test_resolve_entry_point_qualified_module_not_found():
    modules = {"ModuleA": "Sub X()\nEnd Sub\n"}
    with pytest.raises(wrapper.EntryPointNotFoundError):
        wrapper.resolve_program_entry_point(modules, entry_point="Missing.X")


def test_resolve_entry_point_none_found_raises():
    modules = {"ModuleA": "Dim x As Integer\n"}
    with pytest.raises(wrapper.EntryPointNotFoundError):
        wrapper.resolve_program_entry_point(modules)


def test_wrap_program_adds_orchestrator_module():
    modules = {"ModuleA": "Function DoThing() As Long\n    DoThing = 1\nEnd Function\n"}
    injected, entry_module, entry_proc, is_function = wrapper.wrap_program(modules)
    assert entry_module == "ModuleA"
    assert entry_proc == "DoThing"
    assert is_function is True
    assert injected["ModuleA"] == modules["ModuleA"]  # verbatim, untouched
    orchestrator = injected[wrapper.ORCHESTRATOR_MODULE_NAME]
    assert "Sub PyPrint(ByVal Msg As Variant)" in orchestrator
    assert "PyBridge.Core.PyBridge_SetSuccess(DoThing())" in orchestrator
    assert "On Error GoTo ErrHandler" in orchestrator


def test_wrap_program_call_is_always_unqualified_even_with_qualified_entry_point():
    modules = {"ModuleA": "Function DoThing() As Long\nEnd Function\n"}
    injected, *_ = wrapper.wrap_program(modules, entry_point="ModuleA.DoThing")
    orchestrator = injected[wrapper.ORCHESTRATOR_MODULE_NAME]
    assert "DoThing()" in orchestrator
    assert "ModuleA.DoThing" not in orchestrator


def test_wrap_program_sub_entry_point_uses_call():
    modules = {"ModuleA": "Sub DoThing()\nEnd Sub\n"}
    injected, *_ = wrapper.wrap_program(modules)
    orchestrator = injected[wrapper.ORCHESTRATOR_MODULE_NAME]
    assert "Call DoThing()" in orchestrator
    assert "PyBridge.Core.PyBridge_SetSuccess(Empty)" in orchestrator


def test_wrap_program_with_args_uses_typed_array_preamble():
    modules = {"ModuleA": "Function Average(ByVal nums() As Double) As Double\nEnd Function\n"}
    injected, *_ = wrapper.wrap_program(modules, args=[[1.0, 2.0, 3.0]])
    orchestrator = injected[wrapper.ORCHESTRATOR_MODULE_NAME]
    assert "Dim __pbArg0(2) As Double" in orchestrator
    assert "Average(__pbArg0)" in orchestrator


def test_wrap_program_rejects_reserved_orchestrator_name():
    modules = {wrapper.ORCHESTRATOR_MODULE_NAME: "Sub X()\nEnd Sub\n"}
    with pytest.raises(ValueError):
        wrapper.wrap_program(modules)


def test_wrap_program_class_module_source_untouched():
    modules = {
        "MyClass": "Public Function Method() As Long\n    Method = 1\nEnd Function\n",
        "Driver": "Function Run() As Long\n    Dim c As New MyClass\n    Run = c.Method()\nEnd Function\n",
    }
    injected, entry_module, entry_proc, _ = wrapper.wrap_program(
        modules, class_modules=["MyClass"]
    )
    assert entry_module == "Driver"
    assert injected["MyClass"] == modules["MyClass"]
    assert injected["Driver"] == modules["Driver"]
