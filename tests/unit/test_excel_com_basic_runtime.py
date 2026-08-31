import pytest

from tests.unit.fakes_excel_com import FakeApplication
from vba_bridge.backends.excel_com import basic_runtime
from vba_bridge.exceptions import StaleRunError


@pytest.fixture
def runtime():
    app = FakeApplication()
    rt = basic_runtime.VBAProjectRuntime(app)
    rt.setup()
    yield rt
    rt.close()


def test_setup_creates_core_then_agent_workbook(runtime):
    created = runtime.application.Workbooks.created
    assert len(created) == 2
    assert created[0] is runtime.core_workbook
    assert created[1] is runtime.agent_workbook


def test_setup_saves_core_workbook_as_macro_enabled(runtime):
    path, file_format = runtime.core_workbook.saved_as[0]
    assert path.endswith(".xlsm")
    assert file_format == basic_runtime._XL_MACRO_ENABLED_WORKBOOK


def test_setup_renames_core_project_to_pybridge(runtime):
    assert runtime.core_workbook.VBProject.Name == basic_runtime.CORE_PROJECT_NAME


def test_setup_saves_core_project_before_agent_references_it(runtime):
    # References.AddFromFile resolves a path on disk -- if Save() hadn't
    # happened yet this would be referencing a file that doesn't exist.
    assert runtime.core_workbook.saved is True
    assert runtime.agent_workbook.VBProject.References.added_files == [
        runtime.core_workbook.FullName
    ]


def test_setup_adds_core_module_with_pybridge_source(runtime):
    components = runtime.core_workbook.VBProject.VBComponents.items
    assert len(components) == 1
    core_component = components[0]
    assert core_component.Name == basic_runtime.CORE_MODULE_NAME
    assert core_component.component_type == basic_runtime._VBEXT_CT_STDMODULE
    assert core_component.CodeModule.added_source == basic_runtime.CORE_MODULE_SOURCE


def test_inject_module_adds_new_standard_module(runtime):
    runtime.inject_module("Main", "Sub Foo()\nEnd Sub\n")
    components = runtime.agent_workbook.VBProject.VBComponents.items
    assert len(components) == 1
    assert components[0].Name == "Main"
    assert components[0].component_type == basic_runtime._VBEXT_CT_STDMODULE
    assert components[0].CodeModule.added_source == "Sub Foo()\nEnd Sub\n"


def test_inject_module_marks_class_modules(runtime):
    runtime.inject_module("MyClass", "Public Function M()\nEnd Function\n", is_class=True)
    component = runtime.agent_workbook.VBProject.VBComponents.items[0]
    assert component.component_type == basic_runtime._VBEXT_CT_CLASSMODULE


def test_inject_module_replaces_existing_module_by_remove_then_add(runtime):
    components = runtime.agent_workbook.VBProject.VBComponents
    runtime.inject_module("Main", "Sub V1()\nEnd Sub\n")
    first_component = components.items[0]

    runtime.inject_module("Main", "Sub V2()\nEnd Sub\n")

    assert components.remove_calls == ["Main"]
    assert len(components.items) == 1
    assert components.items[0] is not first_component
    assert components.items[0].Name == "Main"
    assert components.items[0].CodeModule.added_source == "Sub V2()\nEnd Sub\n"


def test_run_invokes_entry_sub_then_reads_packed_result(runtime):
    runtime.application.run_results = [
        None,  # PyBridgeRun's own return value is unused
        (True, 0, "", "", 7.0, 1, ["hello"], "tok-123"),
    ]
    result = runtime.run("Main", expected_token="tok-123")

    assert runtime.application.run_calls == [
        f"'{runtime.agent_workbook.Name}'!Main.PyBridgeRun",
        f"'{runtime.core_workbook.Name}'!Core.PyBridge_GetResultPacked",
    ]
    assert result.success is True
    assert result.output == ["hello"]
    assert result.return_value == 7.0


def test_run_raises_stale_run_error_on_token_mismatch(runtime):
    runtime.application.run_results = [
        None,
        (True, 0, "", "", 1.0, 0, [], "old-token"),
    ]
    with pytest.raises(StaleRunError):
        runtime.run("Main", expected_token="new-token")


def test_run_failure_reports_error_fields_and_no_return_value(runtime):
    runtime.application.run_results = [
        None,
        (False, 11, "Division by zero.", "Main", 0.0, 0, [], "tok"),
    ]
    result = runtime.run("Main", expected_token="tok")
    assert result.success is False
    assert result.err_number == 11
    assert result.err_description == "Division by zero."
    assert result.return_value is None


def test_is_alive_false_before_setup():
    rt = basic_runtime.VBAProjectRuntime(FakeApplication())
    assert rt.is_alive() is False


def test_is_alive_true_after_setup(runtime):
    assert runtime.is_alive() is True


def test_is_alive_false_after_close(runtime):
    runtime.close()
    assert runtime.is_alive() is False


def test_close_closes_both_workbooks_without_saving(runtime):
    core_wb, agent_wb = runtime.core_workbook, runtime.agent_workbook
    runtime.close()
    assert agent_wb.closed is True
    assert agent_wb.close_save_changes is False
    assert core_wb.closed is True
    assert core_wb.close_save_changes is False
