"""Lightweight in-memory stand-ins for the pywin32 COM object graph
(Application/Workbooks/Workbook/VBProject/VBComponents/...) that
vba_bridge.backends.excel_com touches -- same spirit as FakeBackend in
test_session_fake_backend.py: enough behavior to exercise the real logic in
basic_runtime.py without a real Excel/COM install, not a full COM emulator.
"""


class FakeCodeModule:
    def __init__(self):
        self.added_source = None
        self.delete_lines_calls = []

    @property
    def CountOfLines(self):
        return 0 if self.added_source is None else len(self.added_source.splitlines())

    def DeleteLines(self, start, count):
        self.delete_lines_calls.append((start, count))
        self.added_source = None

    def AddFromString(self, source):
        self.added_source = source


class FakeComponent:
    def __init__(self, name, component_type):
        self.Name = name
        self.component_type = component_type
        self.CodeModule = FakeCodeModule()


class FakeComponents:
    def __init__(self):
        self.items = []
        self.add_calls = []
        self.remove_calls = []

    def __iter__(self):
        return iter(self.items)

    def Add(self, component_type):
        component = FakeComponent(f"_new_{len(self.items)}", component_type)
        self.items.append(component)
        self.add_calls.append(component_type)
        return component

    def Remove(self, component):
        self.items.remove(component)
        self.remove_calls.append(component.Name)


class FakeReferences:
    def __init__(self):
        self.added_files = []

    def AddFromFile(self, path):
        self.added_files.append(path)


class FakeVBProject:
    def __init__(self):
        self.Name = "VBAProject"
        self.VBComponents = FakeComponents()
        self.References = FakeReferences()


class FakeWorkbook:
    def __init__(self, name):
        self._name = name
        self.FullName = f"C:\\fake\\{name}"
        self.VBProject = FakeVBProject()
        self.saved_as = []
        self.saved = False
        self.closed = False
        self.close_save_changes = None
        self.alive = True

    @property
    def Name(self):
        # Real COM: touching a property on a closed workbook raises -- mirror
        # that so is_alive()'s try/except has something real to catch.
        if not self.alive:
            raise RuntimeError("workbook is closed")
        return self._name

    def SaveAs(self, path, FileFormat=None):
        self.saved_as.append((path, FileFormat))
        self.FullName = path

    def Save(self):
        self.saved = True

    def Close(self, SaveChanges=None):
        self.closed = True
        self.close_save_changes = SaveChanges
        self.alive = False


class FakeWorkbooks:
    def __init__(self):
        self.created = []

    def Add(self):
        workbook = FakeWorkbook(f"Book{len(self.created) + 1}")
        self.created.append(workbook)
        return workbook


class FakeApplication:
    def __init__(self):
        self.Workbooks = FakeWorkbooks()
        self.run_calls = []
        self.run_results = []

    def Run(self, run_string):
        self.run_calls.append(run_string)
        if self.run_results:
            return self.run_results.pop(0)
        return None
