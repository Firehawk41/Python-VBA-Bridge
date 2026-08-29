# Python-VBA-Bridge

Run VBA-compatible Basic code programmatically and iterate on errors -- like a
Python REPL, but for VBA. Write a `Sub`/`Function`, run it, get back a
structured result (success/failure, captured output, return value, error
number/description), fix the bug, and re-run -- all from Python, without
touching a GUI.

```python
from vba_bridge import VBASession

with VBASession() as session:
    result = session.run(
        """
        Function Average(ByVal nums() As Double) As Double
            Dim total As Double, i As Long
            For i = LBound(nums) To UBound(nums)
                total = total + nums(i)
            Next i
            Average = total / (UBound(nums) - LBound(nums) + 1)
        End Function
        """,
        args=[[1.0, 2.0, 3.0, 4.0]],
    )
    print(result.success, result.return_value)  # True 2.5
```

## v1 backend: LibreOffice (not real Excel)

This runs code through **LibreOffice Calc's Basic interpreter in VBA
compatibility mode** (`Option VBASupport 1`), driven headlessly via the
`uno` Python bindings -- not real Excel. It's close to real VBA for most
control-flow/error-handling code, but not identical (see "Known limitations"
below). A real-Excel backend via `pywin32`/COM (Windows only) is planned for
v2; `VBASession`'s public API is designed not to change when that lands.

## Setup

1. Install LibreOffice's Calc component and the Basic scripting bits:
   ```
   sudo apt-get install -y libreoffice-calc
   ```
   (`libreoffice-core` alone is not enough -- Calc itself is a separate
   package.)

2. Make sure the `uno` Python module is importable. This is the tricky part:
   `uno`/`pyuno` is not published to PyPI -- it's a compiled artifact (`pyuno.so`)
   installed alongside LibreOffice, tied to the Python ABI it was built
   against. Three supported setups, most robust first:

   - **Run under the system Python** that LibreOffice's `python3-uno` package
     targets (check with `python3 -c "import uno"`). No path hacks needed.
   - **A venv with `--system-site-packages`**, created from that same system
     Python (matching minor version), so `uno`/`pyuno.so` resolve automatically.
   - **Manual `sys.path` insertion** of LibreOffice's dist-packages directory
     (commonly `/usr/lib/python3/dist-packages`). Works only if your
     interpreter's ABI happens to match the one `pyuno.so` was built against --
     fragile, last resort. `LibreOfficeBackend.connect()` already tries this
     automatically and raises `UnoNotAvailableError` with an actionable
     message if it still fails.

3. `pip install -e .` (or just make sure `vba_bridge/` is on your `PYTHONPATH`).

## Usage

```python
from vba_bridge import VBASession

session = VBASession()  # starts a headless LibreOffice instance, reused across calls

result = session.run("""
    Function DivErr() As Double
        Dim x As Double
        x = 0
        DivErr = 1 / x
    End Function
""")
# result.success -> False
# result.error.number -> 11
# result.error.description -> "Division by zero."

fixed = session.run("""
    Function DivErr() As Double
        DivErr = 42
    End Function
""")
# fixed.success -> True
# fixed.return_value -> 42

session.close()
```

`VBASession` is also a context manager (`with VBASession() as session: ...`),
which is the recommended way to use it so the underlying process always gets
cleaned up.

### `run()` options

```python
session.run(
    vba_source,           # a str (single snippet) or dict {module_name: source} (multi-module program)
    entry_point=None,     # name to call; auto-detected (first Sub/Function) if omitted
    class_modules=(),     # dict path only: which module_name keys are VBA class modules
    args=(),              # positional args: scalars (int/float/str/bool) or lists
    timeout=None,         # per-call timeout in seconds; defaults to VBASession(run_timeout=...)
)
```

Lists passed as args become real typed VBA arrays (`Double()`, `String()`,
etc.), constructed explicitly -- not via Basic's `Array()`, which produces an
untyped Variant that breaks on typed array parameters.

### Multi-module and class-module programs

For a real program spanning several modules (and optionally classes), pass a
dict instead of a single string -- each key is a module name, each value its
full source. No PyPrint/error-handling code is spliced into your modules
(that lives in a separate generated module); each one just gets its own
`Option VBASupport 1`/`Option Explicit` if it doesn't already have them:

```python
modules = {
    "Calculator": """
        Option VBASupport 1
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
    "Driver": """
        Function UseClass() As Double
            Dim c As New Calculator
            c.Value = 21
            UseClass = c.DoubleIt()
        End Function
    """,
}

result = session.run(modules, class_modules=["Calculator"], entry_point="UseClass")
# result.return_value -> 42.0
```

- `class_modules` names which dict keys should support `New ModuleName`
  instantiation, properties, and instance state. Everything else is a plain
  standard module. Modules call each other by bare (unqualified) name, same
  as real VBA within one project -- no need to qualify calls.
- `entry_point` can be a bare name (searched across all non-class modules) or
  `"ModuleName.ProcName"` if you need to disambiguate. It must be a
  standalone `Sub`/`Function`, not a class's method -- write a small driver
  module (like `Driver` above) that instantiates the class and calls into it.
- Iterating on a bug works the same way as the single-string case: rebuild
  the dict with the fixed module's source changed, call `run()` again with
  the same `entry_point`/`args` -- other modules' state and source are
  untouched unless you also change them.

**Getting existing VBA into this shape**: export each module from the VBA
editor (right-click a module → *Export File...*) as `.bas`/`.cls` text files,
then read them into a dict -- or just paste the exported text directly, no
manual cleanup needed:

```python
import pathlib

modules = {p.stem: p.read_text() for p in pathlib.Path("my_vba_project").glob("*.bas")}
class_modules = [p.stem for p in pathlib.Path("my_vba_project").glob("*.cls")]
modules.update({p.stem: p.read_text() for p in pathlib.Path("my_vba_project").glob("*.cls")})
```

A real `.cls` export starts with a `VERSION 1.0 CLASS` / `BEGIN...END` header
block that's only meaningful to VBA's own project-file importer -- not valid
Basic when injected directly, and it's stripped automatically before
injection, along with any `Option VBASupport 1`/`Option Explicit` line the
export already has (vba_bridge adds its own). `Attribute VB_Name = "..."`
lines are harmless and left alone.

### Safety net: a run that doesn't execute raises, it never returns stale data

If a module fails to compile (a genuine syntax error, or something the
header-stripping above doesn't catch), LibreOffice Basic can silently return
without running anything at all -- with no exception on its own. Left
unguarded, that would surface as the *previous* successful call's leftover
result, silently reported as if it were this call's real answer. Every
`run()` writes a fresh, unique token at the start of execution and verifies
it on read-back; a mismatch raises `vba_bridge.exceptions.StaleRunError`
(via `result.raw_exception`) instead of a wrong answer.

### Working with worksheets and real workbook files

`Range`, `Selection`, `Cells`, `ActiveSheet`, and `ThisWorkbook` all work
against the hidden document each `VBASession` runs in -- including
`Range(...).Value = x` direct chained writes. `Cells(...)` is one narrow
exception: a direct chained write (`Cells(1,1).Value = x`) silently does
nothing; assign it to a variable first instead:

```vb
Dim c As Object
Set c = Cells(1, 1)
c.Value = 42          ' works
```

Full multi-workbook I/O also works -- `Workbooks.Open`, `.Close`, `.SaveAs`
against real files on disk, exactly as in real VBA:

```vb
Dim srcBook As Workbook, tplBook As Workbook
Set srcBook = Workbooks.Open("input.xlsx")
total = srcBook.Sheets(1).Range("A1").Value + srcBook.Sheets(1).Range("A2").Value
srcBook.Close False

Set tplBook = Workbooks.Open("template.xlsx")
tplBook.Sheets(1).Range("B1").Value = total
tplBook.SaveAs "output.xlsx", 51   ' 51 = xlOpenXMLWorkbook (.xlsx)
tplBook.Close False
```

Give the source workbook/template paths that exist on disk in this
environment (upload/attach the files, or point to wherever they already are)
and the resulting output file can be read back and verified.

### Printing output

VBA/Basic has no stdout. Call the injected `PyPrint("message")` from your
code instead of `Debug.Print`; everything you `PyPrint` during a run comes
back in `result.output` (a list of strings, in call order).

### `VBAResult` fields

- `success: bool`
- `output: list[str]` -- your `PyPrint()` calls, in order
- `return_value: Any` -- a `Function`'s return value (`None` for a `Sub`)
- `error: VBAError | None` -- `.number`, `.description`, `.source` when `success` is `False`
- `raw_exception: Exception | None` -- set instead of `error` for a transport-level
  failure (bad source with no entry point found, a run timeout, a dead backend) --
  distinct from an ordinary caught VBA error
- `duration_ms: float`
- `entry_point: str`

### Session lifecycle

- `session.reset()` -- clear shared run-state, keep the process warm
- `session.restart()` -- recycle the underlying LibreOffice process
- `session.close()` -- shut everything down (also happens automatically via `with`)

The whole point of keeping one `VBASession` around across many `run()` calls
is speed: each call only does two lightweight round-trips against an
already-running LibreOffice instance, not a fresh process launch. Relaunching
per call would make the "iterate quickly on errors" workflow unusably slow.

## Known limitations

- **Not real Excel.** `Option VBASupport 1` covers most VBA control-flow,
  error-handling, and collection semantics well, but has real gaps around
  some Excel-specific object-model members (charts, PivotTables, ActiveX
  controls), `Declare` Windows API calls, and occasional differences in
  exact error numbers/messages from real VBA. Code that works under this v1
  backend is not guaranteed to be portable to real Excel without adjustment.
- **No precise error line numbers in v1.** `VBAError.line` is always `None`
  right now. An automatic `Erl()`-based line-labeling scheme was built and
  tested, but line labels reliably break LibreOffice Basic's compiler
  whenever the same procedure also has a `Dim`/`ReDim`/`Static`/`Private`
  declaration -- which is virtually all real code -- and it fails *silently*
  (the run just does nothing, no error reported at all). Since the whole
  point of this tool is trustworthy error reporting, shipping a heuristic
  that can silently produce false negatives was judged worse than not having
  line numbers. `error.number`/`.description`/`.source` are reliable and
  are what's reported; an agent iterating on errors can still correlate
  those against the source. (The field is kept in `VBAError` for forward
  compatibility -- a future Excel/COM backend may make this more tractable.)
- **First-run latency**: launching the underlying LibreOffice process takes a
  few seconds. Keep a `VBASession` alive across many `run()` calls rather
  than creating a new one per call.
- **Linux/LibreOffice only for v1.** A real-Excel backend via `pywin32` (Windows
  only) is planned as v2, behind the same `Backend` interface, so `VBASession`
  won't need to change to use it.
- **Class modules: confirmed working, not exhaustively.** `Property Get`/`Let`,
  multiple independent instances, and re-running with updated class source all
  work. Not yet exercised: `Property Set` (object-reference properties),
  `Class_Initialize`/`Class_Terminate`, `WithEvents`, or collections/arrays of
  class instances -- these may work but haven't been tested.
- **`Implements` (interface polymorphism) does not work.** A class declaring
  `Implements SomeInterface` compiles and its own methods are callable
  directly, but `Set var = New ThatClass` where `var` is declared as the
  *interface* type fails with runtime error 425 ("Invalid use of an
  object"). Encapsulation and composition (one class holding/using another
  concretely) are unaffected -- this is specifically interface-based
  polymorphism (treating different classes uniformly through a shared
  interface type) that doesn't work.

## Testing

```
pytest tests/unit           # fast, no LibreOffice process needed
pytest tests/integration    # starts a real headless LibreOffice instance
```

Integration tests are automatically skipped if `soffice` isn't on `PATH`.
