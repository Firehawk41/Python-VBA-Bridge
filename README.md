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
        Function Average(nums() As Double) As Double
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
below), and its VBA-compat layer has real, non-obvious gaps once code gets
past small self-contained snippets -- see "v2 backend" below for a real-Excel
alternative once you hit them. A real-Excel backend via `pywin32`/COM
(Windows only) exists as v2 -- see below; `VBASession`'s public API doesn't
change to use it.

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

## v2 backend: real Excel via pywin32 (Windows only)

`ExcelComBackend` drives a real, locally-installed Excel via `pywin32`/COM
automation instead of headless LibreOffice -- same `Backend` interface, so
`VBASession` and every example above works unchanged, just pass a different
backend:

```python
from vba_bridge import VBASession
from vba_bridge.backends.excel_com import ExcelComBackend

with VBASession(backend=ExcelComBackend()) as session:
    result = session.run("Function F() As Long\n    F = 42\nEnd Function\n")
```

**Status: verified against real Excel** (Windows 11, Excel 16.0) -- the
basic call/return, typed-array-argument, error-handling, iterate-on-a-fix,
`PyPrint` output, multi-module/class-module, and timeout/force-kill-recovery
paths all confirmed working end-to-end. The design mirrors the LibreOffice
backend closely (see `vba_bridge/backends/excel_com/basic_runtime.py`'s
module docstring for why it uses two workbooks -- a persistent `PyBridge`
project holding a `Core` module, and an `Agent` workbook that references it
and holds the modules injected on every `run()`).

Getting there surfaced several places where real VBA is stricter than
LibreOffice Basic (all fixed in `wrapper.py`/`basic_runtime.py`, not
user-facing):

- **No leading underscore in identifiers.** Real VBA's `Application.Run`
  can't resolve a Sub name starting with `_` -- LibreOffice's UNO invoke
  path tolerates it. Affected both the internal entry-point Sub name and
  generated local variable names for array arguments.
- **No `Option VBASupport 1`.** That's a LibreOffice-only pragma; real VBA
  doesn't recognize it at all, and leaving it in breaks compilation of the
  whole module. Stripped before injection on this backend.
- **Fixed-size arrays can't be `ReDim`'d.** `Dim arr(63) As String` is a
  fixed array in real VBA; only a dynamic `Dim arr() As String` can be
  `ReDim`'d later. LibreOffice Basic tolerates `ReDim` on either.
- **Multi-argument `Sub` calls need `Call`.** `Foo(a, b, c)` without `Call`
  is a real VBA syntax error for more than one argument (a single argument
  in parens without `Call` is fine).
- **A new VBA component isn't always a blank slate.** When Excel's VBE
  option "Require Variable Declaration" (off by default, but per-machine)
  is on, every new module is pre-seeded with its own `Option Explicit`
  line before injection ever touches it; adding another produces a
  duplicate-statement compile error. Cleared via `CodeModule.DeleteLines()`
  before injecting, regardless of that setting.
- **A compile error in injected code can pop an interactive, uncatchable
  "Compile error" dialog** instead of failing `Application.Run` fast with a
  COM exception -- confirmed specifically for user code sharing a module
  with the cross-project `PyBridge.Core.*` calls. `On Error` can never trap
  a compile-time error regardless. There's no way to detect this faster
  than a run's timeout; the existing timeout/force-kill path (below)
  recovers cleanly afterward, but a syntax mistake in new, untested VBA
  will hang for the full configured timeout first. Consider passing a
  shorter `timeout=` while iterating on unfamiliar code with this backend.
- **Releasing COM objects against a wedged apartment blocks forever.**
  `run_macro()`'s timeout path force-kills the OS process by PID specifically
  *before* dropping the last Python references to the Agent/Core workbook
  COM wrappers -- dropping them first (even after the process is already
  confirmed stuck) lets `Release()` block indefinitely waiting on a
  single-threaded apartment that can never service it.

Cross-project qualified calls (`PyBridge.Core.PyBridge_Reset(...)`) via the
Agent workbook's `VBProject.References`, the COM cross-thread marshaling in
`ExcelComBackend.run_macro()` (`CoMarshalInterThreadInterfaceInStream`/
`CoGetInterfaceAndReleaseStream`), and the `Application.Run(...)` string
format / `vbext_ComponentType` literals all work as documented once the
above were fixed.

### Setup

1. Windows, with Excel installed.
2. `pip install vba-bridge[excel]` (installs `pywin32`).
3. One-time, per-machine: enable **File > Options > Trust Center > Trust
   Center Settings > Macro Settings > "Trust access to the VBA project
   object model"** in Excel. `ExcelComBackend` injects code by manipulating
   `VBComponents` directly, which Excel blocks without this -- it can't be
   turned on programmatically (that would itself be a security hole), so it
   has to be a manual, deliberate step. `connect()` checks for this and
   raises `VbomAccessDeniedError` with the same instructions if it's off.

`ExcelComBackend(visible=True)` (the default) leaves the Excel window
on-screen while it works -- deliberate, so you can watch what it's doing,
since this drives your real Excel install rather than a disposable sandbox
process. Pass `visible=False` once you trust it.

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
- **Linux/LibreOffice only for v1.** A real-Excel backend via `pywin32`
  (Windows only) exists as v2, behind the same `Backend` interface, so
  `VBASession` doesn't need to change to use it -- see "v2 backend" above.
  Recommended once code outgrows small, self-contained snippets: LibreOffice's
  VBA-compat layer has real gaps (see the rest of this list) that cost real
  time to triage against a large or idiom-heavy real-world program, and none
  of them exist in real Excel.
- **Class modules: confirmed working, not exhaustively.** `Property Get`/`Let`,
  multiple independent instances, and re-running with updated class source all
  work. Not yet exercised: `Property Set` (object-reference properties),
  `Class_Initialize`/`Class_Terminate`, `WithEvents`, or collections/arrays of
  class instances -- these may work but haven't been tested.
- **`Implements` (interface polymorphism) does not work -- confirmed a
  genuine LibreOffice Basic limitation, not fixable from vba_bridge.** A
  class declaring `Implements SomeInterface` compiles and its own methods
  are callable directly, but the interface relationship itself never
  registers: `TypeOf obj Is SomeInterface` is always `False` and
  `TypeName(New ThatClass)` reports generic `"Object"` instead of the real
  class name, so `Set var = New ThatClass` into an interface-typed `var`
  fails with runtime error 425 ("Invalid use of an object"). Confirmed this
  isn't an artifact of vba_bridge's own module injection -- the same failure
  happens in the document's own default `Standard` library, and adding
  `Option Compatible` doesn't change it either. Encapsulation and
  composition (one class holding/using another concretely) are unaffected;
  this is specifically interface-based polymorphism that doesn't work.
  **Workaround for polymorphism-shaped code**: `CallByName(obj, "MethodName",
  VbMethod, args...)` gives working duck-typed dispatch across differently-typed
  objects that share a method name/signature convention, with no `Implements`
  or shared interface type needed -- confirmed working (an array of `Dog`/`Cat`
  instances, each with its own `Speak` method, dispatched correctly through
  one loop calling `CallByName(animals(i), "Speak", VbMethod)`).

## Testing

```
pytest tests/unit           # fast, no LibreOffice/Excel needed
pytest tests/integration    # starts a real headless LibreOffice instance
```

Integration tests are automatically skipped if `soffice` isn't on `PATH`.
`tests/unit` includes `ExcelComBackend`'s tests -- they run anywhere (Linux
included) against an in-memory fake of the COM object graph
(`tests/unit/fakes_excel_com.py`), which verifies the backend's own logic
but is not a substitute for actually running it against real Excel; there
are no `tests/integration` tests for it yet since that requires a Windows
machine to run on.

## License

[MIT](LICENSE)
