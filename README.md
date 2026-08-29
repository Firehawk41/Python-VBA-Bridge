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
    vba_source,           # a Sub or Function definition (plus any helpers it calls)
    entry_point=None,     # name to call; auto-detected (first Sub/Function) if omitted
    args=(),              # positional args: scalars (int/float/str/bool) or lists
    timeout=None,         # per-call timeout in seconds; defaults to VBASession(run_timeout=...)
)
```

Lists passed as args become real typed VBA arrays (`Double()`, `String()`,
etc.), constructed explicitly -- not via Basic's `Array()`, which produces an
untyped Variant that breaks on typed array parameters.

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

## Testing

```
pytest tests/unit           # fast, no LibreOffice process needed
pytest tests/integration    # starts a real headless LibreOffice instance
```

Integration tests are automatically skipped if `soffice` isn't on `PATH`.
