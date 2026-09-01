# Testing a real, existing VBA project

Bringing a real, already-deployed VBA project (exported `.bas`/`.cls` files
from an actual working macro, not written from scratch for `vba_bridge`)
under test surfaces a handful of practical issues beyond what the README's
basic examples cover. This is the accumulated recipe from doing exactly
that against a ~2000-line, multi-module, multi-class production project
with a real Access database dependency -- so the next project doesn't have
to re-pay the same discovery cost.

## Reading real exports: use cp1252, not UTF-8

Real VBA `.bas`/`.cls` exports are Windows-1252 encoded -- comments with
em-dashes, degree signs, smart quotes, etc. are common in anything that's
been edited in the VBE over the years. Reading them as UTF-8 raises
`UnicodeDecodeError` partway through the file. Always:

```python
with open(path, "r", encoding="cp1252") as f:
    source = f.read()
```

## External library references (Scripting.Dictionary, ADODB.Connection, ...)

Real-world VBA commonly declares early-bound external types --
`Dim d As Scripting.Dictionary`, `Dim conn As ADODB.Connection` -- which
requires an actual VBProject reference, not just the type name resolving
on its own. `vba_bridge` doesn't add these automatically (they're
project-specific). `ExcelComBackend.add_reference(guid, major, minor)`
adds one to the Agent workbook before you inject/run code that needs it:

```python
from vba_bridge import VBASession
from vba_bridge.backends.excel_com import (
    ExcelComBackend,
    MICROSOFT_SCRIPTING_RUNTIME,
    MICROSOFT_ACTIVEX_DATA_OBJECTS,
)

backend = ExcelComBackend()
session = VBASession(backend=backend)  # connects (auto_start=True default)
backend.add_reference(*MICROSOFT_SCRIPTING_RUNTIME)
backend.add_reference(*MICROSOFT_ACTIVEX_DATA_OBJECTS)
```

Those two constants are confirmed on a Windows 11 / Office 2016+ (x86)
machine. If a different major/minor is registered on yours, check
`HKLM:\SOFTWARE\Classes\TypeLib\{guid}` (or `...\Wow6432Node\...` when
querying a 64-bit OS about a 32-bit-only library) and adjust -- ADO in
particular is usually safe to pin at an old version like 2.8 since it
stays binary-compatible with newer installs, but that isn't guaranteed on
every machine.

## UserForms aren't supported

`ExcelComBackend` (and `LibreOfficeBackend`) can only inject standard
modules and class modules -- there is no support for injecting a real
`.frm`/`.frx` UserForm component. If a code path you need to test
references one (`Dim f As New frmSomething`, `Unload SomeForm`), either:

- **Avoid exercising that branch** with your test input. This is the
  common case -- UserForm usage is very often a fallback path for "not
  found in a lookup," easy to dodge by making sure your test data already
  matches something in whatever it's looking up against.
- **If the branch is genuinely unavoidable**, replace it in your in-memory
  copy of the source with a targeted string replacement before injecting
  -- never edit the real file on disk. For a branch that your test data
  should never actually reach, replace it with a loud `Err.Raise` rather
  than a silent no-op, so a wrong assumption about reachability fails
  loudly instead of masking a real gap:

```python
old_block = '''        If Ambiguous Then
            With New SomeUserForm
                .Show vbModal
                ...
            End With
        End If'''
new_block = '''        If Ambiguous Then
            Err.Raise vbObjectError + 999, "MyFunc", "SomeUserForm is not available in this test harness; this branch should not be reached by the test's synthetic data."
        End If'''
assert old_block in source, "block text did not match -- re-check against the current file"
source = source.replace(old_block, new_block)
```

## Calling a `Private` entry point

The function you actually want to exercise is often `Private` -- the real
work behind a `Public` wrapper that also does things you can't automate
(`Application.GetSaveAsFilename`, a closing `MsgBox`, etc., both of which
show a real modal dialog that hangs unattended automation regardless of
any "debug mode" flag the project has). You can't call a `Private`
procedure across modules -- cross-module calls require `Public`. Append a
small `Public` driver `Sub` to your in-memory copy of *that same module's*
source; being in the same module, it can call the `Private` function
directly:

```python
driver = '''
Public Sub RunForTest()
    Call TheRealPrivateFunction(...)
End Sub
'''
real_source = real_source + driver
# inject real_source under its real module name; entry_point="ThatModule.RunForTest"
```

This exercises the real logic while skipping only the interactive parts
you genuinely can't automate, without touching the file on disk or the
function you're actually trying to verify.

## Isolating from production config

Real projects often hardcode paths/credentials in a config module
(sometimes deliberately gitignored). Never edit that file on disk for a
test run -- apply a targeted string replacement to an in-memory copy
instead, and **assert the replacement actually matched** before
proceeding. A silent no-op replacement (the text you expected isn't
there, e.g. because the file changed) is worse than an error, because it
means your "isolated" test just quietly ran against production:

```python
config_source = config_source.replace(
    'Public Const DB_PATH As String = "\\\\PRODSERVER\\real.accdb"',
    f'Public Const DB_PATH As String = "{local_mock_db_path}"',
)
assert local_mock_db_path in config_source, "override did not match -- check the real file's exact text"
```

## Diagnosing hangs: use the project's own logging, not live window inspection

If the project already has its own logging class, redirect its output
path (same in-memory-override technique as above) to a local folder and
read the newest log file when a run hangs or fails. It shows exactly
which named function logged last before the problem -- far faster than
live-diagnosing a hung Excel dialog by enumerating Win32 windows and
reading control text out of a modal box. (A compile error in real Excel
can pop an interactive, un-catchable dialog instead of failing
`Application.Run` fast with a COM exception -- see the main README's "v2
backend" section. When that happens there's no log line at all past the
last one written, which itself tells you the break happened during
compilation/setup of whatever ran next, not inside logged business logic.)

## A local mock Access database, without going anywhere near production

If the project talks to a real Access database, build a local `.accdb`
mock instead of pointing anywhere near production -- even a read-only
connection to the real file is a risk not worth taking just to verify
code paths.

**Bitness gotcha**: on a machine where Office is 32-bit, ADOX/ACE OLEDB is
only registered for 32-bit COM. A 64-bit Python process trying
`win32com.client.Dispatch("ADOX.Catalog")` to *create* a new `.accdb` (as
opposed to just opening an existing one via `ADODB.Connection`, which
works fine either bitness) fails with "Class not registered" -- not an
informative error about *why*. Check your Office bitness
(`HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration`, `Platform`
value) and, if it's x86, run the DB-creation step under 32-bit PowerShell
instead of Python:

```powershell
& "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy Bypass -Command {
    $dbPath = 'C:\path\to\mock.accdb'
    $cat = New-Object -ComObject ADOX.Catalog
    $cat.Create('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + $dbPath)
    $conn = New-Object -ComObject ADODB.Connection
    $conn.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + $dbPath) | Out-Null
    $conn.Execute("CREATE TABLE elements (ID COUNTER PRIMARY KEY, symbol TEXT(10))") | Out-Null
    $conn.Execute("INSERT INTO elements (symbol) VALUES ('Fe')") | Out-Null
    $conn.Close()
}
```

**Schema**: derive it from the actual SQL the VBA issues (grep the project
for `SELECT`/`INSERT INTO`/`FROM \w+`) rather than guessing table shapes.

**Seed at least one matching row per table a lazy-loaded dictionary/cache
joins across.** An empty *result* (not just an empty table) from a query
some code expects rows from is a common way to hit an unrelated bug in
the code under test before you even reach what you're actually trying to
verify -- e.g. VBA's `UBound(arr, 2)` on a 1-D empty array (what
`Recordset.GetRows()` returns for zero matching rows) raises "Subscript
out of range." That's a real bug in the code being tested, not something
`vba_bridge` can paper over -- but it's also not what you're usually
trying to test on a given run, so seed data to route around it
deliberately, and separately report the underlying fragility.
