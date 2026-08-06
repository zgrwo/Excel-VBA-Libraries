---
name: python-SKILL
description: >
  Python development for Excel COM automation, numpy/scipy cross-validation,
  and test infrastructure. Use when writing pywin32 Excel automation scripts,
  numpy/scipy statistical validation, openpyxl data handling, or Python-VBA
  cross-validation tests. Covers COM threading safety, SafeArray marshaling,
  datetime timezone handling, and Python-VBA type bridging.
---

# Python — Excel COM Automation & Test Infrastructure

Battle-tested patterns from the Excel VBA Libraries test suite.

## 1. Core Principles

- **COM-first testing**: All VBA validation goes through pywin32 COM, not direct VBA execution.
  This tests the actual UDF call path — cell → VBA → return value — the same way Excel does.
- **Python as referee, not implementation**: Cross-validation compares VBA output against
  numpy/scipy reference values. Python never reimplements VBA logic; it independently computes
  the expected answer.
- **Cleanup is non-negotiable**: Orphaned Excel processes from failed test runs accumulate and
  consume memory. Every script must use `try/finally` teardown — no exceptions.
- **Late binding in VBA, early dispatch in Python**: VBA modules use `CreateObject` for portability;
  Python uses `win32.gencache.EnsureDispatch` for type-library caching and speed.

## 2. Excel COM Automation (pywin32)

### 2.1 Startup Pattern

```python
import win32com.client as win32

def ensure_excel():
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False
    excel.ScreenUpdating = False
    excel.AutomationSecurity = 1  # msoAutomationSecurityLow
    return excel
```

Always disable `ScreenUpdating` and `EnableEvents` — they generate unnecessary COM round-trips.
Use `EnsureDispatch` not `Dispatch` — it caches the type library and avoids `makepy` regeneration.

### 2.2 Workbook Creation with VBA Module Injection

```python
def create_workbook(excel, output_path, source_paths):
    wb = excel.Workbooks.Add()
    ws = wb.Sheets(1)
    ws.Name = "TestResults"
    wb.Sheets.Add().Name = "TestData"

    vbproj = wb.VBProject
    for src_path in source_paths:
        mod_name = os.path.splitext(os.path.basename(src_path))[0]
        for comp in list(vbproj.VBComponents):
            if comp.Name == mod_name:
                vbproj.VBComponents.Remove(comp)
        content = read_module(src_path)  # strips Attribute VB_Name line
        comp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
        comp.Name = mod_name
        comp.CodeModule.AddFromString(content)

    wb.SaveAs(output_path, FileFormat=52)  # xlOpenXMLWorkbookMacroEnabled
    wb.Close(SaveChanges=True)
    return excel.Workbooks.Open(output_path)
```

**Critical:** `FileFormat=52` is `xlOpenXMLWorkbookMacroEnabled` (.xlsm). Without it, macros are stripped on save.

### 2.3 Running VBA Macros and Reading Results

```python
def run_tests(excel, wb):
    wb.Sheets("TestResults").Activate()
    excel.Application.Run("TestRunner.RunAllTests")

def read_results(wb):
    ws = wb.Sheets("TestResults")
    last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # xlUp
    for r in range(2, last_row + 1):
        result_val = str(ws.Cells(r, 5).Value or "")
        if result_val == "PASS": ...
        elif result_val == "FAIL": ...
```

Use `End(-4162)` (xlUp) instead of `UsedRange` — `UsedRange` can retain stale cell formatting.

## 3. COM Type Bridging Gotchas

### 3.1 Datetime Marshaling

**Recommended approach — ISO date string:**
```python
from datetime import date, datetime

def _to_com_arg(arg):
    """datetime/date → ISO date string. VBA CDate() parses unambiguously."""
    if isinstance(arg, (date, datetime)):
        return arg.strftime("%Y-%m-%d")
    return arg
```

Why ISO strings: naive `datetime` → COM shifts by local UTC offset (CST=+8h shifts 1 day).
Excel serials are complex (`(d-epoch).days+2`). ISO `"2024-01-15"` → VBA `CDate()` works everywhere.

Apply before `Application.Run`: `args = tuple(_to_com_arg(a) for a in args)`.
Also convert in VBA: functions receiving dates should use `ByVal d As Variant` + `IsDate(d)` + `CDate(d)`.

### 3.2 Array Return Types

```python
# VBA 2D arrays come back as tuple-of-tuples
result = excel.Application.Run("SomeUDF", rng)
# result = ((1.0, 2.0), (3.0, 4.0))  — tuple of row tuples

# Convert to numpy:
import numpy as np
def com_to_numpy(vba_result):
    if isinstance(vba_result, np.ndarray):
        return vba_result
    rows = len(vba_result)
    cols = len(vba_result[0]) if isinstance(vba_result[0], tuple) else 1
    out = np.zeros((rows, cols))
    for i in range(rows):
        row = vba_result[i]
        if isinstance(row, tuple):
            for j in range(cols):
                out[i, j] = float(row[j])
        else:
            out[i, 0] = float(row)
    return out
```

### 3.3 UDF Support — is_udf Flag + Multi-Range Writing

Functions taking `ByRef rng As Range` need data written to a sheet before calling.
The test framework supports this via `"is_udf": True` in test case dicts:

```python
# Test case with is_udf — args written to TestData sheet as Range objects
{"name": "GroupBy_SUM", "func": "GroupBy",
 "args": lambda: (SALES_HDR, 1, 3, "SUM"),
 "py_ref": lambda a: _py_groupby(a[0], a[1], a[2], a[3]),
 "result_type": "array", "is_udf": True},
```

**Multi-Range UDFs** (e.g., CrossJoin takes two Ranges): the `_call_udf` method
writes each array arg to a separate sheet area with a blank row separator.
Scalar args pass through unchanged.

**Key implementation detail in `_call_udf`:**
```python
r_offset = 1
for a in args:
    if isinstance(a, (list, tuple)):
        arr = np.asarray(a, dtype=object)
        write_range(ws, arr, r_offset, 1)
        nr, nc = arr.shape
        rng_args.append(ws.Range(ws.Cells(r_offset, 1), ws.Cells(r_offset + nr - 1, nc)))
        r_offset += nr + 1  # blank row separator
    else:
        rng_args.append(a)
return run_macro(excel, wb, macro, *rng_args)
```

### 3.4 Array Comparison Patterns

#### Empty Array Guard
`np.max` on zero-size arrays raises ValueError. Always guard before comparison:
```python
if va.size == 0 and pa.size == 0:
    me, ok = 0.0, True
elif va.size == 0 or pa.size == 0:
    me, ok = float("inf"), False
else:
    me = float(np.nanmax(np.abs(np.nan_to_num(va) - np.nan_to_num(pa))))
    ok = me <= tol
```

#### String Array Detection
COM returns tuples-of-tuples that may contain strings. Detect before numeric conversion:
```python
is_str = False
if isinstance(vba_result, (tuple, list)) and len(vba_result) > 0:
    first = vba_result[0]
    if isinstance(first, (tuple, list)) and len(first) > 0:
        first = first[0]
    if isinstance(first, str):  # note: do NOT add 'and first' — "" is falsy!
        try: float(first)
        except (ValueError, TypeError): is_str = True
```
**Trap:** `if isinstance(first, str) and first:` skips empty strings (`""` is falsy).
Remove the `and first` guard — an empty string IS a string and should trigger string compare.

#### VBA Empty → COM None
`Array()` in VBA returns Empty, which COM marshals to `None`.
Handle in `com_to_numpy`: `if vba_result is None: return np.array([])`.

#### NaN Padding (ArrayChunk)
VBA `Empty` → Python `None` → numpy `NaN`. Use `np.nanmax` + `np.nan_to_num`
to compare arrays with padding: `np.nanmax(np.abs(np.nan_to_num(va) - np.nan_to_num(pa)))`.

#### Date Comparison in Scalars
VBA `Date` returns → COM `pywintypes.datetime`. Compare dates directly:
```python
if isinstance(vba_result, (date, datetime)):
    vba_dt = datetime(vba_result.year, vba_result.month, vba_result.day)
    py_dt = datetime(py_val.year, py_val.month, py_val.day)
    ok = abs((vba_dt - py_dt).total_seconds() / 86400.0) < 1.0
```

## 4. numpy/scipy Cross-Validation

### 4.1 Statistical Functions — Python Reference

| VBA Function | Python Equivalent | Tolerance |
|-------------|-------------------|-----------|
| Mean, Median, StdDev, Var | `np.mean`, `np.median`, `np.std(ddof=1)`, `np.var(ddof=1)` | 1e-6 |
| Skewness, Kurtosis | `scipy.stats.skew(bias=False)`, `scipy.stats.kurtosis(bias=False)` | 1e-2 |
| GeometricMean, HarmonicMean | `scipy.stats.gmean`, `scipy.stats.hmean` | 1e-6 |
| Percentile | `np.percentile` | 1e-4 |

### 4.2 Linear Algebra — Python Reference

| VBA Function | Python Equivalent | Tolerance |
|-------------|-------------------|-----------|
| SVD | `np.linalg.svd(full_matrices=False)` | 1e-5 (sigma), 1e-8 (reconstruction) |
| Determinant | `np.linalg.det` | 1e-6 |
| QR Decomposition | `np.linalg.qr` | 1e-8 (reconstruction) |
| Pseudoinverse | `np.linalg.pinv` | 1e-5 |

### 4.3 Floating-Point Comparison

```python
def check_close(name, vba_val, py_val, tol=1e-8):
    try:
        v, p = float(vba_val), float(py_val)
        ok = abs(v - p) <= tol or (np.isnan(v) and np.isnan(p))
    except (ValueError, TypeError, OverflowError):
        ok = str(vba_val).strip() == str(py_val).strip()
```

Always handle `NaN` and non-numeric cases. `float("NaN") == float("NaN")` is `False` in Python.

## 5. Excel Data Reading (openpyxl)

```python
import openpyxl

def get_numeric_column(xlsx_path, sheet, col_name):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet]
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    ci = headers.index(col_name) + 1
    vals = []
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, ci).value
        if v is None or isinstance(v, bool):
            vals.append(np.nan)
        elif isinstance(v, (int, float)):
            vals.append(float(v))
        else:
            vals.append(np.nan)
    wb.close()
    return np.array(vals)
```

Always use `data_only=True` — otherwise formulas return cached values or None.

## 6. Cleanup Discipline

```python
def teardown(excel, wb):
    try:
        wb.Close(SaveChanges=True)
    except Exception:
        pass
    finally:
        excel.Quit()
        del excel  # release COM reference
```

**Never skip `excel.Quit()`** — orphaned Excel processes accumulate and consume memory. Use `try/finally` to ensure cleanup even on failure.

## 7. VBA Code in Python Strings

VBA test code is embedded in Python `build_*.py` files using triple-quoted raw strings:

```python
def build_test_module():
    return r"""
' TestRunner module
Option Explicit
...
"""
```

**Critical gotcha:** VBA double-quote sequences (`""""`) collide with Python `"""` terminators.
Use `Chr$(34)` instead of consecutive double-quotes in VBA strings embedded in Python `r"""..."""`.

## 8. Module Build Pattern

Each module has a `build_*.py` that follows this structure:
1. `ensure_excel()` → start Excel
2. `create_workbook()` → new .xlsm + import .bas modules
3. `add_test_module()` → inject TestRunner VBA code
4. `run_tests()` → execute `TestRunner.RunAllTests` macro
5. `read_results()` → parse TestResults sheet
6. `print_report()` → display results
7. `teardown()` → close Excel

## 9. Related Skills

其他 Skill 文件及文档的加载时机见 [AGENTS.md](../../../AGENTS.md) 文档路由表。
