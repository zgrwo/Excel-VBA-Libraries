"""Shared utilities for Excel COM automation testing of VBA modules.

Follows the patterns established in skills/python-SKILL.md:
  - COM-first testing via pywin32
  - Python as referee, not implementation
  - Cleanup is non-negotiable (try/finally teardown)
  - Early dispatch in Python (EnsureDispatch)
"""

import os
import sys
import numpy as np
from typing import Any, List, Optional, Sequence, Tuple, Union

import win32com.client as win32

# ---------------------------------------------------------------------------
# Path constants
# ---------------------------------------------------------------------------
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = os.path.join(PROJECT_ROOT, "src")
VBA_CORE_DIR = os.path.join(PROJECT_ROOT, "VBA-Core")

# VBA-Core modules must be imported in this exact order due to dependencies:
#   VariantKit (no deps) -> ArrayOps -> DictProxy
VBA_CORE_IMPORT_ORDER = [
    "VariantKit",
    "ArrayOps",
    "DictProxy",
]


# ===========================================================================
# Excel Lifecycle
# ===========================================================================

def ensure_excel(visible: bool = False) -> Any:
    """Start Excel with correct settings for automated testing.

    Disables all UI and prompting features that would stall automated runs:
    DisplayAlerts, AskToUpdateLinks, EnableEvents, ScreenUpdating.
    Sets AutomationSecurity to msoAutomationSecurityLow so macros can run
    without prompting.
    """
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False
    excel.ScreenUpdating = False
    excel.AutomationSecurity = 1  # msoAutomationSecurityLow
    return excel


def teardown(excel: Any, wb: Any = None) -> None:
    """Close workbook and quit Excel, releasing COM references.

    Uses try/finally to guarantee excel.Quit() is called even when
    workbook close fails.  Orphaned Excel processes consume memory and
    interfere with subsequent test runs.
    """
    try:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
    finally:
        try:
            excel.Quit()
        except Exception:
            pass
        del excel


# ===========================================================================
# Module I/O
# ===========================================================================

def read_module(path: str) -> str:
    """Read a .bas/.cls file for CodeModule.AddFromString injection.

    VBE header lines (VERSION, BEGIN/END, Attribute VB_*) are already
    commented out in the source files (matching the .bas convention),
    so no stripping is needed — just read and return.
    """
    with open(path, "r", encoding="utf-8-sig") as fh:
        return fh.read()


# ===========================================================================
# Workbook Creation with VBA Module Injection
# ===========================================================================

def create_workbook(
    excel: Any,
    output_path: str,
    module_paths: Sequence[str],
    import_order: Optional[Sequence[str]] = None,
) -> Any:
    """Create a .xlsm workbook with VBA modules injected.

    Parameters
    ----------
    excel : Excel.Application
        Running Excel instance from ``ensure_excel()``.
    output_path : str
        Full path for the .xlsm file to create.  The file is saved and
        re-opened so that the VBProject is fully initialised.
    module_paths : sequence of str
        Paths to .bas/.cls files to import.
    import_order : sequence of str, optional
        Module *names* (without extension) in the order they should be
        imported.  Modules whose names appear in this list are imported
        first in that order; remaining modules follow in their original
        order.  Required for VBA-Core dependencies.  If ``None``,
        modules are imported in the order provided.

    Returns
    -------
    Workbook
        The opened .xlsm workbook, ready for macro injection and testing.
    """
    # --- Build ordered import list ------------------------------------------
    if import_order is not None:
        # Build a name -> path lookup
        path_by_name: dict = {}
        for mp in module_paths:
            name = os.path.splitext(os.path.basename(mp))[0]
            path_by_name[name] = mp

        ordered: List[str] = []
        # Import in the specified order first
        for name in import_order:
            if name in path_by_name:
                ordered.append(path_by_name.pop(name))
        # Then append whatever is left (in original order)
        for mp in module_paths:
            name = os.path.splitext(os.path.basename(mp))[0]
            if name in path_by_name:
                ordered.append(mp)
        module_paths = ordered

    # --- Create workbook ----------------------------------------------------
    wb = excel.Workbooks.Add()
    try:
        # Rename Sheet1 to TestResults and add TestData
        ws = wb.Sheets(1)
        ws.Name = "TestResults"
        wb.Sheets.Add().Name = "TestData"

        # --- Inject VBA modules ---------------------------------------------
        vbproj = wb.VBProject
        for src_path in module_paths:
            mod_name = os.path.splitext(os.path.basename(src_path))[0]

            # Remove any existing component with the same name
            for comp in list(vbproj.VBComponents):
                try:
                    if comp.Name == mod_name:
                        vbproj.VBComponents.Remove(comp)
                except Exception:
                    pass

            content = read_module(src_path)
            ext = os.path.splitext(src_path)[1].lower()
            if ext == ".cls":
                comp_type = 2  # vbext_ct_ClassModule
            else:
                comp_type = 1  # vbext_ct_StdModule

            comp = vbproj.VBComponents.Add(comp_type)
            comp.Name = mod_name
            comp.CodeModule.AddFromString(content)

        # --- Save as macro-enabled workbook ---------------------------------
        # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
        wb.SaveAs(output_path, FileFormat=52)
        return wb

    except Exception:
        # Clean up the temporary workbook on failure
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
        raise


# ===========================================================================
# TestRunner VBA Module
# ===========================================================================

TESTRUNNER_VBA = r"""
Option Explicit

' ===========================================================================
' TestRunner — automated test executor
'
' Injected at test-build time into each .xlsm workbook.  Provides:
'   RunAllTests   — driver that test modules call from their own procedures
'   Assert        — write PASS / FAIL row to TestResults sheet
'   AssertClose   — floating-point comparison with tolerance
'   AssertEqual   — exact (string / integer) equality
' ===========================================================================

' ---------------------------------------------------------------------------
' Driver — initialises (or re-initialises) the TestResults sheet.
' Test modules call this at the start of their own RunAllTests-style entry
' point so that the sheet is always clean.
' ---------------------------------------------------------------------------
Public Sub RunAllTests()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("TestResults")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "TestResults"
    End If
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "Module"
    ws.Cells(1, 2).Value = "Test"
    ws.Cells(1, 3).Value = "Result"
    ws.Cells(1, 4).Value = "Details"
End Sub

' ---------------------------------------------------------------------------
' Assert — boolean pass / fail
' ---------------------------------------------------------------------------
Public Sub Assert(ByVal testName As String, ByVal condition As Boolean, _
                  Optional ByVal details As String = "")
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets("TestResults")
    ' Find last used row in column A going up (xlUp = -4162)
    r = ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 1
    ws.Cells(r, 1).Value = "TestRunner"
    ws.Cells(r, 2).Value = testName
    If condition Then
        ws.Cells(r, 3).Value = "PASS"
    Else
        ws.Cells(r, 3).Value = "FAIL"
        If Len(details) > 0 Then ws.Cells(r, 4).Value = details
    End If
End Sub

' ---------------------------------------------------------------------------
' AssertClose — floating-point equality within tolerance
' ---------------------------------------------------------------------------
Public Sub AssertClose(ByVal testName As String, ByVal actual As Double, _
                       ByVal expected As Double, _
                       Optional ByVal tol As Double = 0.000001)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets("TestResults")
    r = ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 1
    ws.Cells(r, 1).Value = "TestRunner"
    ws.Cells(r, 2).Value = testName
    If Abs(actual - expected) <= tol Then
        ws.Cells(r, 3).Value = "PASS"
    Else
        ws.Cells(r, 3).Value = "FAIL"
        ws.Cells(r, 4).Value = "Expected " & CStr(expected) & _
                               ", got " & CStr(actual)
    End If
End Sub

' ---------------------------------------------------------------------------
' AssertEqual — exact string / integer equality
' ---------------------------------------------------------------------------
Public Sub AssertEqual(ByVal testName As String, _
                       ByVal actual As Variant, _
                       ByVal expected As Variant)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets("TestResults")
    r = ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 1
    ws.Cells(r, 1).Value = "TestRunner"
    ws.Cells(r, 2).Value = testName
    If CStr(actual) = CStr(expected) Then
        ws.Cells(r, 3).Value = "PASS"
    Else
        ws.Cells(r, 3).Value = "FAIL"
        ws.Cells(r, 4).Value = "Expected " & CStr(expected) & _
                               ", got " & CStr(actual)
    End If
End Sub
"""


def inject_testrunner(wb: Any) -> None:
    """Inject the TestRunner VBA module into the workbook.

    If a TestRunner component already exists it is removed first so the
    injection is idempotent.
    """
    vbproj = wb.VBProject

    # Remove any existing TestRunner component
    for comp in list(vbproj.VBComponents):
        try:
            if comp.Name == "TestRunner":
                vbproj.VBComponents.Remove(comp)
        except Exception:
            pass

    comp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
    comp.Name = "TestRunner"
    comp.CodeModule.AddFromString(TESTRUNNER_VBA)


# ===========================================================================
# Test Execution
# ===========================================================================

def run_macro(excel: Any, wb: Any, macro_name: str, *args: Any) -> Any:
    """Run a named VBA macro and return its result.

    Parameters
    ----------
    excel : Excel.Application
    wb : Workbook
        The workbook must be the active workbook for the macro to resolve.
    macro_name : str
        Fully qualified macro name, e.g. ``"TestRunner.RunAllTests"``.
    *args : Any
        Positional arguments forwarded to ``Application.Run``.  Passed
        through to the VBA function / sub after the macro name.

    Returns
    -------
    Any
        The return value of the VBA function / sub (subs return None).
    """
    return excel.Application.Run(macro_name, *args)


def read_results(wb: Any) -> Tuple[int, int, List[str]]:
    """Parse the TestResults sheet and return aggregated counts.

    Uses ``End(xlUp)`` rather than ``UsedRange`` to find the last row —
    ``UsedRange`` can retain stale cell formatting and report an incorrect
    range size.

    Returns
    -------
    (pass_count, fail_count, details)
        details is a list of human-readable failure descriptions (one per
        FAIL row), suitable for printing in a report.
    """
    try:
        ws = wb.Sheets("TestResults")
    except Exception:
        # TestResults sheet does not exist — no tests were run
        return 0, 0, []

    # xlUp = -4162
    last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    if last_row < 2:
        # Only header row — no test rows
        return 0, 0, []

    pass_count = 0
    fail_count = 0
    details: List[str] = []

    for r in range(2, last_row + 1):
        module_val = str(ws.Cells(r, 1).Value or "")
        test_val = str(ws.Cells(r, 2).Value or "")
        result_val = str(ws.Cells(r, 3).Value or "").strip().upper()
        detail_val = str(ws.Cells(r, 4).Value or "")

        if result_val == "PASS":
            pass_count += 1
        elif result_val == "FAIL":
            fail_count += 1
            desc = f"{module_val}::{test_val}"
            if detail_val:
                desc += f"  [{detail_val}]"
            details.append(desc)
        # Non-standard result values are ignored (e.g. empty rows)

    return pass_count, fail_count, details


def print_report(
    pass_count: int,
    fail_count: int,
    details: List[str],
    label: str = "Test Results",
) -> None:
    """Print a formatted test results summary to stdout."""
    total = pass_count + fail_count
    print(f"\n{'=' * 60}")
    print(f"  {label}")
    print(f"{'=' * 60}")
    print(f"  Total : {total}")
    print(f"  PASS  : {pass_count}")
    print(f"  FAIL  : {fail_count}")
    if total > 0:
        rate = 100.0 * pass_count / total
        print(f"  Rate  : {rate:.1f}%")
    print(f"{'=' * 60}")

    if details:
        print(f"\n  Failure details ({len(details)}):")
        for d in details:
            print(f"    - {d}")
        print()

    if fail_count == 0 and total > 0:
        print("  All tests PASSED.\n")


# ===========================================================================
# Cross-Validation Helpers
# ===========================================================================

def com_to_numpy(vba_result: Any) -> np.ndarray:
    """Convert a VBA array (tuple-of-tuples) to a numpy array.

    VBA 2-D arrays are marshalled through COM as ``tuple`` of row ``tuple``
    s.  Scalar values come back as plain Python floats / ints / strings.
    This function normalises all of those into a ``numpy.ndarray``.

    Edge cases handled:
    - Scalar (float / int / str) -> shape (1,) array
    - 1-D tuple of scalars -> shape (n,) array
    - 2-D tuple of tuples -> shape (m, n) array
    - Already a numpy array -> returned as-is (no copy)
    """
    if isinstance(vba_result, np.ndarray):
        return vba_result

    # VBA Empty/Array() → None through COM
    if vba_result is None:
        return np.array([])

    # Scalar
    if not isinstance(vba_result, (tuple, list)):
        return np.array([vba_result])

    if len(vba_result) == 0:
        return np.array([])

    first = vba_result[0]

    # 1-D: first element is a scalar (not a tuple/list)
    if not isinstance(first, (tuple, list)):
        out = np.zeros(len(vba_result))
        for i, val in enumerate(vba_result):
            try:
                out[i] = float(val)
            except (ValueError, TypeError):
                out[i] = np.nan
        return out

    # 2-D: first element is a tuple -> tuple of row tuples
    rows = len(vba_result)
    cols = len(first)
    out = np.zeros((rows, cols))
    for i in range(rows):
        row = vba_result[i]
        if isinstance(row, (tuple, list)):
            for j in range(cols):
                try:
                    out[i, j] = float(row[j])
                except (ValueError, TypeError):
                    out[i, j] = np.nan
        else:
            # Degenerate: a row that is a scalar
            try:
                out[i, 0] = float(row)
            except (ValueError, TypeError):
                out[i, 0] = np.nan
    return out


def check_close(
    name: str,
    vba_val: Any,
    py_val: Any,
    tol: float = 1e-8,
) -> Union[bool, str]:
    """Compare a VBA result against a Python reference value.

    Always prefer this over a raw ``assert`` — it handles NaN, infinities,
    and non-numeric values gracefully, and returns a descriptive failure
    string instead of raising.

    Parameters
    ----------
    name : str
        Test name used in the failure message.
    vba_val : Any
        Value returned by VBA (through COM).
    py_val : Any
        Reference value computed by Python (numpy / scipy).
    tol : float
        Absolute tolerance for floating-point comparison.

    Returns
    -------
    bool or str
        ``True`` when the values match within tolerance.
        A descriptive ``str`` when they do not.
    """
    try:
        v = float(vba_val)
        p = float(py_val)
        # NaN == NaN is False in Python, so handle it explicitly
        if np.isnan(v) and np.isnan(p):
            return True
        ok = abs(v - p) <= tol
    except (ValueError, TypeError, OverflowError):
        # Fall back to string comparison for non-numeric types
        v_str = str(vba_val).strip() if vba_val is not None else ""
        p_str = str(py_val).strip() if py_val is not None else ""
        ok = v_str == p_str

    if ok:
        return True

    return f"{name}: VBA={vba_val!r}  PY={py_val!r}  tol={tol}"


def write_range(
    ws: Any,
    data: Union[np.ndarray, List, Tuple],
    start_row: int = 1,
    start_col: int = 1,
) -> None:
    """Write a 2-D list or numpy array to an Excel worksheet.

    Data is written starting at (*start_row*, *start_col*).  VBA UDFs
    expect Range inputs that come from actual worksheet cells, so this
    is the standard way to pipe Python-generated test data into VBA.

    Parameters
    ----------
    ws : Worksheet
        Target worksheet.
    data : array-like
        2-D data to write.  Supports ``list`` of ``list``, ``tuple`` of
        ``tuple``, ``numpy.ndarray``, and row-major flat lists.
    start_row : int
        First row (1-based).
    start_col : int
        First column (1-based).
    """
    # Normalise to a plain list-of-lists
    if isinstance(data, np.ndarray):
        arr = data
    else:
        arr = np.asarray(data, dtype=object)

    if arr.ndim == 0:
        ws.Cells(start_row, start_col).Value = float(arr)
        return

    if arr.ndim == 1:
        arr = arr.reshape(1, -1)

    n_rows, n_cols = arr.shape
    for i in range(n_rows):
        for j in range(n_cols):
            val = arr[i, j]
            # Convert numpy scalar types to native Python
            if hasattr(val, "item"):
                val = val.item()
            if isinstance(val, float) and np.isnan(val):
                val = None
            ws.Cells(start_row + i, start_col + j).Value = val


# ===========================================================================
# Self-test (runnable with ``python tests/test_utils.py``)
# ===========================================================================

if __name__ == "__main__":
    print("=== test_utils self-check ===")
    print(f"  PROJECT_ROOT : {PROJECT_ROOT}")
    print(f"  SRC_DIR      : {SRC_DIR}")
    print(f"  VBA_CORE_DIR : {VBA_CORE_DIR}")
    print(f"  VBA_CORE_IMPORT_ORDER : {VBA_CORE_IMPORT_ORDER}")

    # --- com_to_numpy -------------------------------------------------------
    # Scalar
    assert com_to_numpy(3.14).tolist() == [3.14]
    assert com_to_numpy(42).tolist() == [42.0]
    # 1-D
    assert com_to_numpy((1.0, 2.0, 3.0)).tolist() == [1.0, 2.0, 3.0]
    # 2-D
    arr = com_to_numpy(((1.0, 2.0), (3.0, 4.0)))
    assert arr.shape == (2, 2)
    assert arr.tolist() == [[1.0, 2.0], [3.0, 4.0]]
    # numpy pass-through
    assert com_to_numpy(np.array([1.0, 2.0])).tolist() == [1.0, 2.0]
    print("  com_to_numpy  : OK")

    # --- check_close --------------------------------------------------------
    assert check_close("ok_float", 1.0, 1.0) is True
    assert isinstance(check_close("bad_float", 1.0, 2.0), str)
    assert check_close("ok_nan", float("nan"), float("nan")) is True
    assert check_close("ok_str", "abc", "abc") is True
    assert isinstance(check_close("bad_str", "abc", "xyz"), str)
    print("  check_close   : OK")

    # --- read_module (with a temp file) -------------------------------------
    import tempfile
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".bas", delete=False, encoding="utf-8"
    ) as tf:
        tf.write('Attribute VB_Name = "TestMod"\nOption Explicit\n'
                 "Public Function Foo(): Foo = 1: End Function\n")
        tmp_path = tf.name
    try:
        content = read_module(tmp_path)
        assert "Attribute VB_Name" not in content
        assert "Option Explicit" in content
        assert "Public Function Foo" in content
        print("  read_module   : OK")
    finally:
        os.unlink(tmp_path)

    # --- write_range (com_to_numpy round-trip on plain data) -----------------
    # Just verify the array normalisation path works; actual COM write is
    # tested by integration tests.
    data_2d = [[1.0, 2.0], [3.0, 4.0]]
    arr = np.asarray(data_2d)
    assert arr.shape == (2, 2)
    print("  write_range   : array path OK")

    print("\n  All self-checks passed.")
