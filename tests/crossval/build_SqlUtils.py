"""Cross-validate SqlUtils functions against Python references.

Usage: python tests/crossval/build_SqlUtils.py

SqlUtils uses ADODB to query the saved workbook. Test data is written to
the TestData sheet before SQL functions are called.
"""

import os, sys
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests.crossval.build_common import CrossValRunner
from tests.test_utils import SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER

MODULE_PATHS = [os.path.join(VBA_CORE_DIR, n + ".cls") for n in VBA_CORE_IMPORT_ORDER]
MODULE_PATHS.append(os.path.join(SRC_DIR, "SqlUtils.bas"))

# Test data: simple 3-column table for SQL queries
# Columns: ID, Name, Score
SQL_TEST_DATA = [
    [1, "Alice", 85.5],
    [2, "Bob", 92.0],
    [3, "Charlie", 78.3],
    [4, "Diana", 88.7],
    [5, "Eve", 95.1],
    [6, "Frank", 100.0],  # 触发字符串比较 bug: "100" < "90" (#45)
]

SQL_COLUMNS = ["ID", "Name", "Score"]

# R5-05/R5-36 regression fixtures: WHERE keyword must be located outside string
# literals / bracket identifiers; ORDER BY inside a WHERE literal is legal.
R5_WHERE_DATA = [
    ["ID", "Name"],
    [1, "A"],
    [2, "B"],
    [3, "C"],
    [4, "D"],
    [5, "E"],
]

R5_ORDERBY_DATA = [
    ["ID", "Name"],
    [1, "ORDER BY x"],
    [2, "B"],
]

# mode -> (fixture, expected data-row count after filtering)
R5_WHERE_PROBE_CASES = {
    "baseline": (R5_WHERE_DATA, 2.0),
    "literal_where": (R5_WHERE_DATA, 2.0),
    "bracket_table": (R5_WHERE_DATA, 2.0),
    "where_paren": (R5_WHERE_DATA, 2.0),
    "where_tab": (R5_WHERE_DATA, 2.0),
    "order_by_literal": (R5_ORDERBY_DATA, 1.0),
}

# R5-27: SqlListColumns table-name normalization (plain / quoted / bracketed).
R5_LISTCOL_CASES = [
    ("plain_TestData", "TestData", ["ID", "Name", "Score"]),
    ("quoted_TestData", "'TestData$'", ["ID", "Name", "Score"]),
    ("bracketed_TestData", "[TestData$]", ["ID", "Name", "Score"]),
    ("space_plain", "Data Sheet", ["Col A", "Col B"]),
    ("space_quoted", "'Data Sheet$'", ["Col A", "Col B"]),
    ("space_bracketed", "[Data Sheet$]", ["Col A", "Col B"]),
]

PROBE_MODULE_NAME = "SqlR5Probe"
PROBE_LINES = [
    "Option Explicit",
    "",
    "Public Function ProbeWhere(ByVal rng As Range, ByVal mode As String) As Double",
    "    On Error GoTo EH",
    "    Dim sqlText As String, res As Variant, ok As Boolean",
    "    Select Case mode",
    "        Case \"baseline\"",
    "            sqlText = \"SELECT * FROM data WHERE ID > 3\"",
    "        Case \"literal_where\"",
    "            sqlText = \"SELECT 'WHERE' AS w FROM data WHERE ID > 3\"",
    "        Case \"bracket_table\"",
    "            sqlText = \"SELECT * FROM [DataWhere$] WHERE ID > 3\"",
    "        Case \"where_paren\"",
    "            sqlText = \"SELECT * FROM data WHERE(ID > 3)\"",
    "        Case \"where_tab\"",
    "            sqlText = \"SELECT * FROM data\" & vbTab & \"WHERE\" & vbTab & \"ID > 3\"",
    "        Case \"order_by_literal\"",
    "            sqlText = \"SELECT * FROM data WHERE Name = 'ORDER BY x'\"",
    "        Case Else",
    "            ProbeWhere = -999",
    "            Exit Function",
    "    End Select",
    "    res = SqlRangeQuery(sqlText, rng, \"data\", ok)",
    "    If Not ok Then",
    "        ProbeWhere = -1",
    "        Exit Function",
    "    End If",
    "    ProbeWhere = CDbl(UBound(res, 1) - 1)",
    "    Exit Function",
    "EH:",
    "    ProbeWhere = -CDbl(Err.Number)",
    "End Function",
]


def _ensure_probe_module(wb):
    """Inject SqlR5Probe module (idempotent) for in-VBA error capture."""
    vbproj = wb.VBProject
    for comp in list(vbproj.VBComponents):
        if comp.Name == PROBE_MODULE_NAME:
            return
    comp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
    comp.Name = PROBE_MODULE_NAME
    comp.CodeModule.AddFromString("\r\n".join(PROBE_LINES))



def _py_list_sheets(args):
    """Expected sheets in the workbook (created by create_workbook)."""
    return ["TestData", "TestResults"]


def _py_list_columns(args):
    """Expected columns (sorted)."""
    return sorted(SQL_COLUMNS)


def _py_query_select_all(args):
    """Execute SELECT * FROM [TestData$] — returns all rows with header."""
    return [SQL_COLUMNS] + [list(row) for row in SQL_TEST_DATA]


def _py_query_select_filter(args):
    """SELECT Name, Score FROM [TestData$] WHERE Score > 90"""
    return [["Name", "Score"], ["Bob", 92.0], ["Eve", 95.1], ["Frank", 100.0]]


def _py_query_select_count(args):
    """SELECT COUNT(*) FROM [TestData$]"""
    return [[6]]


def _r5_where_probe(excel, wb, ws, runner, tc, args):
    """R5-05/R5-36: WHERE keyword parsing regression (errors captured in VBA).

    Returns the number of data rows (header excluded). SqlRangeQuery errors
    surface as negative Err.Number, so a bad parse/filter yields a mismatch
    instead of an exception crossing the COM boundary.
    """
    from tests.test_utils import run_macro, write_range
    mode = args[0]
    data, expected = R5_WHERE_PROBE_CASES[mode]
    probe_ws = _get_probe_sheet(wb)
    probe_ws.UsedRange.ClearContents()
    arr = np.asarray(data, dtype=object)
    write_range(probe_ws, arr, 1, 1)
    rng = probe_ws.Range(probe_ws.Cells(1, 1),
                         probe_ws.Cells(arr.shape[0], arr.shape[1]))
    val = run_macro(excel, wb, f"{PROBE_MODULE_NAME}.ProbeWhere", rng, mode)
    return float(val), float(expected), 0.0


def _get_probe_sheet(wb):
    """Dedicated fixture sheet for the range-query probe (keeps TestData intact)."""
    try:
        return wb.Sheets("SqlR5Data")
    except Exception:
        probe_ws = wb.Sheets.Add()
        probe_ws.Name = "SqlR5Data"
        return probe_ws


# R5-28 activation: previously all five cases were SKIP. These minimal cases run
# without Excel formulas (direct VBA calls / in-VBA probe) and are persisted
# regressions for R5-05, R5-26 (escape output), R5-27 and R5-36.
R5_ACTIVE_CASES = [
    # ---- SqlEscapeString outputs (plain / bracket / forLike) ----
    {
        "name": "SqlEscapeString_plain",
        "func": "SqlEscapeString",
        "args": lambda: ("O'Brien's",),
        "py_ref": lambda a: "O''Brien''s",
        "result_type": "string",
    },
    {
        # forLike=False (默认): 方括号为字面量, 仅 '' 转义; forLike 分支见下例
        "name": "SqlEscapeString_bracket_literal",
        "func": "SqlEscapeString",
        "args": lambda: ("a[b]",),
        "py_ref": lambda a: "a[b]",
        "result_type": "string",
    },
    {
        "name": "SqlEscapeString_forLike",
        "func": "SqlEscapeString",
        "args": lambda: ("100%_x[", True),
        "py_ref": lambda a: "100[%][_]x[[]",
        "result_type": "string",
    },
]

# ---- SqlListColumns: plain / quoted / bracketed / spaced table names (R5-27) ----
R5_ACTIVE_CASES += [
    {
        "name": f"SqlListColumns_{name}",
        "func": "SqlListColumns",
        "args": lambda a=arg: (a,),
        "py_ref": lambda a, cols=cols: cols,
        "compare_mode": "contains",
    }
    for name, arg, cols in R5_LISTCOL_CASES
]

# ---- SqlRangeQuery WHERE keyword boundaries (R5-05) / ORDER BY literal (R5-36) ----
R5_ACTIVE_CASES += [
    {
        "name": f"SqlRangeQuery_where_{mode}",
        "func": "SqlRangeQuery",
        "args": lambda m=mode: (m,),
        "reconstruct": _r5_where_probe,
        "py_ref": lambda a, m=mode: R5_WHERE_PROBE_CASES[m][1],
    }
    for mode in R5_WHERE_PROBE_CASES
]


TEST_CASES = [
    # =========================================================================
    # SqlListSheets — verify workbook contains expected sheets
    # =========================================================================
    {
        "name": "SqlListSheets_contains",
        "func": "SqlListSheets",
        "args": lambda: (),
        "py_ref": _py_list_sheets,
        "result_type": "array",
        "compare_mode": "contains",
        "skip_if": True,
        "skip_reason": "ADODB sheet naming varies ($ suffix, quotes); VBA Test_SqlUtils covers this",
    },

    # =========================================================================
    # SQL query tests — ADODB adds system columns and type marshaling varies.
    # Full SQL logic verified by VBA Test_SqlUtils (30 assertions).
    # =========================================================================
    {
        "name": "SqlListColumns_TestData",
        "func": "SqlListColumns",
        "args": lambda: ("[TestData$]",),
        "py_ref": _py_list_columns,
        "result_type": "array",
        "compare_mode": "sorted",
        "skip_if": True,
        "skip_reason": "ADODB returns extra system columns (ColIndex etc.); VBA Test_SqlUtils covers this",
    },
    {
        "name": "SqlExecute_select_all",
        "func": "SqlExecute",
        "args": lambda: ("SELECT * FROM [TestData$]",),
        "py_ref": _py_query_select_all,
        "result_type": "array",
        "tol": 1e-10,
        "skip_if": True,
        "skip_reason": "COM int→float type marshaling mismatches py_ref; VBA Test_SqlUtils covers this",
    },
    {
        "name": "SqlQuery_filter_score",
        "func": "SqlQuery",
        "args": lambda: ("Name, Score", "[TestData$]", "Score > 90"),
        "py_ref": _py_query_select_filter,
        "result_type": "array",
        "tol": 1e-10,
        "skip_if": True,
        "skip_reason": "ADODB type marshaling; VBA Test_SqlUtils covers this",
    },
    {
        "name": "SqlExecute_count",
        "func": "SqlExecute",
        "args": lambda: ("SELECT COUNT(*) AS cnt FROM [TestData$]",),
        "py_ref": _py_query_select_count,
        "result_type": "array",
        "tol": 1e-10,
        "skip_if": True,
        "skip_reason": "COUNT result structure varies through COM; VBA Test_SqlUtils covers this",
    },
] + R5_ACTIVE_CASES


def _flatten_sort(v):
    """Flatten nested tuples/lists, return sorted unique list of strings."""
    if v is None:
        return []
    if isinstance(v, (list, tuple)):
        items = []
        for x in v:
            if isinstance(x, (list, tuple)):
                items.extend(str(y).strip() for y in x)
            else:
                items.append(str(x).strip())
        return sorted(set(items))
    return sorted(set(str(v).strip().split(",")))


class SqlUtilsRunner(CrossValRunner):
    """Extended runner with support for 'contains' and 'sorted' compare modes,
    and data seeding for SQL tests."""

    def run_all(self, test_cases):
        import tempfile
        from tests.test_utils import (
            ensure_excel, teardown, create_workbook,
            inject_testrunner, write_range,
        )

        excel = ensure_excel()
        wb = None
        try:
            output = os.path.join(
                tempfile.gettempdir(),
                f"vba_crossval_{self.module_name}.xlsm",
            )
            wb = create_workbook(
                excel, output, self.module_paths,
                import_order=self._import_order,
            )
            inject_testrunner(wb)
            _ensure_probe_module(wb)

            # Write SQL test data to TestData sheet
            td = wb.Sheets("TestData")
            for j, col_name in enumerate(SQL_COLUMNS):
                td.Cells(1, j + 1).Value = col_name
            arr = np.asarray(SQL_TEST_DATA, dtype=object)
            write_range(td, arr, 2, 1)

            # R5-27 fixture: sheet name with a space (ACE OpenSchema → 'Data Sheet$')
            sd = wb.Sheets.Add()
            sd.Name = "Data Sheet"
            sd.Cells(1, 1).Value = "Col A"
            sd.Cells(1, 2).Value = "Col B"
            wb.Save()

            self.results = []
            for tc in test_cases:
                self._run_one_sql(excel, wb, td, tc)

            return self.results
        finally:
            teardown(excel, wb)

    def _run_one_sql(self, excel, wb, ws, tc):
        """Run one test case with custom comparison modes."""
        label = f"{self.module_name}.{tc['func']}.{tc['name']}"

        if tc.get("skip_if", False):
            print(f"  SKIP  {label} — {tc.get('skip_reason', 'no reason')}")
            self.results.append((self.module_name, tc["name"], "SKIP",
                                 tc.get("skip_reason", "")))
            return

        try:
            args = tc["args"]() if callable(tc["args"]) else tc.get("args", ())
            if not isinstance(args, (tuple, list)):
                args = (args,)

            # Convert datetime args to ISO strings
            args = tuple(self._to_com_arg(a) for a in args)

            if tc.get("reconstruct"):
                # In-VBA probe: multi-step call + VBA-side error capture
                vba_result, py_val, tol = tc["reconstruct"](
                    excel, wb, ws, self, tc, args)
                self._compare_scalar(label, vba_result, py_val, tol)
                return

            if tc.get("is_udf"):
                vba_result = self._call_udf(excel, wb, ws, tc, args)
            else:
                macro = f"{self.module_name}.{tc['func']}"
                vba_result = self._call_vba_raw(excel, wb, macro, *args)

            py_val = tc["py_ref"](args) if callable(tc["py_ref"]) else tc["py_ref"]
            cmp_mode = tc.get("compare_mode", "exact")

            if cmp_mode == "contains":
                self._compare_contains(label, vba_result, py_val, tc)
            elif cmp_mode == "sorted":
                self._compare_sorted(label, vba_result, py_val, tc)
            else:
                result_type = tc.get("result_type", "scalar")
                tol = tc.get("tol", 1e-10)
                self._compare(label, vba_result, py_val, result_type, tol, tc)

        except Exception as exc:
            self.results.append((self.module_name, tc["name"], "FAIL",
                                f"exception: {exc}"))
            print(f"  FAIL  {label} — exception: {exc}")

    @staticmethod
    def _call_vba_raw(excel, wb, macro, *args):
        """Call VBA macro (uses run_macro from test_utils)."""
        from tests.test_utils import run_macro
        return run_macro(excel, wb, macro, *args)

    def _compare_contains(self, label, vba_result, py_val, tc):
        """Check that VBA result contains ALL items in py_val."""
        vba_items = set(_flatten_sort(vba_result))
        py_items = set(_flatten_sort(py_val))
        missing = py_items - vba_items
        ok = len(missing) == 0
        status = "PASS" if ok else "FAIL"
        detail = "" if ok else f"missing: {sorted(missing)}"
        self.results.append((self.module_name, tc["name"], status, detail))
        if ok:
            print(f"  PASS  {label}  contains check OK")
        else:
            print(f"  FAIL  {label} — {detail}")

    def _compare_sorted(self, label, vba_result, py_val, tc):
        """Sort both sides before array comparison."""
        vba_flat = sorted(_flatten_sort(vba_result))
        py_flat = sorted(_flatten_sort(py_val))
        ok = vba_flat == py_flat
        status = "PASS" if ok else "FAIL"
        detail = "" if ok else f"VBA={vba_flat!r} PY={py_flat!r}"
        self.results.append((self.module_name, tc["name"], status, detail))
        if ok:
            print(f"  PASS  {label}  sorted compare OK")
        else:
            print(f"  FAIL  {label} — {detail}")


def main() -> int:
    runner = SqlUtilsRunner("SqlUtils", MODULE_PATHS)
    runner.run_all(TEST_CASES)
    passed, failed = runner.print_summary()
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
