"""Cross-validate PivotUtils functions against Python reference implementations.

Usage: python tests/build_PivotUtils.py
"""

import os, sys
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests.crossval.build_common import CrossValRunner
from tests.test_utils import SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER

MODULE_PATHS = [os.path.join(VBA_CORE_DIR, n + ".cls") for n in VBA_CORE_IMPORT_ORDER]
MODULE_PATHS.append(os.path.join(SRC_DIR, "RangeUtils.bas"))
MODULE_PATHS.append(os.path.join(SRC_DIR, "PivotUtils.bas"))

SALES_HDR = [["产品","月份","销量"]] + [
         ["产品A","1月",100],["产品A","2月",150],["产品A","3月",200],
         ["产品B","1月",80],["产品B","2月",120],["产品B","3月",160]]

def _py_groupby_list(data, gc, ac, fn):
    """Python ref for GroupBy (matches VBA output WITH header row)."""
    grp = {}
    for r in data[1:]:  # skip header row
        k, v = r[gc-1], float(r[ac-1])
        grp.setdefault(k, []).append(v)
    fn_u = fn.upper()
    out = [[str(data[0][gc-1]), fn_u + "(" + str(data[0][ac-1]) + ")"]]
    for k in sorted(grp, key=str):
        vs = grp[k]
        if fn_u == "SUM":    out.append([k, sum(vs)])
        elif fn_u == "AVG":  out.append([k, sum(vs)/len(vs)])
        elif fn_u == "COUNT": out.append([k, len(vs)])
        elif fn_u == "MIN":  out.append([k, min(vs)])
        elif fn_u == "MAX":  out.append([k, max(vs)])
    return out

def _py_cross(a, b):
    """CrossJoin Python ref — cartesian product as list of lists (numeric strings → float)."""
    result = []
    for x in a:
        for y in b:
            yv = float(y) if isinstance(y, str) and y.replace('.','').isdigit() else y
            result.append([x, yv])
    return result

# VL_DATA needs a header row — VLookupArray treats row 1 as header (i starts at 2)
VL_DATA = [["ID", "Name", "Score"], [1, "Alice", 85.5], [2, "Bob", 92.0], [3, "Charlie", 78.3]]


def _py_vlookup(data, lookup_val, lookup_col, return_col):
    """Python ref for VLookupArray — case-insensitive text compare on 2D data (row 0 = header)."""
    for row in data[1:]:  # skip header
        if str(row[lookup_col - 1]).lower() == str(lookup_val).lower():
            v = row[return_col - 1]
            return float(v) if isinstance(v, (int, float)) else v
    return None  # not found


def _py_raw_convert(args):
    """Python ref for RawConversion — simple pivot: colDim=product, rowDim=month.
    Input is 2D list with header [产品, 月份, 销量], valueCol=3, colDim=1, rowDim=2."""
    data = args[0]
    value_col = int(args[1]) - 1
    col_dim = int(args[2]) - 1
    row_dim = int(args[3]) - 1
    pivot = {}
    col_set = set()
    for row in data[1:]:
        rk = str(row[row_dim])
        ck = str(row[col_dim])
        v = float(row[value_col])
        col_set.add(ck)
        pivot.setdefault(rk, {})[ck] = v
    cols = sorted(col_set)
    rows = sorted(pivot.keys())
    result = [[data[0][row_dim]] + cols]
    for rk in rows:
        result.append([rk] + [pivot[rk].get(c, 0.0) for c in cols])
    return result


TEST_CASES = [
    # =========================================================================
    # GroupBy — SUM / AVG / COUNT
    # =========================================================================
    {"name": "GroupBy_SUM", "func": "GroupBy",
     "args": lambda: (SALES_HDR, 1, 3, "SUM"),
     "py_ref": lambda a: _py_groupby_list(a[0], a[1], a[2], a[3]),
     "result_type": "array", "is_udf": True},
    {"name": "GroupBy_AVG", "func": "GroupBy",
     "args": lambda: (SALES_HDR, 1, 3, "AVG"),
     "py_ref": lambda a: _py_groupby_list(a[0], a[1], a[2], a[3]),
     "result_type": "array", "is_udf": True},
    {"name": "GroupBy_COUNT", "func": "GroupBy",
     "args": lambda: (SALES_HDR, 1, 3, "COUNT"),
     "py_ref": lambda a: _py_groupby_list(a[0], a[1], a[2], a[3]),
     "result_type": "array", "is_udf": True},
    # =========================================================================
    # CrossJoin — cartesian product
    # =========================================================================
    {"name": "CrossJoin", "func": "CrossJoin",
     "args": lambda: (["X","A","B"], ["Y","1","2","3"]),
     "py_ref": lambda a: [[a[0][0], a[1][0]]] + _py_cross(a[0][1:], a[1][1:]),
     "result_type": "array", "is_udf": True},
    # =========================================================================
    # VLookupArray — in-memory VLOOKUP on 2D array (non-UDF, direct call)
    # =========================================================================
    {"name": "VLookupArray_found", "func": "VLookupArray",
     "args": lambda: (VL_DATA, "Bob", 2, 3),
     "py_ref": lambda a: _py_vlookup(a[0], a[1], a[2], a[3]),
     "result_type": "scalar", "tol": 1e-10},
    {"name": "VLookupArray_case_insensitive", "func": "VLookupArray",
     "args": lambda: (VL_DATA, "ALICE", 2, 3),
     "py_ref": lambda a: _py_vlookup(a[0], a[1], a[2], a[3]),
     "result_type": "scalar", "tol": 1e-10},
    {"name": "VLookupArray_number_lookup", "func": "VLookupArray",
     "args": lambda: (VL_DATA, 3, 1, 3),
     "py_ref": lambda a: _py_vlookup(a[0], a[1], a[2], a[3]),
     "result_type": "scalar", "tol": 1e-10},
    # =========================================================================
    # RawConversion — pivot/crosstab (UDF, Range input)
    #   NOTE: Chinese headers break through COM string comparison
    #   (encoding mismatch between Python str and VBA Variant).
    # =========================================================================
    {"name": "RawConversion_basic", "func": "RawConversion",
     "args": lambda: (SALES_HDR, 3, 1, 2),
     "py_ref": lambda a: _py_raw_convert(a),
     "result_type": "array", "is_udf": True, "tol": 1e-10,
     "skip_if": True,
     "skip_reason": "Chinese headers break COM string encoding; tested manually in Excel"},

    # UDF wrappers
    {"name": "UDF_PIVOT_VLOOKUP", "func": "UDF_PIVOT_VLOOKUP",
     "args": lambda: ([[1, "A"], [2, "B"]], 2, 1, 2),
     "py_ref": lambda a: "B", "result_type": "string", "is_udf": True},
    {"name": "UDF_PIVOT_GROUPBY", "func": "UDF_PIVOT_GROUPBY",
     "args": lambda: ([["G", "V"], ["A", 10], ["A", 20], ["B", 30]], 1, 2, "SUM"),
     "py_ref": lambda a: [["G", "SUM(V)"], ["A", 30.0], ["B", 30.0]],
     "result_type": "array", "is_udf": True, "tol": 1e-10},
    {"name": "UDF_PIVOT_CROSSJOIN", "func": "UDF_PIVOT_CROSSJOIN",
     "args": lambda: ([[1], [2]], [[3], [4]]),
     "py_ref": lambda a: [[1, 3], [1, 4], [2, 3], [2, 4]],
     "result_type": "array", "is_udf": True, "tol": 1e-10,
     "skip_if": True, "skip_reason": "UDF_PIVOT_CROSSJOIN needs live Range objects (not passable through COM Application.Run)."},

    # =====================================================================
    # Migrated from VBA Test_PivotUtils — coverage gaps (2026-06-16)
    # =====================================================================

    # ---- RawConversionFromArray — pivots; VBA appends SuffixLetter to col headers ----
    {"name": "RawConversionFromArray_basic", "func": "RawConversionFromArray",
     "args": lambda: ([["Region","Product","Sales"],["East","A",100],["West","B",200]],
                      "Sales", "Product", "Region"),
     "py_ref": lambda a: [["Region", "AA", "BA"], ["East", 100, ""], ["West", "", 200]],
     "result_type": "array", "tol": 1e-10},

    # Sub procedures (FilterTable, TransposeTable, etc.) — write to Range, no COM testable

]

def main() -> int:
    runner = CrossValRunner("PivotUtils", MODULE_PATHS, extra_imports=["RangeUtils"])
    runner.run_all(TEST_CASES)
    passed, failed = runner.print_summary()
    return 0 if failed == 0 else 1

if __name__ == "__main__":
    sys.exit(main())
