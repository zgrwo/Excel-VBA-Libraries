"""Cross-validate DictSetUtils functions against Python reference implementations.

Usage: python tests/build_DictSetUtils.py
"""

import os
import sys
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests.crossval.build_common import CrossValRunner
from tests.test_utils import SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER

MODULE_PATHS = [os.path.join(VBA_CORE_DIR, name + ".cls")
                for name in VBA_CORE_IMPORT_ORDER]
MODULE_PATHS.append(os.path.join(SRC_DIR, "DictSetUtils.bas"))


# =============================================================================
# Python reference helpers for set operations
# =============================================================================

def _py_set_union(a, b):
    result = []
    seen = set()
    for x in list(a) + list(b):
        key = str(x)
        if key not in seen:
            seen.add(key)
            result.append(x)
    return result


def _py_set_intersect(a, b):
    result = []
    b_set = {str(x) for x in b}
    seen = set()
    for x in a:
        key = str(x)
        if key in b_set and key not in seen:
            seen.add(key)
            result.append(x)
    return result


def _py_set_difference(a, b):
    b_set = {str(x) for x in b}
    seen = set()
    result = []
    for x in a:
        key = str(x)
        if key not in b_set and key not in seen:
            seen.add(key)
            result.append(x)
    return result


def _py_set_sym_diff(a, b):
    return _py_set_union(_py_set_difference(a, b), _py_set_difference(b, a))


# =============================================================================
# Test Cases
# =============================================================================

TEST_CASES = [

    # ---- SetUnion ----
    {"name": "SetUnion_basic", "func": "SetUnion",
     "args": lambda: ([1, 2, 3], [3, 4, 5]),
     "py_ref": lambda a: sorted(_py_set_union(a[0], a[1])),
     "result_type": "array"},
    {"name": "SetUnion_strings", "func": "SetUnion",
     "args": lambda: (["A", "B"], ["B", "C"]),
     "py_ref": lambda a: sorted(_py_set_union(a[0], a[1])),
     "result_type": "array"},
    {"name": "SetUnion_empty_first", "func": "SetUnion",
     "args": lambda: ([], [1, 2]),
     "py_ref": lambda a: sorted(_py_set_union(a[0], a[1])),
     "result_type": "array"},

    # ---- SetIntersect ----
    {"name": "SetIntersect_basic", "func": "SetIntersect",
     "args": lambda: ([1, 2, 3], [2, 3, 4]),
     "py_ref": lambda a: sorted(_py_set_intersect(a[0], a[1])),
     "result_type": "array"},
    {"name": "SetIntersect_no_overlap", "func": "SetIntersect",
     "args": lambda: ([1, 2], [3, 4]),
     "py_ref": lambda a: [],
     "result_type": "array"},
    {"name": "SetIntersect_strings", "func": "SetIntersect",
     "args": lambda: (["apple", "banana"], ["banana", "cherry"]),
     "py_ref": lambda a: ["banana"],
     "result_type": "array"},

    # ---- SetDifference ----
    {"name": "SetDifference_basic", "func": "SetDifference",
     "args": lambda: ([1, 2, 3], [2, 4]),
     "py_ref": lambda a: _py_set_difference(a[0], a[1]),
     "result_type": "array"},
    {"name": "SetDifference_empty_result", "func": "SetDifference",
     "args": lambda: ([1, 2], [1, 2, 3]),
     "py_ref": lambda a: [],
     "result_type": "array"},

    # ---- SetSymDifference ----
    {"name": "SetSymDifference_basic", "func": "SetSymDifference",
     "args": lambda: ([1, 2, 3], [2, 3, 4]),
     "py_ref": lambda a: _py_set_sym_diff(a[0], a[1]),
     "result_type": "array"},

    # ---- SetIsSubset ----
    {"name": "SetIsSubset_true", "func": "SetIsSubset",
     "args": lambda: ([1, 2], [1, 2, 3]),
     "py_ref": lambda a: True, "result_type": "bool"},
    {"name": "SetIsSubset_false", "func": "SetIsSubset",
     "args": lambda: ([1, 4], [1, 2, 3]),
     "py_ref": lambda a: False, "result_type": "bool"},
    {"name": "SetIsSubset_empty", "func": "SetIsSubset",
     "args": lambda: ([], [1, 2]),
     "py_ref": lambda a: True, "result_type": "bool"},

    # ---- SetEqual ----
    {"name": "SetEqual_true_diff_order", "func": "SetEqual",
     "args": lambda: ([1, 2, 3], [3, 2, 1]),
     "py_ref": lambda a: True, "result_type": "bool"},
    {"name": "SetEqual_false", "func": "SetEqual",
     "args": lambda: ([1, 2], [1, 2, 3]),
     "py_ref": lambda a: False, "result_type": "bool"},

    # ---- GroupCount ----
    {"name": "GroupCount_basic", "func": "GroupCount",
     "args": lambda: (["A", "B", "A", "C", "B", "A"],),
     "py_ref": lambda a: [["A", 3], ["B", 2], ["C", 1]],
     "result_type": "array"},
    {"name": "GroupCount_single_group", "func": "GroupCount",
     "args": lambda: (["X", "X", "X"],),
     "py_ref": lambda a: [["X", 3]],
     "result_type": "array"},

    # ---- SetCartesianProduct ----
    {"name": "SetCartesianProduct_basic", "func": "SetCartesianProduct",
     "args": lambda: (["A", "B"], ["1", "2"]),
     "py_ref": lambda a: [[x, y] for x in a[0] for y in a[1]],
     "result_type": "array"},
    # GroupCount edge
    {"name": "GroupCount_empty", "func": "GroupCount",
     "args": lambda: ([],),
     "py_ref": lambda a: [], "result_type": "array"},

    # DictKeys / DictValues / DictTo2DArray / DictCount
    # (skip — require live Dictionary object, not constructable through COM)
    {"name": "DictKeys_basic", "func": "DictKeys",
     "args": lambda: ("a", 1,),
     "py_ref": lambda a: None,
     "skip_if": True, "skip_reason": "DictKeys requires a Dictionary object — not constructable through COM. verified manually in Excel."},
    {"name": "DictTo2DArray_basic", "func": "DictTo2DArray",
     "args": lambda: ("x", 10,),
     "py_ref": lambda a: None,
     "skip_if": True, "skip_reason": "DictTo2DArray requires a Dictionary object. verified manually in Excel."},
    {"name": "DictCount_basic", "func": "DictCount",
     "args": lambda: ("a", 1,),
     "py_ref": lambda a: None,
     "skip_if": True, "skip_reason": "DictCount requires a Dictionary object. verified manually in Excel."},

    # CountFrequency (skip — returns Dict, not comparable through COM)
    {"name": "CountFrequency_basic", "func": "CountFrequency",
     "args": lambda: ([1, 2, 1, 3, 2, 1],),
     "py_ref": lambda a: None, "result_type": "scalar",
     "skip_if": True, "skip_reason": "CountFrequency returns Dictionary — not comparable through COM. verified manually in Excel."},

    # UDF wrappers
    {"name": "UDF_DICT_INTERSECT", "func": "UDF_DICT_INTERSECT",
     "args": lambda: ([1, 2, 3], [2, 3, 4]),
     "py_ref": lambda a: sorted(set(a[0]) & set(a[1])), "result_type": "array"},
    {"name": "UDF_DICT_DIFFERENCE", "func": "UDF_DICT_DIFFERENCE",
     "args": lambda: ([1, 2, 3], [2, 3, 4]),
     "py_ref": lambda a: sorted(set(a[0]) - set(a[1])), "result_type": "array"},
    {"name": "UDF_DICT_SYM_DIFF", "func": "UDF_DICT_SYM_DIFF",
     "args": lambda: ([1, 2, 3], [2, 3, 4]),
     "py_ref": lambda a: sorted(set(a[0]) ^ set(a[1])), "result_type": "array"},
    {"name": "UDF_DICT_CARTESIAN", "func": "UDF_DICT_CARTESIAN",
     "args": lambda: ([1, 2], [3, 4]),
     "py_ref": lambda a: [[1, 3], [1, 4], [2, 3], [2, 4]],
     "result_type": "array", "tol": 1e-10},
    {"name": "UDF_DICT_ISSUBSET", "func": "UDF_DICT_ISSUBSET",
     "args": lambda: ([1, 2], [1, 2, 3]),
     "py_ref": lambda a: True, "result_type": "bool"},
    {"name": "UDF_DICT_ISEQUAL", "func": "UDF_DICT_ISEQUAL",
     "args": lambda: ([1, 2], [2, 1]),
     "py_ref": lambda a: True, "result_type": "bool"},
    # NOTE: VBA GroupCount uses DictProxy.DictKey() internally, which
    # converts numeric keys to strings. py_ref uses string keys ("1", "2")
    # to match this behavior. If DictKey implementation changes, update.
    {"name": "UDF_DICT_GROUPCOUNT", "func": "UDF_DICT_GROUPCOUNT",
     "args": lambda: ([1, 2, 1],),
     "py_ref": lambda a: [["1", 2], ["2", 1]], "result_type": "array"},

    # UDF wrapper
    {"name": "UDF_DICT_UNION", "func": "UDF_DICT_UNION",
     "args": lambda: ([1, 2, 3], [3, 4, 5]),
     "py_ref": lambda a: sorted(set(a[0]) | set(a[1])),
     "result_type": "array"},

    # =========================================================================
    # Test extraction from VBA Test_DictSetUtils (2026-06-16)
    # =========================================================================

    # ---- SetUnion — scalar + array ----
    {"name": "SetUnion_scalar_array",
     "func": "SetUnion",
     "args": lambda: (42, [1, 2]),
     "py_ref": lambda a: sorted(_py_set_union([a[0]], a[1])),
     "result_type": "array",
     "skip_if": True,
     "skip_reason": "scalar + array COM marshaling differs from VBA Variant semantics"},

    # ---- SetUnion — both empty ----
    {"name": "SetUnion_both_empty",
     "func": "SetUnion",
     "args": lambda: ([], []),
     "py_ref": lambda a: [],
     "result_type": "array"},

    # ---- SetIntersect — both empty ----
    {"name": "SetIntersect_both_empty",
     "func": "SetIntersect",
     "args": lambda: ([], []),
     "py_ref": lambda a: [],
     "result_type": "array"},

    # ---- SetDifference — empty - nonempty ----
    {"name": "SetDifference_empty_nonempty",
     "func": "SetDifference",
     "args": lambda: ([], [1, 2]),
     "py_ref": lambda a: [],
     "result_type": "array"},

    # ---- SetSymDifference — both empty ----
    {"name": "SetSymDifference_both_empty",
     "func": "SetSymDifference",
     "args": lambda: ([], []),
     "py_ref": lambda a: [],
     "result_type": "array"},

    # ---- SetIsSubset — both empty ----
    {"name": "SetIsSubset_both_empty",
     "func": "SetIsSubset",
     "args": lambda: ([], []),
     "py_ref": lambda a: True,
     "result_type": "bool"},

    # ---- SetEqual — both empty ----
    {"name": "SetEqual_both_empty",
     "func": "SetEqual",
     "args": lambda: ([], []),
     "py_ref": lambda a: True,
     "result_type": "bool"},

    # ---- SetEqual — empty vs nonempty ----
    {"name": "SetEqual_empty_vs_nonempty",
     "func": "SetEqual",
     "args": lambda: ([], [1]),
     "py_ref": lambda a: False,
     "result_type": "bool"},

    # ---- SetCartesianProduct — one empty ----
    {"name": "SetCartesianProduct_empty_first",
     "func": "SetCartesianProduct",
     "args": lambda: ([], [1, 2]),
     "py_ref": lambda a: [],
     "result_type": "array"},

    # ---- SetCartesianProduct — other empty ----
    {"name": "SetCartesianProduct_empty_second",
     "func": "SetCartesianProduct",
     "args": lambda: ([1, 2], []),
     "py_ref": lambda a: [],
     "result_type": "array"},

    # ---- Boundary: duplicates / equal-sets / single-elements ----
    {"name": "SetUnion_all_duplicates", "func": "SetUnion",
     "args": lambda: ([1, 1, 1], [1, 1, 1]),
     "py_ref": lambda a: [1], "result_type": "array"},
    {"name": "SetIntersect_all_duplicates", "func": "SetIntersect",
     "args": lambda: ([1, 1, 1], [1, 2, 1]),
     "py_ref": lambda a: [1], "result_type": "array"},
    {"name": "SetIsSubset_equal_sets", "func": "SetIsSubset",
     "args": lambda: ([1, 2, 3], [1, 2, 3]),
     "py_ref": lambda a: True, "result_type": "bool"},
    {"name": "SetEqual_different_length", "func": "SetEqual",
     "args": lambda: ([1, 2], [1, 2, 3]),
     "py_ref": lambda a: False, "result_type": "bool"},
    {"name": "SetCartesianProduct_single_each", "func": "SetCartesianProduct",
     "args": lambda: (["A"], ["B"]),
     "py_ref": lambda a: [["A", "B"]], "result_type": "array"},
]


def main() -> int:
    runner = CrossValRunner("DictSetUtils", MODULE_PATHS)
    runner.run_all(TEST_CASES)
    passed, failed = runner.print_summary()
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
