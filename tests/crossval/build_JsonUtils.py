"""Cross-validate JsonUtils functions against Python reference implementations.

Usage: python tests/build_JsonUtils.py
"""

import os
import sys
import json as _json

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests.crossval.build_common import CrossValRunner
from tests.test_utils import SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER

MODULE_PATHS = [os.path.join(VBA_CORE_DIR, name + ".cls")
                for name in VBA_CORE_IMPORT_ORDER]
MODULE_PATHS.append(os.path.join(SRC_DIR, "JsonUtils.bas"))


# =============================================================================
# Test Cases
# =============================================================================

TEST_CASES = [

    # ---- JsonParse — core parser (new) ----
    {
        "name": "JsonParse_number",
        "func": "JsonParse",
        "args": lambda: ("42",),
        "py_ref": lambda a: 42.0,
        "result_type": "scalar", "tol": 1e-10,
    },
    {
        "name": "JsonParse_string",
        "func": "JsonParse",
        "args": lambda: ('"hello"',),
        "py_ref": lambda a: "hello",
        "result_type": "string",
    },
    {
        "name": "JsonParse_array",
        "func": "JsonParse",
        "args": lambda: ("[1,2,3]",),
        "py_ref": lambda a: [1.0, 2.0, 3.0],
        "result_type": "array", "tol": 1e-10,
    },
    {
        "name": "JsonParse_bool_true",
        "func": "JsonParse",
        "args": lambda: ("true",),
        "py_ref": lambda a: True,
        "result_type": "bool",
    },
    {
        "name": "JsonParse_null",
        "func": "JsonParse",
        "args": lambda: ("null",),
        "py_ref": lambda a: None,
        "result_type": "scalar",
        "skip_if": True,
        "skip_reason": "VBA Null marshals to COM None; comparison needs special handling",
    },
    {
        "name": "JsonParse_negative",
        "func": "JsonParse",
        "args": lambda: ("-17.5",),
        "py_ref": lambda a: -17.5,
        "result_type": "scalar", "tol": 1e-10,
    },

    # ---- JsonIsValid ----
    {
        "name": "JsonIsValid_true_for_object",
        "func": "JsonIsValid",
        "args": lambda: ('{"a":1}',),
        "py_ref": lambda a: True,
        "result_type": "bool",
    },
    {
        "name": "JsonIsValid_true_for_string",
        "func": "JsonIsValid",
        "args": lambda: ('"hello"',),
        "py_ref": lambda a: True,
        "result_type": "bool",
    },
    {
        "name": "JsonIsValid_true_for_number",
        "func": "JsonIsValid",
        "args": lambda: ("42",),
        "py_ref": lambda a: True,
        "result_type": "bool",
    },
    {
        "name": "JsonIsValid_false_for_bad_json",
        "func": "JsonIsValid",
        "args": lambda: ("{bad}",),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },
    {
        "name": "JsonIsValid_false_for_empty",
        "func": "JsonIsValid",
        "args": lambda: ("",),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },
    {
        "name": "JsonIsValid_true_for_boolean",
        "func": "JsonIsValid",
        "args": lambda: ("true",),
        "py_ref": lambda a: True,
        "result_type": "bool",
    },
    {
        "name": "JsonIsValid_true_for_null",
        "func": "JsonIsValid",
        "args": lambda: ("null",),
        "py_ref": lambda a: True,
        "result_type": "bool",
    },

    # ---- JsonGet ----
    {
        "name": "JsonGet_nested_path",
        "func": "JsonGet",
        "args": lambda: ('{"users":[{"name":"Alice"},{"name":"Bob"}]}', "users[0].name"),
        "py_ref": lambda a: "Alice",
        "result_type": "string",
    },
    {
        "name": "JsonGet_simple_key",
        "func": "JsonGet",
        "args": lambda: ('{"name":"John","age":30}', "name"),
        "py_ref": lambda a: "John",
        "result_type": "string",
    },
    {
        "name": "JsonGet_array_index",
        "func": "JsonGet",
        "args": lambda: ('[10,20,30]', "[1]"),
        "py_ref": lambda a: "20",
        "result_type": "string",
    },

    # ---- JsonGetKeys ----
    {
        "name": "JsonGetKeys_two_keys",
        "func": "JsonGetKeys",
        "args": lambda: ('{"a":1,"b":2}',),
        "py_ref": lambda a: ["a", "b"],
        "result_type": "array",
    },
    # J-02 回归: RFC 8259 键大小写敏感 — "a"/"A" 必须保留为两个键
    {
        "name": "JsonGetKeys_case_sensitive",
        "func": "JsonGetKeys",
        "args": lambda: ('{"a":1,"A":2}',),
        "py_ref": lambda a: ["a", "A"],
        "result_type": "array",
    },
    {
        "name": "JsonGetKeys_empty_object",
        "func": "JsonGetKeys",
        "args": lambda: ('{}',),
        "py_ref": lambda a: [],
        "result_type": "array",
    },
    {
        "name": "JsonGetKeys_nested_keys",
        "func": "JsonGetKeys",
        "args": lambda: ('{"data":{"x":1},"meta":"info"}',),
        "py_ref": lambda a: ["data", "meta"],
        "result_type": "array",
    },

    # ---- JsonStringify ----
    {
        "name": "JsonStringify_array",
        "func": "JsonStringify",
        "args": lambda: ([1, 2, 3],),
        "py_ref": lambda a: _json.dumps(a[0], separators=(',', ':')),
        "result_type": "string",
    },
    {
        "name": "JsonStringify_empty_array",
        "func": "JsonStringify",
        "args": lambda: ([],),
        "py_ref": lambda a: "[]",
        "result_type": "string",
    },
    {
        "name": "JsonStringify_nested",
        "func": "JsonStringify",
        "args": lambda: ([[1, 2], [3, 4]],),
        "py_ref": lambda a: _json.dumps(a[0], separators=(',', ':')),
        "result_type": "string",
    },
    {
        "name": "JsonStringify_mixed_types",
        "func": "JsonStringify",
        "args": lambda: (["hello", 42, True],),
        "py_ref": lambda a: _json.dumps(a[0], separators=(',', ':')),
        "result_type": "string",
    },

    # =========================================================================
    # Unicode / non-BMP tests (C1 sync — surrogate pair handling)
    # =========================================================================

    # ---- JsonParse — emoji (U+1F600 😀, surrogate pair in VBA) ----
    # NOTE: VBA JsonParse returns arrays from COM as nested tuples;
    # _compare_array handles flattening for correct comparison.
    {"name": "JsonParse_emoji",
     "func": "JsonParse",
     "args": lambda: ('["\U0001F600"]',),
     "py_ref": lambda a: ["😀"],
     "result_type": "array"},

    # ---- JsonParse — empty array round-trip (M6 sync) ----
    {"name": "JsonParse_empty_array",
     "func": "JsonParse",
     "args": lambda: ("[]",),
     "py_ref": lambda a: [],
     "result_type": "array"},

    # ---- JsonParse — CJK Ext-B character (U+2000B 𠀋, surrogate pair) ----
    {"name": "JsonParse_cjk_extb",
     "func": "JsonParse",
     "args": lambda: ('"\U0002000B"',),
     "py_ref": lambda a: "\U0002000B",
     "result_type": "string"},

    # =========================================================================
    # Edge case tests (P2-5 additions)
    # =========================================================================

    # ---- JsonParse — string escapes ----
    {"name": "JsonParse_string_escapes",
     "func": "JsonParse",
     "args": lambda: (r'"hello\nworld"',),
     "py_ref": lambda a: "hello\nworld",
     "result_type": "string"},

    # ---- JsonParse — scientific notation ----
    {"name": "JsonParse_scientific",
     "func": "JsonParse",
     "args": lambda: ("1.5e3",),
     "py_ref": lambda a: 1500.0,
     "result_type": "scalar", "tol": 1e-10},

    # ---- JsonParse — large number ----
    {"name": "JsonParse_large_number",
     "func": "JsonParse",
     "args": lambda: ("9999999999",),
     "py_ref": lambda a: 9999999999.0,
     "result_type": "scalar", "tol": 1e-6},

    # ---- JsonParse — nested array ----
    {"name": "JsonParse_nested_array",
     "func": "JsonParse",
     "args": lambda: ("[[1,2],[3,4]]",),
     "py_ref": lambda a: [[1.0, 2.0], [3.0, 4.0]],
     "result_type": "array", "tol": 1e-10},

    # ---- JsonIsValid — array type (new) ----
    {"name": "JsonIsValid_true_for_array",
     "func": "JsonIsValid",
     "args": lambda: ("[1,2,3]",),
     "py_ref": lambda a: True,
     "result_type": "bool"},

    # ---- JsonIsValid — scientific notation ----
    {"name": "JsonIsValid_true_for_scientific",
     "func": "JsonIsValid",
     "args": lambda: ("1e10",),
     "py_ref": lambda a: True,
     "result_type": "bool"},

    # ---- JsonIsValid — negative number ----
    {"name": "JsonIsValid_true_for_negative",
     "func": "JsonIsValid",
     "args": lambda: ("-42.5",),
     "py_ref": lambda a: True,
     "result_type": "bool"},

    # =========================================================================
    # Test extraction from VBA Test_JsonUtils (2026-06-16)
    # =========================================================================

    # ---- JsonParse — bool false ----
    {"name": "JsonParse_bool_false",
     "func": "JsonParse",
     "args": lambda: ("false",),
     "py_ref": lambda a: False,
     "result_type": "bool"},

    # ---- JsonGet — dot path into nested object ----
    {"name": "JsonGet_dot_path",
     "func": "JsonGet",
     "args": lambda: ('{"a":{"b":2}}', "a.b"),
     "py_ref": lambda a: "2",
     "result_type": "string"},

    # ---- JsonGet — array index inside object property ----
    {"name": "JsonGet_array_in_object",
     "func": "JsonGet",
     "args": lambda: ('{"arr":[10,20,30]}', "arr[1]"),
     "py_ref": lambda a: "20",
     "result_type": "string"},

    # ---- JsonGet — nonexistent path returns null ----
    {"name": "JsonGet_nonexistent_path",
     "func": "JsonGet",
     "args": lambda: ('{"a":1}', "nonexistent"),
     "py_ref": lambda a: None,
     "result_type": "scalar",
     "skip_if": True,
     "skip_reason": "VBA raises ERR_PATH_NOT_FOUND for nonexistent key; error becomes COM exception"},
    {"name": "JsonGet_empty_path",
     "func": "JsonGet",
     "args": lambda: ('{"a":1}', ""),
     "py_ref": lambda a: None,
     "result_type": "scalar",
     "skip_if": True,
     "skip_reason": "VBA raises error for empty path; error becomes COM exception"},

    # ---- JsonIsValid — large scientific notation (1e308) ----
    {"name": "JsonIsValid_true_for_1e308",
     "func": "JsonIsValid",
     "args": lambda: ("1e308",),
     "py_ref": lambda a: True,
     "result_type": "bool"},

    # ---- JsonIsValid — "undefined" is not valid JSON ----
    {"name": "JsonIsValid_false_for_undefined",
     "func": "JsonIsValid",
     "args": lambda: ("undefined",),
     "py_ref": lambda a: False,
     "result_type": "bool"},

    # ---- JsonIsValid — "NaN" is not valid JSON ----
    {"name": "JsonIsValid_false_for_NaN",
     "func": "JsonIsValid",
     "args": lambda: ("NaN",),
     "py_ref": lambda a: False,
     "result_type": "bool"},
]


def main() -> int:
    runner = CrossValRunner("JsonUtils", MODULE_PATHS)
    runner.run_all(TEST_CASES)
    passed, failed = runner.print_summary()
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
