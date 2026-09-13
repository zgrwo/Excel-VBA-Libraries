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


# 2026-09-13 回归: 空数组路径/控制字符/溢出/Range 双路径
def _json_err_probe(excel, wb, ws, runner, tc, args):
    """VBA 探针: 在 VBA 内部捕获错误码并返回值 (COM 不传递 VBA 错误描述)."""
    vbproj = wb.VBProject
    for comp in list(vbproj.VBComponents):
        if comp.Name == "JsonProbe":
            vbproj.VBComponents.Remove(comp)
    comp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
    comp.Name = "JsonProbe"
    comp.CodeModule.AddFromString(
        "Public Function ProbeErr(ByVal which As String) As Double\r\n"
        "    On Error GoTo EH\r\n"
        "    Dim v As Variant\r\n"
        "    Dim s As String\r\n"
        "    If which = \"overflow\" Then\r\n"
        "        v = JsonParse(\"9e308\")\r\n"
        "    ElseIf which = \"bigint\" Then\r\n"
        "        v = JsonParse(String$(400, \"9\"))\r\n"
        "    ElseIf which = \"bigfrac\" Then\r\n"
        "        v = JsonParse(String$(400, \"9\") & \".5\")\r\n"
        "    ElseIf which = \"depth128\" Then\r\n"
        "        v = JsonParse(String$(128, \"[\") & \"1\" & String$(128, \"]\"))\r\n"
        "        ProbeErr = 1\r\n"
        "        Exit Function\r\n"
        "    ElseIf which = \"depth_over\" Then\r\n"
        "        v = JsonParse(String$(200, \"[\") & \"1\" & String$(200, \"]\"))\r\n"
        "    ElseIf which = \"stringify_empty\" Then\r\n"
        "        If JsonStringify(JsonParse(\"[]\")) = \"[]\" Then ProbeErr = 1 Else ProbeErr = 0\r\n"
        "        Exit Function\r\n"
        "    ElseIf which = \"stringify_unalloc\" Then\r\n"
        "        Dim arrU() As Variant\r\n"
        "        If JsonStringify(arrU) = \"[]\" Then ProbeErr = 1 Else ProbeErr = 0\r\n"
        "        Exit Function\r\n"
        "    ElseIf which = \"fraction_tiny\" Then\r\n"
        "        s = JsonStringify(0.000000000000001)\r\n"
        "        If Not JsonIsValid(s) Then\r\n"
        "            ProbeErr = -1\r\n"
        "            Exit Function\r\n"
        "        End If\r\n"
        "        ProbeErr = CDbl(JsonParse(s))\r\n"
        "        Exit Function\r\n"
        "    ElseIf which = \"fraction_1e7\" Then\r\n"
        "        s = JsonStringify(0.0000001)\r\n"
        "        If Not JsonIsValid(s) Then\r\n"
        "            ProbeErr = -1\r\n"
        "            Exit Function\r\n"
        "        End If\r\n"
        "        ProbeErr = CDbl(JsonParse(s))\r\n"
        "        Exit Function\r\n"
        "    ElseIf which = \"get_array_root\" Then\r\n"
        "        ProbeErr = JsonGet(JsonParse(\"[10,20,30]\"), \"[0]\")\r\n"
        "        Exit Function\r\n"
        "    ElseIf which = \"get_array_root_str\" Then\r\n"
        "        If JsonGet(JsonParse(\"[\"\"a\"\",\"\"b\"\"]\"), \"[1]\") = \"b\" Then ProbeErr = 1 Else ProbeErr = 0\r\n"
        "        Exit Function\r\n"
        "    Else\r\n"
        "        v = JsonGet(\"[]\", \"[0]\")\r\n"
        "    End If\r\n"
        "    ProbeErr = 0\r\n"
        "    Exit Function\r\n"
        "EH: ProbeErr = Err.Number\r\n"
        "End Function")
    from tests.test_utils import run_macro
    err_num = run_macro(excel, wb, "JsonProbe.ProbeErr", args[0])
    tol = float(args[2]) if len(args) > 2 else 0.0
    return float(err_num), float(args[1]), tol


TEST_CASES += [
    {"name": "JsonIsValid_control_char", "func": "JsonIsValid",
     "args": lambda: ('"a\nb"',), "py_ref": lambda a: False, "result_type": "bool"},
    # ERR_INVALID_JSON = vbObjectError + 1301 = -2147220203
    {"name": "JsonParse_overflow", "func": "JsonParse",
     "args": lambda: ("overflow", -2147220203.0),
     "reconstruct": _json_err_probe, "py_ref": lambda a: -2147220203.0},
    # ERR_PATH_NOT_FOUND = vbObjectError + 1302 = -2147220202
    {"name": "JsonGet_empty_array_index", "func": "JsonGet",
     "args": lambda: ("emptyidx", -2147220202.0),
     "reconstruct": _json_err_probe, "py_ref": lambda a: -2147220202.0},
    {"name": "JsonStringify_Range", "func": "JsonStringify", "is_udf": True,
     "args": lambda: ([[1, 2], [3, 4]],),
     "py_ref": lambda a: "[[1,2],[3,4]]", "result_type": "string"},
    # 内部往返: JsonStringify(JsonParse("[]")) 必须得 "[]" (此前 Error 9)
    {"name": "JsonStringify_parsed_empty", "func": "JsonParse",
     "args": lambda: ("stringify_empty", 1.0),
     "reconstruct": _json_err_probe, "py_ref": lambda a: 1.0},
    # R4-05 回归: 转义序列之后的代理对/高位 BMP 不得被 AscW 有符号比较误判为控制字符
    {"name": "JsonParse_escaped_emoji", "func": "JsonParse",
     "args": lambda: (r'"a\u0041' + "\U0001F600" + '"',),
     "py_ref": lambda a: "aA\U0001F600", "result_type": "string"},
    {"name": "JsonParse_escaped_ufffd", "func": "JsonParse",
     "args": lambda: (r'"x\u0041' + "\uFFFD" + '"',),
     "py_ref": lambda a: "xA\uFFFD", "result_type": "string"},

    # ---- R5-03: |x|<1 补前导 0 (json.dumps 独立参考) + JsonIsValid 往返 ----
    {"name": "JsonStringify_fraction_half", "func": "JsonStringify",
     "args": lambda: (0.5,),
     "py_ref": lambda a: _json.dumps(a[0], separators=(',', ':')),
     "result_type": "string"},
    {"name": "JsonStringify_fraction_neg_quarter", "func": "JsonStringify",
     "args": lambda: (-0.25,),
     "py_ref": lambda a: _json.dumps(a[0], separators=(',', ':')),
     "result_type": "string"},
    {"name": "JsonStringify_fraction_tiny_roundtrip", "func": "JsonParse",
     "args": lambda: ("fraction_tiny", 1e-15, 1e-30),
     "reconstruct": _json_err_probe, "py_ref": lambda a: 1e-15},
    {"name": "JsonStringify_fraction_1e7_roundtrip", "func": "JsonParse",
     "args": lambda: ("fraction_1e7", 1e-7, 1e-22),
     "reconstruct": _json_err_probe, "py_ref": lambda a: 1e-7},
    # ---- R5-17: 超长数字字面量 → ERR_INVALID_JSON (而非裸 Error 6) ----
    {"name": "JsonParse_bigint_overflow", "func": "JsonParse",
     "args": lambda: ("bigint", -2147220203.0),
     "reconstruct": _json_err_probe, "py_ref": lambda a: -2147220203.0},
    {"name": "JsonParse_bigfrac_overflow", "func": "JsonParse",
     "args": lambda: ("bigfrac", -2147220203.0),
     "reconstruct": _json_err_probe, "py_ref": lambda a: -2147220203.0},
    # ---- R5-18: 深度上限 128 (128 层 OK / 200 层 ERR_INVALID_JSON) ----
    {"name": "JsonParse_depth128_ok", "func": "JsonParse",
     "args": lambda: ("depth128", 1.0),
     "reconstruct": _json_err_probe, "py_ref": lambda a: 1.0},
    {"name": "JsonParse_depth_over_limit", "func": "JsonParse",
     "args": lambda: ("depth_over", -2147220203.0),
     "reconstruct": _json_err_probe, "py_ref": lambda a: -2147220203.0},
    # ---- R5-34: 未分配数组 Stringify 必须返回 "[]" (而非 Error 9) ----
    {"name": "JsonStringify_unallocated_array", "func": "JsonParse",
     "args": lambda: ("stringify_unalloc", 1.0),
     "reconstruct": _json_err_probe, "py_ref": lambda a: 1.0},
    # ---- R5-48: 已解析数组根支持 "[0]" 数字/字符串索引 ----
    {"name": "JsonGet_parsed_array_root", "func": "JsonParse",
     "args": lambda: ("get_array_root", 10.0),
     "reconstruct": _json_err_probe, "py_ref": lambda a: 10.0},
    {"name": "JsonGet_parsed_array_root_string", "func": "JsonParse",
     "args": lambda: ("get_array_root_str", 1.0),
     "reconstruct": _json_err_probe, "py_ref": lambda a: 1.0},
]


def main() -> int:
    runner = CrossValRunner("JsonUtils", MODULE_PATHS)
    runner.run_all(TEST_CASES)
    passed, failed = runner.print_summary()
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
