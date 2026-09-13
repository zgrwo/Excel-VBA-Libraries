"""Cross-validate RegexUtils functions against Python reference implementations.

Usage: python tests/build_RegexUtils.py
"""

import os
import sys
import re

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests.crossval.build_common import CrossValRunner
from tests.test_utils import SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER

MODULE_PATHS = [os.path.join(VBA_CORE_DIR, name + ".cls")
                for name in VBA_CORE_IMPORT_ORDER]
MODULE_PATHS.append(os.path.join(SRC_DIR, "RegexUtils.bas"))


# 2026-09-13 R4-03/R4-06/R5-19 回归: 错误值不得字符串化; 1D Variant 数组路径;
# RegExp Static 缓存连续 pattern/flags 切换不串状态
def _regex_r4_probe(excel, wb, ws, runner, tc, args):
    """注入专用 VBA 探针 (在 VBA 内捕获错误, 避免错误变体跨 COM)。"""
    from tests.test_utils import run_macro
    vbproj = wb.VBProject
    for comp in list(vbproj.VBComponents):
        if comp.Name == "RegexR4Probe":
            vbproj.VBComponents.Remove(comp)
    comp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
    comp.Name = "RegexR4Probe"
    comp.CodeModule.AddFromString(
        "Option Explicit\r\n"
        "Public Function Probe(ByVal which As String) As Double\r\n"
        "    On Error GoTo EH\r\n"
        "    Dim v As Variant\r\n"
        "    Dim n As Long\r\n"
        "    Dim errNum As Long\r\n"
        "    If which = \"count_1d\" Then\r\n"
        "        v = RegexCount(Array(\"a1\", \"b2\"), \"\\d+\")\r\n"
        "        Probe = CDbl(v)\r\n"
        "        Exit Function\r\n"
        "    ElseIf which = \"udf_count_1d\" Then\r\n"
        "        v = UDF_REGEX_COUNT(Array(\"a1\", \"b2\"), \"\\d+\")\r\n"
        "        Probe = CDbl(v)\r\n"
        "        Exit Function\r\n"
        "    ElseIf which = \"err_count\" Then\r\n"
        "        v = RegexCount(CVErr(2007), \"\\d+\")\r\n"
        "        Probe = 0\r\n"
        "        Exit Function\r\n"
        "    ElseIf which = \"err_udf\" Then\r\n"
        "        v = UDF_REGEX_ISMATCH(CVErr(2007), \"\\d+\")\r\n"
        "        If IsError(v) Then Probe = 1 Else Probe = 0\r\n"
        "        Exit Function\r\n"
        "    ElseIf which = \"cache_switch\" Then\r\n"
        "        If RegexIsMatch(\"ABC\", \"abc\", True) Then n = n + 1\r\n"
        "        If Not RegexIsMatch(\"ABC\", \"abc\", False) Then n = n + 2\r\n"
        "        If RegexIsMatch(\"a\" & vbLf & \"b\", \"^b$\", True, True) Then n = n + 4\r\n"
        "        If Not RegexIsMatch(\"a\" & vbLf & \"b\", \"^b$\", True, False) Then n = n + 8\r\n"
        "        If RegexIsMatch(\"ABC\", \"abc\", True) Then n = n + 16\r\n"
        "        v = RegexExtractGroups(\"x=1;y=2\", \"(\\w)=(\\d)\", True, True)\r\n"
        "        If UBound(v, 1) = 1 Then\r\n"
        "            If v(1, 0) = \"y\" Then n = n + 32\r\n"
        "        End If\r\n"
        "        If Not RegexIsMatch(\"ABC\", \"abc\", False) Then n = n + 64\r\n"
        "        If RegexIsFullMatch(\"1\", \"\\d\", True, True) Then n = n + 128\r\n"
        "        v = RegexExtract(\"a1b2\", \"\\d\", 0, True, True)\r\n"
        "        If CStr(v) = \"1, 2\" Then n = n + 256\r\n"
        "        v = RegexReplace(\"a1b2\", \"\\d\", \"X\", -1)\r\n"
        "        If CStr(v) = \"a1bX\" Then n = n + 512\r\n"
        "        v = RegexExtract(\"a1b2\", \"\\d\", 0, True, True)\r\n"
        "        If CStr(v) = \"1, 2\" Then n = n + 1024\r\n"
        "        Probe = CDbl(n)\r\n"
        "        Exit Function\r\n"
        "    ElseIf which = \"cache_after_error\" Then\r\n"
        "        On Error Resume Next\r\n"
        "        v = RegexIsMatch(\"abc\", \"abc\", True)\r\n"
        "        errNum = Err.Number\r\n"
        "        Err.Clear\r\n"
        "        If errNum = 0 Then\r\n"
        "            v = RegexIsMatch(\"abc\", \"[\")\r\n"
        "            errNum = Err.Number\r\n"
        "            Err.Clear\r\n"
        "        End If\r\n"
        "        On Error GoTo EH\r\n"
        "        If errNum = 0 Then\r\n"
        "            Probe = -2\r\n"
        "        ElseIf RegexIsMatch(\"ABC\", \"abc\", True) Then\r\n"
        "            Probe = 1\r\n"
        "        End If\r\n"
        "        Exit Function\r\n"
        "    End If\r\n"
        "    Probe = -1\r\n"
        "    Exit Function\r\n"
        "EH:\r\n"
        "    If which = \"err_count\" Then Probe = 1 Else Probe = -Err.Number\r\n"
        "End Function")
    mode = args[0]
    val = run_macro(excel, wb, "RegexR4Probe.Probe", mode)
    expected = {"count_1d": 2.0, "udf_count_1d": 2.0,
                "err_count": 1.0, "err_udf": 1.0,
                "cache_switch": 2047.0, "cache_after_error": 1.0}[mode]
    return float(val), expected, 0.0

# =============================================================================
# Test Cases
# =============================================================================

TEST_CASES = [

    # ---- RegexIsMatch ----
    {
        "name": "RegexIsMatch_digits_in_text",
        "func": "RegexIsMatch",
        "args": lambda: ("abc123", r"\d+"),
        "py_ref": lambda a: bool(re.search(a[1], a[0])),
        "result_type": "bool",
    },
    {
        "name": "RegexIsMatch_no_match",
        "func": "RegexIsMatch",
        "args": lambda: ("abc", r"^\d+$"),
        "py_ref": lambda a: bool(re.search(a[1], a[0])),
        "result_type": "bool",
    },
    {
        "name": "RegexIsMatch_empty_input",
        "func": "RegexIsMatch",
        "args": lambda: ("", r"\d+"),
        "py_ref": lambda a: bool(re.search(a[1], a[0])),
        "result_type": "bool",
    },
    {
        "name": "RegexIsMatch_word_boundary",
        "func": "RegexIsMatch",
        "args": lambda: ("hello world", r"\bworld\b"),
        "py_ref": lambda a: bool(re.search(a[1], a[0])),
        "result_type": "bool",
    },

    # ---- RegexExtract ----
    {
        "name": "RegexExtract_phone_number",
        "func": "RegexExtract",
        "args": lambda: ("Phone: 555-1234", r"\d{3}-\d{4}"),
        "py_ref": lambda a: re.search(a[1], a[0]).group() if re.search(a[1], a[0]) else "",
        "result_type": "string",
    },
    {
        "name": "RegexExtract_no_match",
        "func": "RegexExtract",
        "args": lambda: ("no numbers here", r"\d+"),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "RegexExtract_email",
        "func": "RegexExtract",
        "args": lambda: ("Contact: user@example.com for help", r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"),
        "py_ref": lambda a: re.search(a[1], a[0]).group() if re.search(a[1], a[0]) else "",
        "result_type": "string",
    },

    # ---- RegexExtractAll ----
    {
        "name": "RegexExtractAll_word_digit_pairs",
        "func": "RegexExtractAll",
        "args": lambda: ("a1 b2 c3", r"\w\d"),
        "py_ref": lambda a: re.findall(a[1], a[0]),
        "result_type": "array",
    },
    {
        "name": "RegexExtractAll_no_match",
        "func": "RegexExtractAll",
        "args": lambda: ("abcdef", r"\d+"),
        "py_ref": lambda a: [],
        "result_type": "array",
    },
    {
        "name": "RegexExtractAll_numbers",
        "func": "RegexExtractAll",
        "args": lambda: ("x=10, y=20, z=30", r"\d+"),
        "py_ref": lambda a: re.findall(a[1], a[0]),
        "result_type": "array",
    },

    # ---- RegexExtractGroups ----
    {
        "name": "RegexExtractGroups_name_age",
        "func": "RegexExtractGroups",
        "args": lambda: ("Name: John, Age: 30", r"(\w+): (\w+)"),
        "py_ref": lambda a: [list(m) for m in re.findall(a[1], a[0])],
        "result_type": "array",
    },
    {
        "name": "RegexExtractGroups_date",
        "func": "RegexExtractGroups",
        "args": lambda: ("2024-01-15", r"(\d{4})-(\d{2})-(\d{2})"),
        "py_ref": lambda a: [list(m) for m in re.findall(a[1], a[0])],
        "result_type": "array",
    },

    # ---- RegexIsFullMatch ----
    {
        "name": "RegexIsFullMatch_all_digits",
        "func": "RegexIsFullMatch",
        "args": lambda: ("12345", r"\d+"),
        "py_ref": lambda a: bool(re.fullmatch(a[1], a[0])),
        "result_type": "bool",
    },
    {
        "name": "RegexIsFullMatch_partial_match",
        "func": "RegexIsFullMatch",
        "args": lambda: ("abc123", r"\d+"),
        "py_ref": lambda a: bool(re.fullmatch(a[1], a[0])),
        "result_type": "bool",
    },
    {
        "name": "RegexIsFullMatch_exact_word",
        "func": "RegexIsFullMatch",
        "args": lambda: ("hello", r"\w+"),
        "py_ref": lambda a: bool(re.fullmatch(a[1], a[0])),
        "result_type": "bool",
    },
    {
        "name": "RegexIsFullMatch_dot_star",
        "func": "RegexIsFullMatch",
        "args": lambda: ("abc", r".*"),
        "py_ref": lambda a: True,  # .* should match entire string (#77)
        "result_type": "bool",
    },

    # ---- RegexReplace ----
    {
        "name": "RegexReplace_date_reformat",
        "func": "RegexReplace",
        "args": lambda: ("2024-01-15", r"(\d{4})-(\d{2})-(\d{2})", r"$3/$2/$1"),
        "py_ref": lambda a: "15/01/2024",
        "result_type": "string",
    },
    {
        "name": "RegexReplace_collapse_spaces",
        "func": "RegexReplace",
        "args": lambda: ("a   b    c", r"\s+", " "),
        "py_ref": lambda a: re.sub(a[1], a[2], a[0]),
        "result_type": "string",
    },
    {
        "name": "RegexReplace_no_match",
        "func": "RegexReplace",
        "args": lambda: ("hello", r"\d+", "X"),
        "py_ref": lambda a: re.sub(a[1], a[2], a[0]),
        "result_type": "string",
    },
    {
        "name": "RegexReplace_empty_input",
        "func": "RegexReplace",
        "args": lambda: ("", r"\s+", " "),
        "py_ref": lambda a: "",
        "result_type": "string",
    },

    # ---- RegexSplit ----
    {
        "name": "RegexSplit_mixed_delimiters",
        "func": "RegexSplit",
        "args": lambda: ("a,b; c|d", r"[,;|]\s*"),
        # NOTE: py_ref filters empty strings ([s for s in ... if s]).
        # VBA RegexSplit does NOT filter empties. This is safe for current
        # test data (no consecutive delimiters) but could mask issues if
        # test data changes.
        "py_ref": lambda a: [s for s in re.split(a[1], a[0]) if s],
        "result_type": "array",
    },
    {
        "name": "RegexSplit_comma_separated",
        "func": "RegexSplit",
        "args": lambda: ("one,two,three", r","),
        "py_ref": lambda a: re.split(a[1], a[0]),
        "result_type": "array",
    },
    {
        "name": "RegexSplit_no_delimiter_in_string",
        "func": "RegexSplit",
        "args": lambda: ("hello", r","),
        "py_ref": lambda a: re.split(a[1], a[0]),
        "result_type": "array",
    },
    {
        "name": "RegexSplit_empty_input",
        "func": "RegexSplit",
        "args": lambda: ("", r","),
        "py_ref": lambda a: re.split(a[1], a[0]),
        "result_type": "array",
    },

    # ---- RegexCount ----
    {
        "name": "RegexCount_three_letter_words",
        "func": "RegexCount",
        "args": lambda: ("The fat cat sat on the mat", r"\b\w{3}\b"),
        "py_ref": lambda a: len(re.findall(a[1], a[0])),
        "result_type": "scalar",
    },
    {
        "name": "RegexCount_no_match",
        "func": "RegexCount",
        "args": lambda: ("abcdef", r"\d+"),
        "py_ref": lambda a: 0,
        "result_type": "scalar",
    },
    {
        "name": "RegexCount_digits",
        "func": "RegexCount",
        "args": lambda: ("a1b2c3d4", r"\d"),
        "py_ref": lambda a: len(re.findall(a[1], a[0])),
        "result_type": "scalar",
    },

    # ---- RegexEscape ----
    {
        "name": "RegexEscape_math_expression",
        "func": "RegexEscape",
        "args": lambda: ("1+1=2?",),
        "py_ref": lambda a: re.escape(a[0]),
        "result_type": "string",
    },
    {
        "name": "RegexEscape_dots_and_stars",
        "func": "RegexEscape",
        "args": lambda: ("a.b*c",),
        "py_ref": lambda a: re.escape(a[0]),
        "result_type": "string",
    },
    {
        "name": "RegexEscape_brackets",
        "func": "RegexEscape",
        "args": lambda: ("[test]",),
        "py_ref": lambda a: re.escape(a[0]),
        "result_type": "string",
    },
    {
        "name": "RegexEscape_plain_text",
        "func": "RegexEscape",
        "args": lambda: ("hello",),
        "py_ref": lambda a: "hello",
        "result_type": "string",
    },
    {
        "name": "RegexEscape_empty",
        "func": "RegexEscape",
        "args": lambda: ("",),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    # =====================================================================
    # Migrated from VBA Test_RegexUtils — coverage gaps (2026-06-16)
    # =====================================================================

    # ---- RegexIsMatch — case-insensitive (VBA default IgnoreCase=True) ----
    {"name": "RegexIsMatch_case_insensitive", "func": "RegexIsMatch",
     "args": lambda: ("ABC", r"[a-z]+"),
     "py_ref": lambda a: bool(re.search(a[1], a[0], re.IGNORECASE)),
     "result_type": "bool"},

    # ---- RegexExtract — occurrence index (VBA 3rd param) ----
    {"name": "RegexExtract_occurrence_last", "func": "RegexExtract",
     "args": lambda: ("a1b2c3", r"\d+", -1),
     "py_ref": lambda a: "3",
     "result_type": "string"},
    {"name": "RegexExtract_occurrence_first", "func": "RegexExtract",
     "args": lambda: ("a1b2c3", r"\d+", 1),
     "py_ref": lambda a: "1",
     "result_type": "string"},

    # ---- RegexReplace — empty replacement / multiline ----
    {"name": "RegexReplace_empty_repl", "func": "RegexReplace",
     "args": lambda: ("abc", "b", ""),
     "py_ref": lambda a: "ac",
     "result_type": "string"},
    # VBScript regex treats \r and \n as independent line breaks in Multiline mode
    {"name": "RegexReplace_multiline", "func": "RegexReplace",
     "args": lambda: ("a\r\nb", "^", ">", 0, True, True, True),
     "py_ref": lambda a: ">a\r>\n>b",
     "result_type": "string"},

    # ---- RegexSplit — consecutive delimiters ----
    {"name": "RegexSplit_consecutive_delims", "func": "RegexSplit",
     "args": lambda: ("a,,b,,c", ",+"),
     "py_ref": lambda a: ["a", "b", "c"],
     "result_type": "array"},

    # ---- RegexExtractGroups — no match returns empty array ----
    {"name": "RegexExtractGroups_no_match", "func": "RegexExtractGroups",
     "args": lambda: ("abc", r"(\d+)"),
     "py_ref": lambda a: [],
     "result_type": "array"},

    # ---- Error case: invalid regex (CVErr via COM unreliable) ----
    {"name": "RegexIsMatch_bad_pattern", "func": "RegexIsMatch",
     "args": lambda: ("test", "["), "py_ref": lambda a: None,
     "result_type": "bool",
     "skip_if": True, "skip_reason": "VBA raises runtime error for bad pattern; COM marshaling unreliable"},

    # UDF wrapper
    {"name": "UDF_REGEX_ISMATCH", "func": "UDF_REGEX_ISMATCH",
     "args": lambda: (["abc123", "xyz"], r"\d+"),
     "py_ref": lambda a: [True, False],
     "result_type": "array", "compare_mode": "bool_array",
     "skip_if": True,
     "skip_reason": "UDF Regex wrapper expects 2D Range input through COM; core RegexIsMatch is covered by crossval tests"},

    # =====================================================================
    # 2026-09-13 R4 回归: 错误值不得字符串化 (R4-03); 1D 数组路径 (R4-06)
    # =====================================================================
    {"name": "RegexCount_1d_array", "func": "RegexCount",
     "args": lambda: ("count_1d",), "reconstruct": _regex_r4_probe,
     "py_ref": lambda a: 2.0},
    {"name": "UDF_COUNT_1d_array", "func": "UDF_REGEX_COUNT",
     "args": lambda: ("udf_count_1d",), "reconstruct": _regex_r4_probe,
     "py_ref": lambda a: 2.0},
    {"name": "RegexCount_error_scalar", "func": "RegexCount",
     "args": lambda: ("err_count",), "reconstruct": _regex_r4_probe,
     "py_ref": lambda a: 1.0},
    {"name": "UDF_ISMATCH_error_scalar", "func": "UDF_REGEX_ISMATCH",
     "args": lambda: ("err_udf",), "reconstruct": _regex_r4_probe,
     "py_ref": lambda a: 1.0},

    # =====================================================================
    # 2026-09-13 R5-19 回归: RegExp Static 缓存 — 连续 pattern/flags 切换
    # (IgnoreCase/MultiLine/Global 不串状态), 非法模式不污染缓存
    # =====================================================================
    {"name": "Regex_cache_pattern_switch", "func": "RegexIsMatch",
     "args": lambda: ("cache_switch",), "reconstruct": _regex_r4_probe,
     "py_ref": lambda a: 2047.0},
    {"name": "Regex_cache_survives_invalid_pattern", "func": "RegexIsMatch",
     "args": lambda: ("cache_after_error",), "reconstruct": _regex_r4_probe,
     "py_ref": lambda a: 1.0},

]


def main() -> int:
    runner = CrossValRunner("RegexUtils", MODULE_PATHS)
    runner.run_all(TEST_CASES)
    passed, failed = runner.print_summary()
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
