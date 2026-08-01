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

]


def main() -> int:
    runner = CrossValRunner("RegexUtils", MODULE_PATHS)
    runner.run_all(TEST_CASES)
    passed, failed = runner.print_summary()
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
