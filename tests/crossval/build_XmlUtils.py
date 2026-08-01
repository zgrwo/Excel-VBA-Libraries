"""Cross-validate XmlUtils functions against Python xml.etree.ElementTree.

Usage: python tests/crossval/build_XmlUtils.py
"""

import os
import sys
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests.crossval.build_common import CrossValRunner
from tests.test_utils import SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER

MODULE_PATHS = [os.path.join(VBA_CORE_DIR, name + ".cls")
                for name in VBA_CORE_IMPORT_ORDER]
MODULE_PATHS.append(os.path.join(SRC_DIR, "XmlUtils.bas"))


# =============================================================================
# Python Reference Helpers
# =============================================================================

def py_validate(xml_str):
    try:
        ET.fromstring(xml_str)
        return True
    except ET.ParseError:
        return False


def py_xpath_text(xml_str, xpath):
    root = ET.fromstring(xml_str)
    el = root.find(xpath)
    return (el.text or "") if el is not None else None


def py_get_attr(xml_str, xpath, attr_name):
    root = ET.fromstring(xml_str)
    el = root.find(xpath)
    if el is None:
        return None
    return el.get(attr_name, None) or ""


def py_to_array(xml_str, row_path, col_names):
    root = ET.fromstring(xml_str)
    rows = root.findall(row_path)
    result = []
    for row in rows:
        row_data = []
        for col in col_names:
            cel = row.find(col)
            row_data.append((cel.text or "") if cel is not None else "")
        result.append(row_data)
    return result


# =============================================================================
# Test Cases
# =============================================================================

TEST_CASES = [

    # ---- XmlValidate ----
    {
        "name": "XmlValidate_valid",
        "func": "XmlValidate",
        "args": lambda: ("<root/>", ""),
        "py_ref": lambda a: py_validate(a[0]),
        "result_type": "bool",
    },
    {
        "name": "XmlValidate_invalid_mismatch",
        "func": "XmlValidate",
        "args": lambda: ("<root><bad></root>", ""),
        "py_ref": lambda a: py_validate(a[0]),
        "result_type": "bool",
    },
    {
        "name": "XmlValidate_empty_string",
        "func": "XmlValidate",
        "args": lambda: ("", ""),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },
    {
        "name": "XmlValidate_nested_valid",
        "func": "XmlValidate",
        "args": lambda: ("<a><b><c>deep</c></b></a>", ""),
        "py_ref": lambda a: py_validate(a[0]),
        "result_type": "bool",
    },

    # ---- XmlGet ----
    {
        "name": "XmlGet_simple_path",
        "func": "XmlGet",
        "args": lambda: ("<a><b>hello</b></a>", "/a/b"),
        "py_ref": lambda a: "hello",
        "result_type": "string",
    },
    {
        "name": "XmlGet_deep_nested",
        "func": "XmlGet",
        "args": lambda: ("<a><b><c>deep</c></b></a>", "/a/b/c"),
        "py_ref": lambda a: "deep",
        "result_type": "string",
    },
    {
        "name": "XmlGet_not_found",
        "func": "XmlGet",
        "args": lambda: ("<a><b>hello</b></a>", "/a/c"),
        "py_ref": lambda a: None,
        "result_type": "scalar",
        "skip_if": True,
        "skip_reason": "VBA returns CVErr which does not COM-marshal cleanly",
    },

    # ---- XmlGetAttr ----
    {
        "name": "XmlGetAttr_basic",
        "func": "XmlGetAttr",
        "args": lambda: ("<user id='007' name='Bond'/>", "/user", "id"),
        "py_ref": lambda a: "007",
        "result_type": "string",
    },
    {
        "name": "XmlGetAttr_second",
        "func": "XmlGetAttr",
        "args": lambda: ("<user id='007' name='Bond'/>", "/user", "name"),
        "py_ref": lambda a: "Bond",
        "result_type": "string",
    },
    {
        "name": "XmlGetAttr_missing",
        "func": "XmlGetAttr",
        "args": lambda: ("<user id='007'/>", "/user", "role"),
        "py_ref": lambda a: "",
        "result_type": "string",
        "skip_if": True,
        "skip_reason": "VBA returns CVErr(xlErrNA), Python returns empty string",
    },

    # ---- XmlToRange (numeric data only — crossval uses numpy) ----
    {
        "name": "XmlToRange_two_rows",
        "func": "XmlToRange",
        "args": lambda: (
            "<rows><row><a>1</a><b>10</b></row><row><a>2</a><b>20</b></row></rows>",
            "/rows/row",
            ["a", "b"],
        ),
        "py_ref": lambda a: [[1.0, 10.0], [2.0, 20.0]],
        "result_type": "array",
        "tol": 1e-10,
    },
    {
        "name": "XmlToRange_single_col",
        "func": "XmlToRange",
        "args": lambda: (
            "<rows><row><a>1</a></row><row><a>2</a></row></rows>",
            "/rows/row",
            ["a"],
        ),
        "py_ref": lambda a: [[1.0], [2.0]],
        "result_type": "array",
        "tol": 1e-10,
    },
    {
        "name": "XmlToRange_auto_detect",
        "func": "XmlToRange",
        "args": lambda: (
            "<rows><row><a>1</a><b>10</b></row></rows>",
            "/rows/row",
        ),
        "py_ref": lambda a: [[1.0, 10.0]],
        "result_type": "array",
        "tol": 1e-10,
    },

    # ---- UDF_XML_TABLE ----
    {
        "name": "UDF_XML_TABLE_specified_cols",
        "func": "UDF_XML_TABLE",
        "args": lambda: (
            "<rows><row><a>1</a><b>10</b></row><row><a>2</a><b>20</b></row></rows>",
            "/rows/row",
            ["a", "b"],
        ),
        "py_ref": lambda a: [[1.0, 10.0], [2.0, 20.0]],
        "result_type": "array",
        "tol": 1e-10,
    },

    # ---- Edge Cases: XmlGet with complex XPath ----
    {
        "name": "XmlGet_xpath_index",
        "func": "XmlGet",
        "args": lambda: ("<list><item>A</item><item>B</item></list>", "/list/item[2]"),
        "py_ref": lambda a: "B",
        "result_type": "string",
    },
    {
        "name": "XmlGet_declaration_skip",
        "func": "XmlGet",
        "args": lambda: ('<?xml version="1.0"?><root>ok</root>', "/root"),
        "py_ref": lambda a: "ok",
        "result_type": "string",
    },

    # ---- Edge Cases: XmlValidate ----
    {
        "name": "XmlValidate_no_closing_tag",
        "func": "XmlValidate",
        "args": lambda: ("<root>", ""),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },
    {
        "name": "XmlValidate_plain_text",
        "func": "XmlValidate",
        "args": lambda: ("not xml at all", ""),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },

    # ---- UDF_XML_VALIDATE ----
    {
        "name": "UDF_XML_VALIDATE_true",
        "func": "UDF_XML_VALIDATE",
        "args": lambda: ("<root/>",),
        "py_ref": lambda a: True,
        "result_type": "bool",
    },
    {
        "name": "UDF_XML_VALIDATE_false",
        "func": "UDF_XML_VALIDATE",
        "args": lambda: ("<bad>",),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },

    # ---- UDF_XML_GET ----
    {
        "name": "UDF_XML_GET_value",
        "func": "UDF_XML_GET",
        "args": lambda: ("<a><b>42</b></a>", "/a/b"),
        "py_ref": lambda a: "42",
        "result_type": "string",
    },

    # ---- Edge Cases ----
    {
        "name": "XmlGet_with_entities",
        "func": "XmlGet",
        "args": lambda: ("<root>&lt;tag&gt;</root>", "/root"),
        "py_ref": lambda a: "<tag>",
        "result_type": "string",
    },
    {
        "name": "XmlGetAttr_self_closing",
        "func": "XmlGetAttr",
        "args": lambda: ("<items><item key='A'/><item key='B'/></items>", "/items/item", "key"),
        "py_ref": lambda a: "A",
        "result_type": "string",
    },
    {
        "name": "XmlGet_cdata",
        "func": "XmlGet",
        "args": lambda: ("<root><![CDATA[<inner>text</inner>]]></root>", "/root"),
        "py_ref": lambda a: "<inner>text</inner>",
        "result_type": "string",
    },

    # =====================================================================
    # Migrated from VBA Test_XmlUtils — coverage gaps (2026-06-16)
    # =====================================================================

    # ---- XmlValidate — namespace XML ----
    {
        "name": "XmlValidate_namespace",
        "func": "XmlValidate",
        "args": lambda: ("<root xmlns='http://example.com'><item>val</item></root>", ""),
        "py_ref": lambda a: True,
        "result_type": "bool",
    },

    # ---- UDF_XML_TABLE — auto-detect columns ----
    {
        "name": "UDF_XML_TABLE_auto",
        "func": "UDF_XML_TABLE",
        "args": lambda: (
            "<rows><row><a>1</a><b>X</b></row></rows>",
            "/rows/row",
        ),
        "py_ref": lambda a: [[1.0], ["X"]],  # auto-detect returns first col; string row skipped by numpy
        "result_type": "array",
        "tol": 1e-10,
        "skip_if": True,
        "skip_reason": "Mixed string/numeric data; numpy can't compare. VBA verified by Test_XmlUtils.",
    },

    # ---- Error cases (CVErr via COM unreliable) ----
    {
        "name": "XmlToRange_invalid_xpath",
        "func": "XmlToRange",
        "args": lambda: ("<rows><row><a>1</a></row></rows>", "/rows/missing"),
        "py_ref": lambda a: None,
        "result_type": "array",
        "skip_if": True,
        "skip_reason": "VBA returns CVErr for invalid XPath; COM marshaling unreliable",
    },
    {
        "name": "XmlToRange_empty_rows",
        "func": "XmlToRange",
        "args": lambda: ("<root/>", "/root/missing"),
        "py_ref": lambda a: None,
        "result_type": "array",
        "skip_if": True,
        "skip_reason": "VBA returns CVErr for no matching rows; COM marshaling unreliable",
    },
    {
        "name": "XmlValidate_with_errDetail",
        "func": "XmlValidate",
        "args": lambda: ("<root/>", ""),
        "py_ref": lambda a: True,
        "result_type": "bool",
        "skip_if": True,
        "skip_reason": "2nd param is ByRef errDetail; COM can't read back the mutated variant",
    },
]


def main() -> int:
    runner = CrossValRunner("XmlUtils", MODULE_PATHS)
    runner.run_all(TEST_CASES)
    passed, failed = runner.print_summary()
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
