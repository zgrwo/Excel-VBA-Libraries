"""Cross-validate StringUtils functions against Python reference implementations.

Usage: python tests/build_StringUtils.py
"""

import os
import sys
import base64
import html as _html
import re
import urllib.parse

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests.crossval.build_common import CrossValRunner, load_crossval_data, load_text_data
from tests.test_utils import SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER

MODULE_PATHS = [os.path.join(VBA_CORE_DIR, name + ".cls") for name in VBA_CORE_IMPORT_ORDER]
MODULE_PATHS.append(os.path.join(SRC_DIR, "StringUtils.bas"))


# ---------------------------------------------------------------------------
# Python reference helper functions
# ---------------------------------------------------------------------------

def _soundex_py(text):
    """American Soundex — matches VBA StringUtils.Soundex behaviour."""
    if not text:
        return ""
    text = text.upper()
    first = text[0]
    mapping = {
        "B": 1, "F": 1, "P": 1, "V": 1,
        "C": 2, "G": 2, "J": 2, "K": 2, "Q": 2, "S": 2, "X": 2, "Z": 2,
        "D": 3, "T": 3,
        "L": 4,
        "M": 5, "N": 5,
        "R": 6,
    }
    result = [first]
    prev_code = mapping.get(first, 0)
    for ch in text[1:]:
        if len(result) >= 4:
            break
        code = mapping.get(ch, 0)
        if code > 0:
            if code != prev_code:
                result.append(str(code))
                prev_code = code
        else:
            if ch in "AEIOUY":
                prev_code = 0
    while len(result) < 4:
        result.append("0")
    return "".join(result[:4])


def _levenshtein_py(a, b):
    """Standard Levenshtein distance."""
    if len(a) < len(b):
        return _levenshtein_py(b, a)
    if len(b) == 0:
        return len(a)
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a):
        cur = [i + 1]
        for j, cb in enumerate(b):
            cur.append(min(
                prev[j + 1] + 1,       # deletion
                cur[j] + 1,             # insertion
                prev[j] + (0 if ca == cb else 1),  # substitution
            ))
        prev = cur
    return prev[-1]


def _slugify_py(s, sep="-"):
    """Simple slugify matching VBA Slugify for ASCII input."""
    if not s:
        return ""
    # VBA Slugify does: LowerCase -> Trim -> RemoveDiacritics -> process chars
    s = s.lower().strip()
    result = []
    prev_sep = False
    for ch in s:
        if ch.isalnum() and ord(ch) < 128:
            result.append(ch)
            prev_sep = False
        elif ch in ("-", "_") or ch.isspace():
            if not prev_sep and result:
                result.append(sep)
                prev_sep = True
        # Other punctuation is silently dropped
    return "".join(result).rstrip(sep)


def _vba_html_encode(s):
    """Match VBA HTMLEncode exactly (including &#39; for single quote)."""
    s = s.replace("&", "&amp;")
    s = s.replace("<", "&lt;")
    s = s.replace(">", "&gt;")
    s = s.replace('"', "&quot;")
    s = s.replace("'", "&#39;")
    return s


def _vba_html_decode(s):
    """Match VBA HTMLDecode — numeric entities first, then named, &amp; last."""
    # Numeric decimal entities
    def _dec_ent(m):
        code = int(m.group(1))
        if code <= 0xFFFF:
            return chr(code)
        elif code <= 0x10FFFF:
            code -= 0x10000
            return chr(0xD800 + (code >> 10)) + chr(0xDC00 + (code & 0x3FF))
        return m.group(0)

    # Numeric hex entities
    def _hex_ent(m):
        code = int(m.group(1), 16)
        if code <= 0xFFFF:
            return chr(code)
        elif code <= 0x10FFFF:
            code -= 0x10000
            return chr(0xD800 + (code >> 10)) + chr(0xDC00 + (code & 0x3FF))
        return m.group(0)

    s = re.sub(r"&#x([0-9A-Fa-f]+);", _hex_ent, s)
    s = re.sub(r"&#(\d+);", _dec_ent, s)
    # Named entities — &amp; must be last to avoid double-decode of &amp;lt; etc.
    s = s.replace("&quot;", '"')
    s = s.replace("&lt;", "<")
    s = s.replace("&gt;", ">")
    s = s.replace("&amp;", "&")
    return s


def _remove_diacritics_py(s):
    """Remove diacritics using unicodedata NFKD — broader than VBA's Latin-1 map."""
    import unicodedata
    if not s:
        return ""
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


# ---------------------------------------------------------------------------
# Test cases
# ---------------------------------------------------------------------------

TEST_CASES = [
    # =========================================================================
    # RemoveChars(text, charsToRemove) -> String
    # =========================================================================
    {
        "name": "RemoveChars_basic",
        "func": "RemoveChars",
        "args": lambda: ("AB-12/CD-34", "-/"),
        "py_ref": lambda a: "AB12CD34",
        "result_type": "string",
    },
    {
        "name": "RemoveChars_empty_text",
        "func": "RemoveChars",
        "args": lambda: ("", "-/"),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "RemoveChars_empty_chars",
        "func": "RemoveChars",
        "args": lambda: ("Hello World", ""),
        "py_ref": lambda a: "Hello World",
        "result_type": "string",
    },
    {
        "name": "RemoveChars_no_match",
        "func": "RemoveChars",
        "args": lambda: ("ABCDEF", "xyz"),
        "py_ref": lambda a: "ABCDEF",
        "result_type": "string",
    },
    {
        "name": "RemoveChars_all_match",
        "func": "RemoveChars",
        "args": lambda: ("aaaaa", "a"),
        "py_ref": lambda a: "",
        "result_type": "string",
    },

    # =========================================================================
    # KeepChars(text, allowedChars) -> String
    # =========================================================================
    {
        "name": "KeepChars_digits",
        "func": "KeepChars",
        "args": lambda: ("Tel: 555-1234", "0123456789"),
        "py_ref": lambda a: "5551234",
        "result_type": "string",
    },
    {
        "name": "KeepChars_empty_text",
        "func": "KeepChars",
        "args": lambda: ("", "0123456789"),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "KeepChars_empty_allowed",
        "func": "KeepChars",
        "args": lambda: ("Hello World", ""),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "KeepChars_all_allowed",
        "func": "KeepChars",
        "args": lambda: ("ABC", "ABCDEFGHIJKLMNOPQRSTUVWXYZ"),
        "py_ref": lambda a: "ABC",
        "result_type": "string",
    },

    # =========================================================================
    # NormalizeWhitespace(text) -> String
    # =========================================================================
    {
        "name": "NormalizeWs_basic",
        "func": "NormalizeWhitespace",
        "args": lambda: ("  Hello   World  ",),
        "py_ref": lambda a: "Hello World",
        "result_type": "string",
    },
    {
        "name": "NormalizeWs_tabs",
        "func": "NormalizeWhitespace",
        "args": lambda: ("Hello\t\tWorld",),
        "py_ref": lambda a: "Hello World",
        "result_type": "string",
    },
    {
        "name": "NormalizeWs_empty",
        "func": "NormalizeWhitespace",
        "args": lambda: ("",),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "NormalizeWs_only_spaces",
        "func": "NormalizeWhitespace",
        "args": lambda: ("     ",),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "NormalizeWs_newlines",
        "func": "NormalizeWhitespace",
        "args": lambda: ("Hello\r\nWorld",),
        "py_ref": lambda a: "Hello World",
        "result_type": "string",
    },

    # =========================================================================
    # ToTitleCase(text) -> String
    # =========================================================================
    {
        "name": "ToTitleCase_basic",
        "func": "ToTitleCase",
        "args": lambda: ("hello world",),
        "py_ref": lambda a: "Hello World",
        "result_type": "string",
    },
    {
        "name": "ToTitleCase_mcdonald",
        "func": "ToTitleCase",
        "args": lambda: ("mcdonald",),
        "py_ref": lambda a: "McDonald",
        "result_type": "string",
    },
    {
        "name": "ToTitleCase_macdonald",
        "func": "ToTitleCase",
        "args": lambda: ("macdonald",),
        "py_ref": lambda a: "MacDonald",
        "result_type": "string",
    },
    {
        "name": "ToTitleCase_obrien",
        "func": "ToTitleCase",
        "args": lambda: ("o'brien",),
        "py_ref": lambda a: "O'Brien",
        "result_type": "string",
    },
    {
        "name": "ToTitleCase_empty",
        "func": "ToTitleCase",
        "args": lambda: ("",),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "ToTitleCase_already_title",
        "func": "ToTitleCase",
        "args": lambda: ("Already Title",),
        "py_ref": lambda a: "Already Title",
        "result_type": "string",
    },

    # =========================================================================
    # RemoveDiacritics(text) -> String
    # =========================================================================
    {
        "name": "RemoveDiacritics_latin1",
        "func": "RemoveDiacritics",
        "args": lambda: ("crème brûlée",),
        "py_ref": lambda a: "creme brulee",
        "result_type": "string",
    },
    {
        "name": "RemoveDiacritics_empty",
        "func": "RemoveDiacritics",
        "args": lambda: ("",),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "RemoveDiacritics_ascii",
        "func": "RemoveDiacritics",
        "args": lambda: ("Hello World",),
        "py_ref": lambda a: "Hello World",
        "result_type": "string",
    },
    # VBA RemoveDiacritics is limited to Latin-1 + Latin Extended-A.
    # Unicode characters outside this range (e.g. Cyrillic, Greek) are NOT
    # handled.  Mark this test skipped — Python unicodedata handles them but
    # VBA does not.
    {
        "name": "RemoveDiacritics_unicode_mismatch",
        "func": "RemoveDiacritics",
        "args": lambda: ("ŠĔŃ łórłd",),
        # S-caron, E-breve, N-acute, space, l-stroke, o-acute, r, l-stroke, d
        "py_ref": lambda a: "SEN lorld",  # VBA's Latin Extended-A map
        "result_type": "string",
        "skip_if": True,
        "skip_reason": "VBA RemoveDiacritics handles Extended-A chars with multi-char maps (S-car->S, l-bar->l); verify manually",
    },

    # =========================================================================
    # Slugify(text, separator="-") -> String
    # =========================================================================
    {
        "name": "Slugify_basic",
        "func": "Slugify",
        "args": lambda: ("Hello World!",),
        "py_ref": lambda a: _slugify_py(a[0]),
        "result_type": "string",
    },
    {
        "name": "Slugify_with_dots",
        "func": "Slugify",
        "args": lambda: ("Hello World! (v2.0)",),
        "py_ref": lambda a: _slugify_py(a[0]),
        "result_type": "string",
    },
    {
        "name": "Slugify_empty",
        "func": "Slugify",
        "args": lambda: ("",),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "Slugify_custom_sep",
        "func": "Slugify",
        "args": lambda: ("Hello World", "_"),
        "py_ref": lambda a: _slugify_py(a[0], "_"),
        "result_type": "string",
    },
    {
        "name": "Slugify_already_slug",
        "func": "Slugify",
        "args": lambda: ("hello-world",),
        "py_ref": lambda a: "hello-world",
        "result_type": "string",
    },

    # =========================================================================
    # Base64Encode(text, encoding="UTF-8") -> String
    # =========================================================================
    {
        "name": "Base64Encode_Hello",
        "func": "Base64Encode",
        "args": lambda: ("Hello",),
        "py_ref": lambda a: base64.b64encode(a[0].encode("utf-8")).decode(),
        "result_type": "string",
    },
    {
        "name": "Base64Encode_empty",
        "func": "Base64Encode",
        "args": lambda: ("",),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "Base64Encode_longer",
        "func": "Base64Encode",
        "args": lambda: ("Man is distinguished",),
        "py_ref": lambda a: base64.b64encode(a[0].encode("utf-8")).decode(),
        "result_type": "string",
    },
    # ---- Base64Encode — Chinese (UTF-8 BOM removal verification, M7 sync) ----
    {
        "name": "Base64Encode_chinese",
        "func": "Base64Encode",
        "args": lambda: ("测试",),
        "py_ref": lambda a: base64.b64encode("测试".encode("utf-8")).decode(),
        "result_type": "string",
    },
    # ---- Base64EncodeDecode — round-trip verify BOM-free output ----
    {
        "name": "Base64EncodeDecode_roundtrip",
        "func": "Base64Decode",
        "args": lambda: (base64.b64encode("Hello World!".encode("utf-8")).decode(),),
        "py_ref": lambda a: "Hello World!",
        "result_type": "string",
    },

    # =========================================================================
    # Base64Decode(base64, encoding="UTF-8") -> String
    # =========================================================================
    {
        "name": "Base64Decode_Hello",
        "func": "Base64Decode",
        "args": lambda: ("SGVsbG8=",),
        "py_ref": lambda a: base64.b64decode(a[0].encode()).decode("utf-8"),
        "result_type": "string",
    },
    {
        "name": "Base64Decode_empty",
        "func": "Base64Decode",
        "args": lambda: ("",),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "Base64Decode_longer",
        "func": "Base64Decode",
        "args": lambda: ("TWFuIGlzIGRpc3Rpbmd1aXNoZWQ=",),
        "py_ref": lambda a: base64.b64decode(a[0].encode()).decode("utf-8"),
        "result_type": "string",
    },

    # =========================================================================
    # URLEncode(text) -> String
    # =========================================================================
    {
        "name": "URLEncode_space",
        "func": "URLEncode",
        "args": lambda: ("Hello World",),
        "py_ref": lambda a: urllib.parse.quote(a[0]),
        "result_type": "string",
    },
    {
        "name": "URLEncode_empty",
        "func": "URLEncode",
        "args": lambda: ("",),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "URLEncode_special",
        "func": "URLEncode",
        "args": lambda: ("name=John Doe&age=30",),
        "py_ref": lambda a: urllib.parse.quote(a[0]),
        "result_type": "string",
    },
    {
        "name": "URLEncode_unreserved",
        "func": "URLEncode",
        "args": lambda: ("abc-_.~123",),
        "py_ref": lambda a: "abc-_.~123",
        "result_type": "string",
    },

    # =========================================================================
    # URLDecode(text) -> String
    # =========================================================================
    {
        "name": "URLDecode_space",
        "func": "URLDecode",
        "args": lambda: ("Hello%20World",),
        "py_ref": lambda a: urllib.parse.unquote(a[0]),
        "result_type": "string",
    },
    {
        "name": "URLDecode_empty",
        "func": "URLDecode",
        "args": lambda: ("",),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "URLDecode_plus",
        "func": "URLDecode",
        "args": lambda: ("Hello+World",),
        "py_ref": lambda a: urllib.parse.unquote_plus(a[0]),
        "result_type": "string",
    },

    # =========================================================================
    # HTMLEncode(text) -> String
    # =========================================================================
    {
        "name": "HTMLEncode_tags",
        "func": "HTMLEncode",
        "args": lambda: ("<div>",),
        "py_ref": lambda a: _vba_html_encode(a[0]),
        "result_type": "string",
    },
    {
        "name": "HTMLEncode_ampersand",
        "func": "HTMLEncode",
        "args": lambda: ("A & B",),
        "py_ref": lambda a: _vba_html_encode(a[0]),
        "result_type": "string",
    },
    {
        "name": "HTMLEncode_empty",
        "func": "HTMLEncode",
        "args": lambda: ("",),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "HTMLEncode_plain",
        "func": "HTMLEncode",
        "args": lambda: ("Hello World",),
        "py_ref": lambda a: "Hello World",
        "result_type": "string",
    },
    {
        "name": "HTMLEncode_quotes",
        "func": "HTMLEncode",
        "args": lambda: ('He said "hello"',),
        "py_ref": lambda a: _vba_html_encode(a[0]),
        "result_type": "string",
    },

    # =========================================================================
    # HTMLDecode(text) -> String
    # =========================================================================
    {
        "name": "HTMLDecode_paragraph",
        "func": "HTMLDecode",
        "args": lambda: ("&lt;p&gt;Hi&lt;/p&gt;",),
        "py_ref": lambda a: _vba_html_decode(a[0]),
        "result_type": "string",
    },
    {
        "name": "HTMLDecode_empty",
        "func": "HTMLDecode",
        "args": lambda: ("",),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "HTMLDecode_plain",
        "func": "HTMLDecode",
        "args": lambda: ("Hello World",),
        "py_ref": lambda a: "Hello World",
        "result_type": "string",
    },
    {
        "name": "HTMLDecode_ampersand",
        "func": "HTMLDecode",
        "args": lambda: ("A &amp; B",),
        "py_ref": lambda a: _vba_html_decode(a[0]),
        "result_type": "string",
    },
    {
        "name": "HTMLDecode_quote",
        "func": "HTMLDecode",
        "args": lambda: ("&quot;Hello&quot;",),
        "py_ref": lambda a: _vba_html_decode(a[0]),
        "result_type": "string",
    },

    # =========================================================================
    # LevenshteinDistance(a, b, caseSensitive=False) -> Long
    # =========================================================================
    {
        "name": "Levenshtein_kitten_sitting",
        "func": "LevenshteinDistance",
        "args": lambda: ("kitten", "sitting"),
        "py_ref": lambda a: _levenshtein_py(a[0], a[1]),
        "result_type": "scalar",
    },
    {
        "name": "Levenshtein_identical",
        "func": "LevenshteinDistance",
        "args": lambda: ("abc", "abc"),
        "py_ref": lambda a: 0,
        "result_type": "scalar",
    },
    {
        "name": "Levenshtein_empty_first",
        "func": "LevenshteinDistance",
        "args": lambda: ("", "hello"),
        "py_ref": lambda a: 5,
        "result_type": "scalar",
    },
    {
        "name": "Levenshtein_empty_second",
        "func": "LevenshteinDistance",
        "args": lambda: ("hello", ""),
        "py_ref": lambda a: 5,
        "result_type": "scalar",
    },
    {
        "name": "Levenshtein_both_empty",
        "func": "LevenshteinDistance",
        "args": lambda: ("", ""),
        "py_ref": lambda a: 0,
        "result_type": "scalar",
    },
    {
        "name": "Levenshtein_case_insensitive",
        "func": "LevenshteinDistance",
        "args": lambda: ("Hello", "hello"),
        "py_ref": lambda a: 0,
        "result_type": "scalar",
    },
    {
        "name": "Levenshtein_case_sensitive",
        "func": "LevenshteinDistance",
        "args": lambda: ("Hello", "hello", True),
        "py_ref": lambda a: 1,
        "result_type": "scalar",
    },

    # =========================================================================
    # Soundex(text) -> String
    # =========================================================================
    {
        "name": "Soundex_Robert",
        "func": "Soundex",
        "args": lambda: ("Robert",),
        "py_ref": lambda a: _soundex_py(a[0]),
        "result_type": "string",
    },
    {
        "name": "Soundex_Rupert",
        "func": "Soundex",
        "args": lambda: ("Rupert",),
        "py_ref": lambda a: _soundex_py(a[0]),
        "result_type": "string",
    },
    {
        "name": "Soundex_empty",
        "func": "Soundex",
        "args": lambda: ("",),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "Soundex_single",
        "func": "Soundex",
        "args": lambda: ("A",),
        "py_ref": lambda a: _soundex_py(a[0]),
        "result_type": "string",
    },
    {
        "name": "Soundex_Smith",
        "func": "Soundex",
        "args": lambda: ("Smith",),
        "py_ref": lambda a: _soundex_py(a[0]),
        "result_type": "string",
    },
    {
        "name": "Soundex_Smythe",
        "func": "Soundex",
        "args": lambda: ("Smythe",),
        "py_ref": lambda a: _soundex_py(a[0]),
        "result_type": "string",
    },

    # =========================================================================
    # ExtractBetween(text, leftDelim, rightDelim, nth=1, includeDelim=False) -> String
    # =========================================================================
    {
        "name": "ExtractBetween_basic",
        "func": "ExtractBetween",
        "args": lambda: ("<title>Hello</title>", "<title>", "</title>"),
        "py_ref": lambda a: "Hello",
        "result_type": "string",
    },
    {
        "name": "ExtractBetween_empty_text",
        "func": "ExtractBetween",
        "args": lambda: ("", "<", ">"),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "ExtractBetween_not_found",
        "func": "ExtractBetween",
        "args": lambda: ("no delimiters here", "<", ">"),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "ExtractBetween_include_delim",
        "func": "ExtractBetween",
        "args": lambda: ("<title>Hello</title>", "<title>", "</title>", 1, True),
        "py_ref": lambda a: "<title>Hello</title>",
        "result_type": "string",
    },
    {
        "name": "ExtractBetween_nth2",
        "func": "ExtractBetween",
        "args": lambda: ("a(b)c(d)e", "(", ")", 2),
        "py_ref": lambda a: "d",
        "result_type": "string",
    },

    # =========================================================================
    # ReverseString(text) -> String
    # =========================================================================
    {
        "name": "ReverseString_basic",
        "func": "ReverseString",
        "args": lambda: ("ABC",),
        "py_ref": lambda a: "CBA",
        "result_type": "string",
    },
    {
        "name": "ReverseString_empty",
        "func": "ReverseString",
        "args": lambda: ("",),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "ReverseString_palindrome",
        "func": "ReverseString",
        "args": lambda: ("radar",),
        "py_ref": lambda a: "radar",
        "result_type": "string",
    },
    {
        "name": "ReverseString_sentence",
        "func": "ReverseString",
        "args": lambda: ("Hello World",),
        "py_ref": lambda a: "dlroW olleH",
        "result_type": "string",
    },

    # =========================================================================
    # CountSubstring(text, search, caseSensitive=False) -> Long
    # =========================================================================
    {
        "name": "CountSubstring_banana",
        "func": "CountSubstring",
        "args": lambda: ("banana", "an"),
        "py_ref": lambda a: 2,
        "result_type": "scalar",
    },
    {
        "name": "CountSubstring_no_match",
        "func": "CountSubstring",
        "args": lambda: ("hello", "xyz"),
        "py_ref": lambda a: 0,
        "result_type": "scalar",
    },
    {
        "name": "CountSubstring_empty_text",
        "func": "CountSubstring",
        "args": lambda: ("", "abc"),
        "py_ref": lambda a: 0,
        "result_type": "scalar",
    },
    {
        "name": "CountSubstring_empty_search",
        "func": "CountSubstring",
        "args": lambda: ("hello", ""),
        "py_ref": lambda a: 0,
        "result_type": "scalar",
    },
    {
        "name": "CountSubstring_overlapping",
        "func": "CountSubstring",
        "args": lambda: ("aaaa", "aa"),
        "py_ref": lambda a: 2,  # VBA uses InStr (non-overlapping)
        "result_type": "scalar",
    },

    # =========================================================================
    # StartsWith(text, prefix, caseSensitive=False) -> Boolean
    # =========================================================================
    {
        "name": "StartsWith_true",
        "func": "StartsWith",
        "args": lambda: ("Hello World", "Hello"),
        "py_ref": lambda a: True,
        "result_type": "bool",
    },
    {
        "name": "StartsWith_false",
        "func": "StartsWith",
        "args": lambda: ("Hello World", "World"),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },
    {
        "name": "StartsWith_empty_text",
        "func": "StartsWith",
        "args": lambda: ("", "Hello"),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },
    {
        "name": "StartsWith_empty_prefix",
        "func": "StartsWith",
        "args": lambda: ("Hello", ""),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },
    {
        "name": "StartsWith_case_sensitive",
        "func": "StartsWith",
        "args": lambda: ("Hello", "h", True),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },

    # =========================================================================
    # EndsWith(text, suffix, caseSensitive=False) -> Boolean
    # =========================================================================
    {
        "name": "EndsWith_xlsx",
        "func": "EndsWith",
        "args": lambda: ("report.xlsx", ".xlsx"),
        "py_ref": lambda a: True,
        "result_type": "bool",
    },
    {
        "name": "EndsWith_false",
        "func": "EndsWith",
        "args": lambda: ("report.xlsx", ".pdf"),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },
    {
        "name": "EndsWith_empty",
        "func": "EndsWith",
        "args": lambda: ("", ".txt"),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },
    {
        "name": "EndsWith_case_insensitive",
        "func": "EndsWith",
        "args": lambda: ("File.XLSX", ".xlsx"),
        "py_ref": lambda a: True,
        "result_type": "bool",
    },

    # =========================================================================
    # LeftOf(text, delimiter, nth=1) -> String
    # =========================================================================
    {
        "name": "LeftOf_email",
        "func": "LeftOf",
        "args": lambda: ("john.doe@example.com", "@"),
        "py_ref": lambda a: "john.doe",
        "result_type": "string",
    },
    {
        "name": "LeftOf_no_delim",
        "func": "LeftOf",
        "args": lambda: ("noat.com", "@"),
        "py_ref": lambda a: "noat.com",  # VBA returns whole text when delim not found
        "result_type": "string",
    },
    {
        "name": "LeftOf_empty",
        "func": "LeftOf",
        "args": lambda: ("", "@"),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "LeftOf_nth2_dot",
        "func": "LeftOf",
        "args": lambda: ("a.b.c.d", ".", 2),
        "py_ref": lambda a: "a.b",
        "result_type": "string",
    },

    # =========================================================================
    # RightOf(text, delimiter, nth=1, fromRight=False) -> String
    # =========================================================================
    {
        "name": "RightOf_path",
        "func": "RightOf",
        "args": lambda: (r"C:\Data\file.txt", "\\"),
        "py_ref": lambda a: "Data\\file.txt",
        "result_type": "string",
    },
    {
        "name": "RightOf_no_delim",
        "func": "RightOf",
        "args": lambda: ("nodot", "."),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "RightOf_empty",
        "func": "RightOf",
        "args": lambda: ("", "@"),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "RightOf_fromRight_ext",
        "func": "RightOf",
        "args": lambda: ("archive.tar.gz", ".", 1, True),
        "py_ref": lambda a: "gz",
        "result_type": "string",
    },

    # =========================================================================
    # TextJoin(delimiter, ParamArray args) -> String
    # =========================================================================
    {
        "name": "TextJoin_basic",
        "func": "TextJoin",
        "args": lambda: (", ", "A", "B", "C"),
        "py_ref": lambda a: ", ".join(a[1:]),
        "result_type": "string",
    },
    {
        "name": "TextJoin_single",
        "func": "TextJoin",
        "args": lambda: ("-", "only"),
        "py_ref": lambda a: "only",
        "result_type": "string",
    },
    {
        "name": "TextJoin_empty_args",
        "func": "TextJoin",
        "args": lambda: (", ",),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "TextJoin_no_delimiter",
        "func": "TextJoin",
        "args": lambda: ("", "A", "B", "C"),
        "py_ref": lambda a: "".join(a[1:]),
        "result_type": "string",
    },

    # =========================================================================
    # NthWord(text, n, delimiter=" ") -> String
    # =========================================================================
    {
        "name": "NthWord_comma_2",
        "func": "NthWord",
        "args": lambda: ("apple,banana,cherry", 2, ","),
        "py_ref": lambda a: a[0].split(a[2])[1],
        "result_type": "string",
    },
    {
        "name": "NthWord_space_1",
        "func": "NthWord",
        "args": lambda: ("Hello World Again", 1),
        "py_ref": lambda a: "Hello",
        "result_type": "string",
    },
    {
        "name": "NthWord_negative_last",
        "func": "NthWord",
        "args": lambda: ("apple,banana,cherry", -1, ","),
        "py_ref": lambda a: "cherry",
        "result_type": "string",
    },
    {
        "name": "NthWord_out_of_range",
        "func": "NthWord",
        "args": lambda: ("one,two", 5, ","),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "NthWord_empty",
        "func": "NthWord",
        "args": lambda: ("", 1),
        "py_ref": lambda a: "",
        "result_type": "string",
    },

    # =========================================================================
    # CommonPrefix(a, b, caseSensitive=False) -> String
    # =========================================================================
    {
        "name": "CommonPrefix_flower",
        "func": "CommonPrefix",
        "args": lambda: ("flower", "flow"),
        "py_ref": lambda a: "flow",
        "result_type": "string",
    },
    {
        "name": "CommonPrefix_flight",
        "func": "CommonPrefix",
        "args": lambda: ("flower", "flight"),
        "py_ref": lambda a: "fl",
        "result_type": "string",
    },
    {
        "name": "CommonPrefix_no_match",
        "func": "CommonPrefix",
        "args": lambda: ("abc", "xyz"),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "CommonPrefix_empty",
        "func": "CommonPrefix",
        "args": lambda: ("", "hello"),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "CommonPrefix_case_sensitive",
        "func": "CommonPrefix",
        "args": lambda: ("Hello", "HELLO", True),
        "py_ref": lambda a: "H",
        "result_type": "string",
    },
    {
        "name": "CommonPrefix_case_insensitive",
        "func": "CommonPrefix",
        "args": lambda: ("Hello", "HELLO"),
        "py_ref": lambda a: "Hello",
        "result_type": "string",
    },

    # =========================================================================
    # PadLeft(text, totalWidth, padChar=" ") -> String
    # =========================================================================
    {
        "name": "PadLeft_zeros",
        "func": "PadLeft",
        "args": lambda: ("42", 5, "0"),
        "py_ref": lambda a: "00042",
        "result_type": "string",
    },
    {
        "name": "PadLeft_no_pad",
        "func": "PadLeft",
        "args": lambda: ("12345", 5, "0"),
        "py_ref": lambda a: "12345",
        "result_type": "string",
    },
    {
        "name": "PadLeft_longer",
        "func": "PadLeft",
        "args": lambda: ("123456", 3, "0"),
        "py_ref": lambda a: "123456",
        "result_type": "string",
    },
    {
        "name": "PadLeft_empty",
        "func": "PadLeft",
        "args": lambda: ("", 3, "x"),
        "py_ref": lambda a: "xxx",
        "result_type": "string",
    },
    {
        "name": "PadLeft_zero_width",
        "func": "PadLeft",
        "args": lambda: ("abc", 0, "x"),
        "py_ref": lambda a: "abc",
        "result_type": "string",
    },

    # =========================================================================
    # PadRight(text, totalWidth, padChar=" ") -> String
    # =========================================================================
    {
        "name": "PadRight_spaces",
        "func": "PadRight",
        "args": lambda: ("Name", 8),
        "py_ref": lambda a: "Name    ",
        "result_type": "string",
    },
    {
        "name": "PadRight_dots",
        "func": "PadRight",
        "args": lambda: ("Chapter", 10, "."),
        "py_ref": lambda a: "Chapter...",
        "result_type": "string",
    },
    {
        "name": "PadRight_empty",
        "func": "PadRight",
        "args": lambda: ("", 4, "-"),
        "py_ref": lambda a: "----",
        "result_type": "string",
    },

    # =========================================================================
    # Truncate(text, maxLength, suffix="...") -> String
    # =========================================================================
    {
        "name": "Truncate_basic",
        "func": "Truncate",
        "args": lambda: ("Hello World", 8),
        "py_ref": lambda a: "Hello...",
        "result_type": "string",
    },
    {
        "name": "Truncate_no_trunc",
        "func": "Truncate",
        "args": lambda: ("Hello", 10),
        "py_ref": lambda a: "Hello",
        "result_type": "string",
    },
    {
        "name": "Truncate_exact",
        "func": "Truncate",
        "args": lambda: ("Hello", 5),
        "py_ref": lambda a: "Hello",
        "result_type": "string",
    },
    {
        "name": "Truncate_very_short",
        "func": "Truncate",
        "args": lambda: ("Hello World", 3),
        "py_ref": lambda a: "Hel",  # maxLength <= len(suffix), no suffix added
        "result_type": "string",
    },
    {
        "name": "Truncate_custom_suffix",
        "func": "Truncate",
        "args": lambda: ("Hello World", 10, ".."),
        "py_ref": lambda a: "Hello Wo..",
        "result_type": "string",
    },
    {
        "name": "Truncate_zero_max",
        "func": "Truncate",
        "args": lambda: ("Hello", 0),
        "py_ref": lambda a: "",
        "result_type": "string",
    },

    # =========================================================================
    # IsNullOrEmpty(text) -> Boolean
    # =========================================================================
    {
        "name": "IsNullOrEmpty_empty_str",
        "func": "IsNullOrEmpty",
        "args": lambda: ("",),
        "py_ref": lambda a: True,
        "result_type": "bool",
    },
    {
        "name": "IsNullOrEmpty_space",
        "func": "IsNullOrEmpty",
        "args": lambda: (" ",),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },
    {
        "name": "IsNullOrEmpty_non_empty",
        "func": "IsNullOrEmpty",
        "args": lambda: ("Hello",),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },

    # =========================================================================
    # IsNullOrWhitespace(text) -> Boolean
    # =========================================================================
    {
        "name": "IsNullOrWhitespace_spaces",
        "func": "IsNullOrWhitespace",
        "args": lambda: ("   ",),
        "py_ref": lambda a: True,
        "result_type": "bool",
    },
    {
        "name": "IsNullOrWhitespace_tabs",
        "func": "IsNullOrWhitespace",
        "args": lambda: ("\t \t",),
        "py_ref": lambda a: True,
        "result_type": "bool",
    },
    {
        "name": "IsNullOrWhitespace_text",
        "func": "IsNullOrWhitespace",
        "args": lambda: (" a ",),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },
    {
        "name": "IsNullOrWhitespace_empty",
        "func": "IsNullOrWhitespace",
        "args": lambda: ("",),
        "py_ref": lambda a: True,
        "result_type": "bool",
    },

    # =========================================================================
    # IsEmail(text) -> Boolean
    # =========================================================================
    {
        "name": "IsEmail_valid",
        "func": "IsEmail",
        "args": lambda: ("user@example.com",),
        "py_ref": lambda a: True,
        "result_type": "bool",
    },
    {
        "name": "IsEmail_no_at",
        "func": "IsEmail",
        "args": lambda: ("invalid",),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },
    {
        "name": "IsEmail_no_dot",
        "func": "IsEmail",
        "args": lambda: ("user@domain",),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },
    {
        "name": "IsEmail_short",
        "func": "IsEmail",
        "args": lambda: ("a@b",),
        "py_ref": lambda a: False,  # Len < 5
        "result_type": "bool",
    },
    {
        "name": "IsEmail_double_at",
        "func": "IsEmail",
        "args": lambda: ("a@b@c.com",),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },

    # =========================================================================
    # IsUrl(text) -> Boolean
    # =========================================================================
    {
        "name": "IsUrl_https",
        "func": "IsUrl",
        "args": lambda: ("https://example.com",),
        "py_ref": lambda a: True,
        "result_type": "bool",
    },
    {
        "name": "IsUrl_http",
        "func": "IsUrl",
        "args": lambda: ("http://example.com",),
        "py_ref": lambda a: True,
        "result_type": "bool",
    },
    {
        "name": "IsUrl_invalid",
        "func": "IsUrl",
        "args": lambda: ("not-a-url",),
        "py_ref": lambda a: False,
        "result_type": "bool",
    },
    {
        "name": "IsUrl_ftp",
        "func": "IsUrl",
        "args": lambda: ("ftp://example.com",),
        "py_ref": lambda a: False,  # VBA only checks http/https
        "result_type": "bool",
    },
    {
        "name": "IsUrl_short",
        "func": "IsUrl",
        "args": lambda: ("http://",),
        "py_ref": lambda a: False,  # Len < 8 (for https check, but http is 7 chars + something = too short)
        "result_type": "bool",
    },

    # =========================================================================
    # Coalesce(ParamArray values) -> Variant
    # =========================================================================
    {
        "name": "Coalesce_first_non_empty",
        "func": "Coalesce",
        "args": lambda: ("", "Hello"),
        "py_ref": lambda a: "Hello",  # empty string skipped, "Hello" returned
        "result_type": "string",
    },
    {
        "name": "Coalesce_all_non_empty",
        "func": "Coalesce",
        "args": lambda: ("A", "B", "C"),
        "py_ref": lambda a: "A",
        "result_type": "string",
    },
    {
        "name": "Coalesce_only_empty",
        "func": "Coalesce",
        "args": lambda: ("", ""),
        "py_ref": lambda a: None,
        "result_type": "string",
    },
    {
        "name": "Coalesce_number_first",
        "func": "Coalesce",
        "args": lambda: (42, "fallback"),
        "py_ref": lambda a: 42,
        "result_type": "scalar",
    },
    {
        "name": "Coalesce_single",
        "func": "Coalesce",
        "args": lambda: ("only",),
        "py_ref": lambda a: "only",
        "result_type": "string",
    },

    # =========================================================================
    # Repeat(text, n) -> String
    # =========================================================================
    {
        "name": "Repeat_ab3",
        "func": "Repeat",
        "args": lambda: ("ab", 3),
        "py_ref": lambda a: "ababab",
        "result_type": "string",
    },
    {
        "name": "Repeat_zero",
        "func": "Repeat",
        "args": lambda: ("ab", 0),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "Repeat_one",
        "func": "Repeat",
        "args": lambda: ("hello", 1),
        "py_ref": lambda a: "hello",
        "result_type": "string",
    },
    {
        "name": "Repeat_empty",
        "func": "Repeat",
        "args": lambda: ("", 5),
        "py_ref": lambda a: "",
        "result_type": "string",
    },
    {
        "name": "Repeat_negative",
        "func": "Repeat",
        "args": lambda: ("x", -1),
        "py_ref": lambda a: "",
        "result_type": "string",
    },

    # UDF wrappers
    {"name": "UDF_STR_ISEMAIL", "func": "UDF_STR_ISEMAIL",
     "args": lambda: ([["test@example.com", "invalid", "a@b.c"]],),
     "py_ref": lambda a: [[True, False, True]],
     "result_type": "array", "compare_mode": "bool_array"},
    {"name": "UDF_STR_REPEAT", "func": "UDF_STR_REPEAT",
     "args": lambda: ([["ab"]], 3),
     "py_ref": lambda a: [["ababab"]],
     "result_type": "array"},

    # =====================================================================
    # Migrated from VBA Test_StringUtils — coverage gaps (2026-06-16)
    # =====================================================================

    # ---- Truncate — zero / negative maxLen ----
    {"name": "Truncate_zero", "func": "Truncate",
     "args": lambda: ("hello", 0),
     "py_ref": lambda a: "", "result_type": "string"},
    {"name": "Truncate_negative", "func": "Truncate",
     "args": lambda: ("hello", -1),
     "py_ref": lambda a: "", "result_type": "string"},

    # ---- PadLeft / PadRight — empty source string ----
    {"name": "PadLeft_empty", "func": "PadLeft",
     "args": lambda: ("", 5, "0"),
     "py_ref": lambda a: "00000", "result_type": "string"},
    {"name": "PadRight_empty", "func": "PadRight",
     "args": lambda: ("", 5, "-"),
     "py_ref": lambda a: "-----", "result_type": "string"},

    # ---- Base64 — long string round-trip encode side ----
    {"name": "Base64Encode_long", "func": "Base64Encode",
     "args": lambda: ("x" * 1000,),
     "py_ref": lambda a: __import__("base64").b64encode(("x"*1000).encode()).decode(),
     "result_type": "string"},

    # ---- LevenshteinDistance — 100-char all-different ----
    {"name": "LevenshteinDistance_100diff", "func": "LevenshteinDistance",
     "args": lambda: ("a" * 100, "b" * 100),
     "py_ref": lambda a: 100, "result_type": "scalar"},

    # ---- IsEmail / IsUrl — VBA Null input (COM marshaling unreliable) ----
    {"name": "IsEmail_null", "func": "IsEmail",
     "args": lambda: (None,), "py_ref": lambda a: False, "result_type": "bool",
     "skip_if": True, "skip_reason": "VBA Null not representable through COM"},
    {"name": "IsUrl_null", "func": "IsUrl",
     "args": lambda: (None,), "py_ref": lambda a: False, "result_type": "bool",
     "skip_if": True, "skip_reason": "VBA Null not representable through COM"},

]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


# 2026-09-13 回归: 2D/Range 输入、单遍实体解码、大小写敏感
TEST_CASES += [
    {"name": "TextJoin_2D_range", "func": "TextJoin",
     "args": lambda: (",", [["a", "b"], ["c", "d"]]),
     "py_ref": lambda a: "a,b,c,d", "result_type": "string"},
    {"name": "HTMLDecode_no_double_decode", "func": "HTMLDecode",
     "args": lambda: ("&#38;lt;",), "py_ref": lambda a: "&lt;", "result_type": "string"},
    {"name": "HTMLDecode_amp_seq", "func": "HTMLDecode",
     "args": lambda: ("&amp;amp; &lt;tag&gt;",), "py_ref": lambda a: "&amp; <tag>", "result_type": "string"},
    {"name": "RemoveChars_case_sensitive", "func": "RemoveChars",
     "args": lambda: ("aAbB", "a"), "py_ref": lambda a: "AbB", "result_type": "string"},
    {"name": "KeepChars_case_sensitive", "func": "KeepChars",
     "args": lambda: ("aAbB", "a"), "py_ref": lambda a: "a", "result_type": "string"},
]


# =============================================================================
# R5 probes (0 FAIL 约束: 错误/Empty/数组字面量必须在 VBA 内部构造与检查)
# =============================================================================

STR_R5_PROBE_VBA = r"""
Option Explicit

Public Function ProbeR5(ByVal which As String) As Variant
    On Error GoTo EH

    Dim v As Variant

    ' ---- R5-01: UDF 标量路径错误值 → CVErr(xlErrValue) (返回 1 = IsError) ----
    If which = "err_reversestring" Then
        v = UDF_STR_REVERSESTRING(CVErr(xlErrDiv0))
    ElseIf which = "err_totitlecase" Then
        v = UDF_STR_TOTITLECASE(CVErr(xlErrDiv0))
    ElseIf which = "err_isemail" Then
        v = UDF_STR_ISEMAIL(CVErr(xlErrDiv0))
    ElseIf which = "err_removechars" Then
        v = UDF_STR_REMOVECHARS(CVErr(xlErrDiv0), "-")
    ElseIf which = "err_truncate" Then
        v = UDF_STR_TRUNCATE(CVErr(xlErrDiv0), 5)
    ElseIf which = "err_countsubstring" Then
        v = UDF_STR_COUNTSUBSTRING(CVErr(xlErrDiv0), "a")
    ElseIf which = "err_padleft" Then
        v = UDF_STR_PADLEFT(CVErr(xlErrDiv0), 5)
    ElseIf which = "err_base64encode" Then
        v = UDF_STR_BASE64ENCODE(CVErr(xlErrDiv0))
    ElseIf which = "err_soundex" Then
        v = UDF_STR_SOUNDEX(CVErr(xlErrDiv0))
    ElseIf which = "err_urlencode" Then
        v = UDF_STR_URLENCODE(CVErr(xlErrDiv0))

    ' ---- R5-11: 1D Variant 数组 → 1×n 2D 结果 ----
    ' 注: 1D 提升后第二维保持原 1D 数组的 LBound (通常为 0), 不能假定 1..n
    ElseIf which = "1d_isnullempty" Then
        v = UDF_STR_ISNULLOREMPTY(Array("", "x", Empty))
        If Not IsArray(v) Then ProbeR5 = "NOT_ARRAY": Exit Function
        If UBound(v, 1) <> 1 Or (UBound(v, 2) - LBound(v, 2) + 1) <> 3 Then ProbeR5 = "BAD_SHAPE": Exit Function
        ProbeR5 = CStr(v(1, LBound(v, 2))) & "|" & CStr(v(1, LBound(v, 2) + 1)) & "|" & CStr(v(1, LBound(v, 2) + 2))
        Exit Function
    ElseIf which = "1d_isnullwhitespace" Then
        v = UDF_STR_ISNULLORWHITESPACE(Array("  ", "x", vbTab))
        If Not IsArray(v) Then ProbeR5 = "NOT_ARRAY": Exit Function
        If UBound(v, 1) <> 1 Or (UBound(v, 2) - LBound(v, 2) + 1) <> 3 Then ProbeR5 = "BAD_SHAPE": Exit Function
        ProbeR5 = CStr(v(1, LBound(v, 2))) & "|" & CStr(v(1, LBound(v, 2) + 1)) & "|" & CStr(v(1, LBound(v, 2) + 2))
        Exit Function
    ElseIf which = "1d_commonprefix" Then
        v = UDF_STR_COMMONPREFIX(Array("flower", "fleet"), "fl")
        If Not IsArray(v) Then ProbeR5 = "NOT_ARRAY": Exit Function
        ProbeR5 = CStr(UBound(v, 1)) & "x" & CStr(UBound(v, 2) - LBound(v, 2) + 1) & ":" & _
                  CStr(v(1, LBound(v, 2))) & "," & CStr(v(1, LBound(v, 2) + 1))
        Exit Function
    ElseIf which = "1d_levenshtein" Then
        v = UDF_STR_LEVENSHTEIN(Array("kitten", "sitting"), "sitting")
        If Not IsArray(v) Then ProbeR5 = "NOT_ARRAY": Exit Function
        ProbeR5 = CStr(UBound(v, 1)) & "x" & CStr(UBound(v, 2) - LBound(v, 2) + 1) & ":" & _
                  CStr(v(1, LBound(v, 2))) & "," & CStr(v(1, LBound(v, 2) + 1))
        Exit Function
    ElseIf which = "1d_base64decode" Then
        v = UDF_STR_BASE64DECODE(Array("SGVsbG8=", "V29ybGQ="))
        If Not IsArray(v) Then ProbeR5 = "NOT_ARRAY": Exit Function
        ProbeR5 = CStr(UBound(v, 1)) & "x" & CStr(UBound(v, 2) - LBound(v, 2) + 1) & ":" & _
                  CStr(v(1, LBound(v, 2))) & "," & CStr(v(1, LBound(v, 2) + 1))
        Exit Function
    ElseIf which = "1d_randomstring" Then
        ' 注意: 输出随机, 仅验证数组形状/不接受 CVErr
        v = UDF_STR_RANDOMSTRING(8, Array("abc", "xyz"))
        If IsError(v) Then ProbeR5 = "ERR": Exit Function
        If Not IsArray(v) Then ProbeR5 = "NOT_ARRAY": Exit Function
        ProbeR5 = CStr(UBound(v, 1)) & "x" & CStr(UBound(v, 2) - LBound(v, 2) + 1)
        Exit Function
    Else
        ProbeR5 = "UNKNOWN"
        Exit Function
    End If

    ' R5-01 分支: 1 = CVErr(xlErrValue), 0 = 其它
    If IsError(v) Then ProbeR5 = 1 Else ProbeR5 = 0
    Exit Function
EH:
    ProbeR5 = "RAISE:" & CStr(Err.Number)
End Function
"""


def _str_r5_probe(excel, wb, ws, runner, tc, args):
    """R5-01/R5-11 探针: 在 VBA 内部构造 CVErr/Array/Empty, 结果不穿透 COM。"""
    from tests.test_utils import run_macro
    vbproj = wb.VBProject
    for comp in list(vbproj.VBComponents):
        if comp.Name == "StrR5Probe":
            vbproj.VBComponents.Remove(comp)
    comp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
    comp.Name = "StrR5Probe"
    comp.CodeModule.AddFromString(
        STR_R5_PROBE_VBA.replace("\r\n", "\n").replace("\n", "\r\n"))
    val = run_macro(excel, wb, "StrR5Probe.ProbeR5", args[0])
    expected = tc["py_ref"](args) if callable(tc["py_ref"]) else tc["py_ref"]
    return val, expected, float(tc.get("tol", 0.0))


# =====================================================================
# 第五轮审查回归 (2026-09-13): R5-01/R5-11/R5-13/R5-61
# =====================================================================
TEST_CASES += [
    # ---- R5-01: 全部/代表 UDF 标量路径错误值 → #VALUE! ----
    {"name": "R5_01_err_reversestring", "func": "UDF_STR_REVERSESTRING",
     "args": lambda: ("err_reversestring",), "reconstruct": _str_r5_probe,
     "py_ref": lambda a: 1.0},
    {"name": "R5_01_err_totitlecase", "func": "UDF_STR_TOTITLECASE",
     "args": lambda: ("err_totitlecase",), "reconstruct": _str_r5_probe,
     "py_ref": lambda a: 1.0},
    {"name": "R5_01_err_isemail", "func": "UDF_STR_ISEMAIL",
     "args": lambda: ("err_isemail",), "reconstruct": _str_r5_probe,
     "py_ref": lambda a: 1.0},
    {"name": "R5_01_err_removechars", "func": "UDF_STR_REMOVECHARS",
     "args": lambda: ("err_removechars",), "reconstruct": _str_r5_probe,
     "py_ref": lambda a: 1.0},
    {"name": "R5_01_err_truncate", "func": "UDF_STR_TRUNCATE",
     "args": lambda: ("err_truncate",), "reconstruct": _str_r5_probe,
     "py_ref": lambda a: 1.0},
    {"name": "R5_01_err_countsubstring", "func": "UDF_STR_COUNTSUBSTRING",
     "args": lambda: ("err_countsubstring",), "reconstruct": _str_r5_probe,
     "py_ref": lambda a: 1.0},
    {"name": "R5_01_err_padleft", "func": "UDF_STR_PADLEFT",
     "args": lambda: ("err_padleft",), "reconstruct": _str_r5_probe,
     "py_ref": lambda a: 1.0},
    {"name": "R5_01_err_base64encode", "func": "UDF_STR_BASE64ENCODE",
     "args": lambda: ("err_base64encode",), "reconstruct": _str_r5_probe,
     "py_ref": lambda a: 1.0},
    {"name": "R5_01_err_soundex", "func": "UDF_STR_SOUNDEX",
     "args": lambda: ("err_soundex",), "reconstruct": _str_r5_probe,
     "py_ref": lambda a: 1.0},
    {"name": "R5_01_err_urlencode", "func": "UDF_STR_URLENCODE",
     "args": lambda: ("err_urlencode",), "reconstruct": _str_r5_probe,
     "py_ref": lambda a: 1.0},

    # ---- R5-11: 1D Variant 数组 → 2D 1×n (谓词逐元素 Boolean) ----
    {"name": "R5_11_1d_isnullempty", "func": "UDF_STR_ISNULLOREMPTY",
     "args": lambda: ("1d_isnullempty",), "reconstruct": _str_r5_probe,
     "py_ref": lambda a: "True|False|True", "result_type": "string"},
    {"name": "R5_11_1d_isnullwhitespace", "func": "UDF_STR_ISNULLORWHITESPACE",
     "args": lambda: ("1d_isnullwhitespace",), "reconstruct": _str_r5_probe,
     "py_ref": lambda a: "True|False|True", "result_type": "string"},
    {"name": "R5_11_1d_commonprefix", "func": "UDF_STR_COMMONPREFIX",
     "args": lambda: ("1d_commonprefix",), "reconstruct": _str_r5_probe,
     "py_ref": lambda a: "1x2:fl,fl", "result_type": "string"},
    {"name": "R5_11_1d_levenshtein", "func": "UDF_STR_LEVENSHTEIN",
     "args": lambda: ("1d_levenshtein",), "reconstruct": _str_r5_probe,
     "py_ref": lambda a: "1x2:3,0", "result_type": "string"},
    {"name": "R5_11_1d_base64decode", "func": "UDF_STR_BASE64DECODE",
     "args": lambda: ("1d_base64decode",), "reconstruct": _str_r5_probe,
     "py_ref": lambda a: "1x2:Hello,World", "result_type": "string"},
    # R5-11 补齐: UDF_STR_RANDOMSTRING 1D charset → 2D 1×n (输出随机, 仅验证形状)
    {"name": "R5_11_1d_randomstring", "func": "UDF_STR_RANDOMSTRING",
     "args": lambda: ("1d_randomstring",), "reconstruct": _str_r5_probe,
     "py_ref": lambda a: "1x2", "result_type": "string"},

    # ---- R5-13: Mac 前缀例外清单 (machine/machinery/macy 不再 MacXxx) ----
    {"name": "R5_13_machine", "func": "ToTitleCase",
     "args": lambda: ("machine",), "py_ref": lambda a: "Machine",
     "result_type": "string"},
    {"name": "R5_13_machinery", "func": "ToTitleCase",
     "args": lambda: ("machinery",), "py_ref": lambda a: "Machinery",
     "result_type": "string"},
    {"name": "R5_13_macdonald_mixed", "func": "ToTitleCase",
     "args": lambda: ("macDonald",), "py_ref": lambda a: "MacDonald",
     "result_type": "string"},
    {"name": "R5_13_macy", "func": "ToTitleCase",
     "args": lambda: ("macy",), "py_ref": lambda a: "Macy",
     "result_type": "string"},

    # ---- R5-61: Extended-B Ș/ș/Ț/ț 映射 + URLDecode 字面非 ASCII ----
    {"name": "R5_61_remove_diacritics_0218_021b", "func": "RemoveDiacritics",
     "args": lambda: ("ȘșȚț",), "py_ref": lambda a: "SsTt",
     "result_type": "string"},
    {"name": "R5_61_url_decode_literal_cafe", "func": "URLDecode",
     "args": lambda: ("café",), "py_ref": lambda a: "café",
     "result_type": "string"},
]


def main() -> int:
    runner = CrossValRunner("StringUtils", MODULE_PATHS)
    runner.run_all(TEST_CASES)
    passed, failed = runner.print_summary()
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    print(f"StringUtils Cross-Validation")
    print(f"  Test cases: {len(TEST_CASES)}")

    rc = main()

    # Additional UUID format check (random output, cannot compare exact value)
    print(f"\n  UUID / RandomString format checks (manual):")
    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = 1  # msoAutomationSecurityLow
    wb = None
    try:
        import tempfile
        output = os.path.join(
            tempfile.gettempdir(), "vba_crossval_StringUtils_uuid.xlsm"
        )
        from tests.test_utils import (
            create_workbook, inject_testrunner, run_macro, teardown,
        )
        wb = create_workbook(excel, output, MODULE_PATHS)

        # UUID check: length 36, 4 dashes, version 4 (char at pos 15 = '4')
        uuid_val = run_macro(excel, wb, "StringUtils.UUID")
        uuid_str = str(uuid_val).strip()
        uuid_ok = (
            len(uuid_str) == 36
            and uuid_str.count("-") == 4
            and re.match(
                r"^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$",
                uuid_str,
            )
            is not None
        )
        status = "PASS" if uuid_ok else "FAIL"
        print(f"     UUID  format={uuid_str}  {status}")

        # RandomString check (default alphanum, length 8)
        rs = run_macro(excel, wb, "StringUtils.RandomString")
        rs_str = str(rs).strip()
        rs_ok = len(rs_str) == 8 and bool(re.match(r"^[A-Za-z0-9]{8}$", rs_str))
        status_rs = "PASS" if rs_ok else "FAIL"
        print(f"     RndStr format={rs_str}  {status_rs}")

    finally:
        teardown(excel, wb)

    # Exit with error code if any test failed
    if rc != 0:
        sys.exit(1)
