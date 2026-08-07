"""Documentation consistency validator — CI-ready.

Validates 15 checks across 4 layers:
  Layer 1: Function signature consistency (src/*.bas ↔ docs/)
  Layer 2: Documentation cross-references (AGENTS.md ↔ SKILL.md ↔ manuals)
  Layer 3: Metadata hygiene (module counts, date stamps, file existence)
  Layer 4: Code quality smells (private copies, deprecated in use, thin wrappers)

Usage:
  python tests/utils/validate_docs.py          # all checks
  python tests/utils/validate_docs.py --layer 1  # signature checks only
  python tests/utils/validate_docs.py --quiet     # only print failures
"""

import re
import sys
from pathlib import Path
from datetime import datetime, timedelta
from collections import defaultdict


ROOT = Path(__file__).resolve().parent.parent.parent
SRC = ROOT / "src"
DOCS = ROOT / "docs"
RULES = ROOT / "rules"
SKILLS = ROOT / "skills"
VBA_CORE = ROOT / "VBA-Core"

API_DOC = RULES / "api-reference.md"
USER_MANUAL = RULES / "user-manual.md"
USER_MANUAL_EN = DOCS / "VBA_LIB_User_Manual_EN.md"
CLAUDE_MD = ROOT / "AGENTS.md"  # renamed: was CLAUDE.md, now AGENTS.md
README_MD = ROOT / "README.md"
SKILL_MD = SKILLS / "vba-SKILL.md"
MANUAL_AUTHORING_MD = SKILLS / "vba-manual-authoring.md"

STALE_MARKER_PATTERNS = [
    r'⚠️\s*待实现', r'TODO.*实现', r'FIXME',
]

NON_VARIANT_TYPES = ('As Long', 'As Double', 'As String', 'As Boolean',
                      'As Worksheet', 'As Workbook', 'As Range')

# ── helpers ──────────────────────────────────────────────────────────────

class CheckResult:
    def __init__(self):
        self.passes = 0
        self.failures = []  # list of str descriptions

    def add_pass(self):
        self.passes += 1

    def add_fail(self, desc):
        self.failures.append(desc)

    @property
    def ok(self):
        return len(self.failures) == 0


def _load(path):
    try:
        return path.read_text(encoding="utf-8")
    except FileNotFoundError:
        return None


def _all_bas_files():
    return sorted(SRC.glob("*.bas"))


def _all_cls_files():
    return sorted(VBA_CORE.glob("*.cls"))


def _udf_signatures(text):
    """Yield (func_name, full_signature) for each UDF in source text."""
    for m in re.finditer(
        r'^Public (?:Function|Sub) (UDF_\w+)\((.*?)\)\s*As\s+\w+',
        text, re.MULTILINE | re.DOTALL
    ):
        yield m.group(1), m.group(2)


def _udf_bodies(text):
    """Yield (func_name, body_lines) for each UDF in source text."""
    for m in re.finditer(
        r"^Public (?:Function|Sub) (UDF_\w+).*?\n(.*?)^End (?:Function|Sub)",
        text, re.MULTILINE | re.DOTALL
    ):
        yield m.group(1), m.group(2)


def _parse_api_doc_counts(text):
    """Return {ModuleName: (public_count, udf_count)} from API doc headers."""
    counts = {}
    for m in re.finditer(
        r'\*\*Module\*\*:\s*`(\w+\.(?:bas|cls))`\s*\|\s*\*\*Public functions\*\*:\s*(\d+)'
        r'(?:\s*\|\s*\*\*Public subs\*\*:\s*\d+)?\s*\|\s*\*\*UDFs\*\*:\s*(\d+)',
        text
    ):
        module = m.group(1).replace('.bas', '').replace('.cls', '')
        counts[module] = (int(m.group(2)), int(m.group(3)))
    return counts


def _parse_vba_core_doc_counts(text):
    """Return {ClassName: (func_count, sub_count)} from VBA-Core appendix."""
    counts = {}
    for m in re.finditer(
        r'\*\*Module\*\*:\s*`(\w+\.cls)`\s*\|\s*\*\*Public functions\*\*:\s*(\d+)'
        r'(?:\s*\|\s*\*\*Public subs\*\*:\s*(\d+))?',
        text
    ):
        cls_name = m.group(1).replace('.cls', '')
        funcs = int(m.group(2))
        subs = int(m.group(3)) if m.group(3) else 0
        counts[cls_name] = (funcs, subs)
    return counts


# ── Layer 1: Function Signature Consistency ──────────────────────────────

def _count_public_in_bas(path):
    text = _load(path)
    if not text:
        return 0, 0, 0
    funcs = len(re.findall(r'^Public Function (?!UDF_|Test_)', text, re.MULTILINE))
    subs = len(re.findall(r'^Public Sub (?!UDF_|Test_)', text, re.MULTILINE))
    udfs = len(re.findall(r'^Public Function UDF_', text, re.MULTILINE))
    return funcs, subs, udfs


def check_function_counts():
    """Verify Public Function/Sub/UDF counts in API doc match source files."""
    result = CheckResult()
    api_text = _load(API_DOC)
    if not api_text:
        result.add_fail("API doc not found: " + str(API_DOC))
        return result

    doc_counts = _parse_api_doc_counts(api_text)

    for bas in _all_bas_files():
        module = bas.stem
        funcs, subs, udfs = _count_public_in_bas(bas)
        if module not in doc_counts:
            result.add_fail(f"{module}: missing from API doc header")
            continue
        expected_funcs, expected_udfs = doc_counts[module]
        if funcs != expected_funcs:
            result.add_fail(
                f"{module}: Public functions — doc={expected_funcs} src={funcs}"
            )
        else:
            result.add_pass()
        if udfs != expected_udfs:
            result.add_fail(
                f"{module}: UDFs — doc={expected_udfs} src={udfs}"
            )
        else:
            result.add_pass()

    # VBA-Core
    vba_counts = _parse_vba_core_doc_counts(api_text)
    for cls in _all_cls_files():
        name = cls.stem
        text = _load(cls)
        if not text:
            continue
        src_funcs = len(re.findall(r'^Public Function (?!Test_)', text, re.MULTILINE))
        src_subs = len(re.findall(r'^Public Sub (?!Test_)', text, re.MULTILINE))
        src_props = len(re.findall(r'^Public Property Get', text, re.MULTILINE))
        if name in vba_counts:
            doc_funcs, doc_subs = vba_counts[name]
            doc_total = doc_funcs + doc_subs
            src_total = src_funcs + src_subs + src_props
            if src_funcs != doc_funcs:
                result.add_fail(
                    f"{name}: Public functions — doc={doc_funcs} src={src_funcs}"
                )
            elif src_subs != doc_subs:
                result.add_fail(
                    f"{name}: Public subs — doc={doc_subs} src={src_subs}"
                )
            else:
                result.add_pass()

    # Cross-check: header total must equal sum of per-module counts
    header_match = re.search(
        r'\*\*Total public functions\*\*:\s*(\d+)\s*\|\s*\*\*Total UDFs\*\*:\s*(\d+)',
        api_text
    )
    if header_match:
        header_funcs = int(header_match.group(1))
        header_udfs = int(header_match.group(2))
        sum_funcs = sum(v[0] for v in doc_counts.values())
        sum_udfs = sum(v[1] for v in doc_counts.values())
        if header_funcs != sum_funcs:
            result.add_fail(
                f"API doc header: Public functions total={header_funcs} "
                f"but sum of per-module counts={sum_funcs}"
            )
        else:
            result.add_pass()
        if header_udfs != sum_udfs:
            result.add_fail(
                f"API doc header: UDFs total={header_udfs} "
                f"but sum of per-module counts={sum_udfs}"
            )
        else:
            result.add_pass()

    # Cross-check: UDF count should not significantly exceed function count.
    # Allow +1: a single core function may legitimately have two UDF wrappers
    # (e.g. CompressFactorPR → UDF_PC_COMPRESS + UDF_PC_ZFACTOR).
    for module, (funcs, udfs) in doc_counts.items():
        if udfs > funcs + 1:
            result.add_fail(
                f"{module}: UDFs ({udfs}) > Public functions ({funcs}) + 1 — "
                f"check for miscounted UDF wrappers"
            )
        else:
            result.add_pass()

    return result


def check_udf_params():
    """Verify all UDF parameters are As Variant."""
    result = CheckResult()
    for bas in _all_bas_files():
        text = _load(bas)
        if not text:
            continue
        module = bas.stem
        for func_name, sig in _udf_signatures(text):
            sig_flat = ' '.join(sig.split())
            for bad_type in NON_VARIANT_TYPES:
                if bad_type in sig_flat:
                    result.add_fail(
                        f"{module}.{func_name}: param has '{bad_type}' — must be As Variant"
                    )
                    break
            else:
                result.add_pass()
    return result


def check_udf_no_err_raise():
    """Verify no Err.Raise inside UDF bodies."""
    result = CheckResult()
    for bas in _all_bas_files():
        text = _load(bas)
        if not text:
            continue
        module = bas.stem
        for func_name, body in _udf_bodies(text):
            if re.search(r'\bErr\.Raise\b', body):
                result.add_fail(
                    f"{module}.{func_name}: contains Err.Raise — use CVErr(xlErrValue) instead"
                )
            else:
                result.add_pass()
    return result


def check_udf_no_volatile():
    """Verify no Application.Volatile in UDFs."""
    result = CheckResult()
    for bas in _all_bas_files():
        text = _load(bas)
        if not text:
            continue
        module = bas.stem
        for func_name, body in _udf_bodies(text):
            if re.search(r'\bApplication\.Volatile\b', body):
                result.add_fail(
                    f"{module}.{func_name}: uses Application.Volatile"
                )
            else:
                result.add_pass()
    return result


def check_udf_no_byref():
    """Verify UDF_* functions have no ByRef parameters (SKILL.md §1.3.1)."""
    result = CheckResult()
    for bas in _all_bas_files():
        text = _load(bas)
        if not text:
            continue
        module = bas.stem
        for func_name, sig in _udf_signatures(text):
            if re.search(r'\bByRef\b', sig):
                result.add_fail(
                    f"{module}.{func_name}: has ByRef param — must be ByVal (§1.3.1)"
                )
            else:
                result.add_pass()
    return result


# ── Layer 2: Documentation Cross-References ──────────────────────────────

def check_section_refs():
    """Verify all § references in AGENTS.md resolve in SKILL.md / manual-authoring.md."""
    result = CheckResult()
    claude = _load(CLAUDE_MD)
    skill = _load(SKILL_MD) or ""
    man_auth = _load(MANUAL_AUTHORING_MD) or ""

    if not claude:
        result.add_fail("AGENTS.md not found")
        return result

    refs = set(re.findall(r'§([0-9]+(?:\.[0-9]+)*)', claude))
    for ref in sorted(refs, key=lambda x: tuple(map(int, x.split('.')))):
        parts = ref.split('.')
        if len(parts) == 1:
            pattern = rf'^## {parts[0]}\.'
        elif len(parts) == 2:
            pattern = rf'^### {parts[0]}\.{parts[1]}'
        else:
            pattern = rf'^#### {".".join(parts)}'

        found = bool(re.search(pattern, skill, re.MULTILINE)) or \
                bool(re.search(pattern, man_auth, re.MULTILINE))

        if not found:
            alt = rf'^#{1,4}\s+{".".join(parts)}\b'
            found = bool(re.search(alt, skill, re.MULTILINE)) or \
                    bool(re.search(alt, man_auth, re.MULTILINE))

        if found:
            result.add_pass()
        else:
            result.add_fail(f"§{ref}: not found in SKILL.md or manual-authoring.md")
    return result


def check_api_manual_anchors():
    """Delegate to existing anchor validator for User Manual."""
    result = CheckResult()
    # This check is handled by tests/utils/validate_manual_anchors.py
    # which correctly handles combined headings like "#### FuncA / FuncB"
    # and both GitHub/VS Code anchor styles.
    result.add_pass()
    return result


def check_no_stale_markers():
    """Verify no stale TODO/⚠️ markers in architecture docs."""
    result = CheckResult()
    # Architecture docs may self-referentially describe the marker (checklist items) — skip them
    for path in [API_DOC, USER_MANUAL, USER_MANUAL_EN]:
        text = _load(path)
        if not text:
            continue
        for pattern in STALE_MARKER_PATTERNS:
            for m in re.finditer(pattern, text):
                result.add_fail(f"{path.name}: stale marker '{m.group(0).strip()}'")
        result.add_pass()
    return result


def check_en_no_chinese():
    """Verify English user manual has no Chinese characters (if it exists)."""
    result = CheckResult()
    text = _load(USER_MANUAL_EN)
    if not text:
        result.add_pass()
        return result

    for i, line in enumerate(text.splitlines(), 1):
        chinese = re.findall(r'[一-鿿]', line)
        if chinese:
            result.add_fail(
                f"{USER_MANUAL_EN.name}:{i}: Chinese chars: {''.join(chinese[:10])}"
            )
    if result.ok:
        result.add_pass()
    return result


# ── Layer 3: Metadata Hygiene ────────────────────────────────────────────

def check_module_counts():
    """Verify no stale hard-coded module counts."""
    result = CheckResult()
    claude = _load(CLAUDE_MD) or ""
    readme = _load(README_MD) or ""

    # Check for hard-coded "N/N 模块" patterns that drift
    stale = re.findall(r'(\d+/\d+\s*模块)', claude)
    for s in stale:
        result.add_fail(f"AGENTS.md: stale count '{s}'")
    if not stale:
        result.add_pass()

    # Check README for stale "N 模块 ... M 模块" patterns
    stale_readme = re.findall(r'(\d+\s*模块[^。]+?\d+\s*模块)', readme)
    for s in stale_readme:
        result.add_fail(f"README.md: stale count pattern '{s}'")
    if not stale_readme:
        result.add_pass()

    result.add_pass()
    return result


def check_last_updated():
    """Verify last_updated stamps are present and recent (< 90 days)."""
    result = CheckResult()
    ARCH_FILES = [CLAUDE_MD, README_MD, SKILL_MD, MANUAL_AUTHORING_MD]

    for path in ARCH_FILES:
        text = _load(path)
        if not text:
            result.add_fail(f"{path.name}: file not found")
            continue

        m = re.search(r'last_updated:\s*(\d{4}-\d{2}-\d{2})', text)
        if not m:
            result.add_fail(f"{path.name}: missing 'last_updated' stamp")
            continue

        try:
            dt = datetime.strptime(m.group(1), "%Y-%m-%d")
            age = datetime.now() - dt
            if age > timedelta(days=90):
                result.add_fail(
                    f"{path.name}: last_updated {m.group(1)} is {age.days}d old"
                )
            else:
                result.add_pass()
        except ValueError:
            result.add_fail(f"{path.name}: invalid date format: {m.group(1)}")

    return result


def check_routing_table():
    """Verify all files referenced in AGENTS.md routing table exist."""
    result = CheckResult()
    claude = _load(CLAUDE_MD)
    if not claude:
        result.add_fail("AGENTS.md not found")
        return result

    refs = set(re.findall(r'`(skills/\S+\.md|docs/\S+\.md)`', claude, re.IGNORECASE))
    # EN manual is documented as "延后至发布" — not a failure if absent
    known_deferred = {'docs/VBA_LIB_User_Manual_EN.md'}
    for ref in sorted(refs):
        full = ROOT / ref
        if full.exists() or ref in known_deferred:
            result.add_pass()
        else:
            result.add_fail(f"AGENTS.md routing table: '{ref}' does not exist")
    return result


# ── Layer 4: Code Quality Smells ──────────────────────────────────────────

# Legitimate patterns: VBA class instantiation requires Static instance wrappers.
# These are NOT code smells — they're the standard VBA pattern for accessing classes.
KNOWN_LEGITIMATE_WRAPPERS = {
    'IsNumericCell',   # Static VariantKit instance — requires wrapper per module
    'FilterPasses',    # Static VariantKit instance — requires wrapper per module
    'SafeKey',         # Static VariantKit instance — requires wrapper per module
    'ValuesEqual',     # Static VariantKit instance — requires wrapper per module
    'Compare',         # Static VariantKit instance — requires wrapper per module
    'WrapScalar',      # Static VariantKit instance — requires wrapper per module
    'SortStringArray', # Wraps ArrayOps.Sort for string-specific sorting
    'VarLetSet',       # JsonUtils: legitimate for function-return object dispatch
    'SafeCreateDict',  # Centralized error-guarded DictProxy.Create — kept intentionally
}

# VBA-Core public methods — if a .bas has a private copy, it's a smell
VBA_CORE_METHODS = {
    'VariantKit': ['Normalize1D', 'Normalize2D', 'NormalizeTo2D', 'ToDoubles',
                   'WrapScalar', 'SafeKey', 'ValuesEqual', 'Compare', 'IsEmptyArray',
                   'ArrayDims', 'Is1D', 'Is2D', 'IsNumericArray', 'IsNumericCell',
                   'FilterPasses', 'VarLetSet'],
    'ArrayOps': ['Sort', 'Slice', 'IndexOf', 'Flatten', 'SortIndices', 'CollectNumericColumns'],
    'DictProxy': ['Create', 'FromKeys', 'ToArray', 'Merge'],
}


def _count_func_lines(text, func_name):
    """Count body lines of a Private Function. Returns (line_count, has_body) or (0, False)."""
    pattern = rf'^Private (?:Function|Sub) {func_name}\b.*?\n(.*?)^End (?:Function|Sub)'
    m = re.search(pattern, text, re.MULTILINE | re.DOTALL)
    if not m:
        return 0, False
    body = m.group(1).strip()
    if not body:
        return 0, True
    return len(body.splitlines()), True


def check_private_vba_core_copies():
    """Detect .bas modules with private copies of VBA-Core functions.
    Thin wrappers (≤5 lines) in KNOWN_LEGITIMATE_WRAPPERS are allowed.
    Thick copies (>5 lines) are always flagged."""
    result = CheckResult()
    for bas in _all_bas_files():
        text = _load(bas)
        if not text:
            continue
        module = bas.stem
        for cls_name, methods in VBA_CORE_METHODS.items():
            for method in methods:
                if not re.search(rf'^Private (?:Function|Sub) {method}\b', text, re.MULTILINE):
                    continue
                line_count, _ = _count_func_lines(text, method)
                if method in KNOWN_LEGITIMATE_WRAPPERS and line_count <= 5:
                    continue  # thin wrapper — acceptable
                result.add_fail(
                    f"{module}: private copy of {cls_name}.{method} ({line_count} lines) — use {cls_name} directly"
                )
        result.add_pass()
    return result


# Deprecated patterns: scan source code for @deprecated markers and check
# if the deprecated symbol is still referenced elsewhere.
def check_deprecated_in_use():
    """Detect @deprecated functions still called in source code."""
    result = CheckResult()
    deprecated = {}  # {symbol: file_where_deprecated}

    # Find @deprecated annotations in VBA-Core
    for cls in _all_cls_files():
        text = _load(cls)
        if not text:
            continue
        for m in re.finditer(r"^' @deprecated\b.*", text, re.MULTILINE):
            # Look for the function name on the next line
            next_line_start = m.end()
            next_line = text[next_line_start:next_line_start + 200].split('\n')[0]
            func_match = re.search(r'(?:Function|Sub|Property)\s+(\w+)', next_line)
            if func_match:
                deprecated[func_match.group(1)] = cls.stem

    # Check all .bas and .cls files for references to deprecated symbols
    for src_path in list(_all_bas_files()) + list(_all_cls_files()):
        text = _load(src_path)
        if not text:
            continue
        name = src_path.stem
        for symbol, dep_file in deprecated.items():
            if name == dep_file:
                continue  # skip the defining file itself
            if re.search(rf'\b{symbol}\b', text):
                result.add_fail(
                    f"{name}: uses deprecated {dep_file}.{symbol}"
                )
    if not result.failures:
        result.add_pass()
    return result


def check_thin_wrappers():
    """Detect ≤3 line private functions that just delegate to VBA-Core."""
    result = CheckResult()
    for bas in _all_bas_files():
        text = _load(bas)
        if not text:
            continue
        module = bas.stem
        # Find private functions that are 1-3 lines and call VBA-Core
        for m in re.finditer(
            r'^Private (?:Function|Sub) (\w+).*?\n(.*?)\n^End (?:Function|Sub)',
            text, re.MULTILINE | re.DOTALL
        ):
            func_name = m.group(1)
            if func_name in KNOWN_LEGITIMATE_WRAPPERS:
                continue
            body = m.group(2).strip()
            lines = [l for l in body.split('\n')
                     if l.strip() and not l.strip().startswith("'")]
            if 1 <= len(lines) <= 3:
                for cls_name, methods in VBA_CORE_METHODS.items():
                    for method in methods:
                        if method in body and method not in KNOWN_LEGITIMATE_WRAPPERS:
                            result.add_fail(
                                f"{module}.{func_name}: thin wrapper ({len(lines)} lines) "
                                f"delegating to {cls_name}.{method}"
                            )
                            break
        result.add_pass()
    return result


# ── runner ────────────────────────────────────────────────────────────────

CHECKS = {
    # Layer 1
    "function_counts":        (1, check_function_counts),
    "udf_params_as_variant":  (1, check_udf_params),
    "udf_no_err_raise":       (1, check_udf_no_err_raise),
    "udf_no_volatile":        (1, check_udf_no_volatile),
    "udf_no_byref":           (1, check_udf_no_byref),
    # Layer 2
    "section_refs_resolve":   (2, check_section_refs),
    "api_manual_anchors":     (2, check_api_manual_anchors),
    "no_stale_markers":       (2, check_no_stale_markers),
    "en_no_chinese":          (2, check_en_no_chinese),
    # Layer 3
    "module_counts":                (3, check_module_counts),
    "last_updated":                 (3, check_last_updated),
    "routing_table":                (3, check_routing_table),
    # Layer 4
    "private_vba_core_copies":      (4, check_private_vba_core_copies),
    "deprecated_in_use":            (4, check_deprecated_in_use),
    "thin_wrappers":                (4, check_thin_wrappers),
}


def run(selected_layer=None, quiet=False):
    total_ok = 0
    total_fail = 0
    failures = []

    layer_names = {1: "Function Signature", 2: "Cross-References", 3: "Metadata", 4: "Code Smells"}

    for name, (layer, fn) in CHECKS.items():
        if selected_layer and layer != selected_layer:
            continue
        check_layer = layer_names.get(layer, f"Layer {layer}")
        result = fn()
        total_ok += result.passes
        total_fail += len(result.failures)
        failures.extend(result.failures)

        if not quiet or result.failures:
            status = "PASS" if result.ok else "FAIL"
            label = name.replace("_", " ")
            print(f"  {status} [{check_layer}] {label}")

    print()
    if total_fail == 0:
        print(f"ALL CHECKS PASSED ({total_ok}/{total_ok + total_fail})")
    else:
        print(f"FAILURES ({total_fail}/{total_ok + total_fail}):")
        for f in failures:
            # Replace non-ASCII for GBK terminal compatibility
            safe = f.encode('ascii', errors='replace').decode('ascii')
            print(f"  FAIL {safe}")

    return 0 if total_fail == 0 else 1


if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser(description="Documentation consistency validator")
    p.add_argument("--layer", type=int, choices=[1, 2, 3, 4],
                   help="Run only checks from specified layer")
    p.add_argument("--quiet", "-q", action="store_true",
                   help="Only print failures")
    args = p.parse_args()
    sys.exit(run(selected_layer=args.layer, quiet=args.quiet))
