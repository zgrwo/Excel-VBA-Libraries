"""VBA Static Lint Checker — Excel-VBA-Libraries.

Detects common VBA pitfalls that historically caused repeated bugs:
  1. Err.Raise misspelling (XxxErr.Raise)
  2. IIf usage (should use If/Else)
  3. ReDim(1 To 0) crash (should use Erase)
  4. Debug.Assert in production code
  5. Private Const inside procedures (should be module-level)
  6. Internal Variant abuse (Private functions using As Variant)
  7. ReDim Preserve inside loop (O(n^2) anti-pattern)
  8. Missing Option Explicit (forces implicit variable declaration)

Usage:
  python scripts/vba_lint.py              # lint all src/*.bas + VBA-Core/*.cls
  python scripts/vba_lint.py --summary    # only show summary
  python scripts/vba_lint.py --json       # JSON output for CI integration

Exit codes:
  0 = zero warnings
  1 = warnings found
  2 = error (files not found, etc.)
"""

import re
import sys
import json
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import List

# Ensure stdout handles Unicode on Windows (GBK console)
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"
VBA_CORE = ROOT / "VBA-Core"

# ── Lint Rules ─────────────────────────────────────────────────────────────

@dataclass
class LintRule:
    id: str
    name: str
    pattern: str
    message: str
    severity: str = "warning"  # warning | info
    regex: re.Pattern = field(init=False, repr=False)

    def __post_init__(self):
        self.regex = re.compile(self.pattern, re.IGNORECASE)


RULES: List[LintRule] = [
    LintRule(
        id="E001",
        name="Err.Raise misspelling",
        pattern=r'\b\w+Err\.Raise',
        message="Should be 'Err.Raise' (no module prefix)",
        severity="warning",
    ),
    LintRule(
        id="E002",
        name="IIf usage",
        pattern=r'\bIIf\s*\(',
        message="Use If/Else instead of IIf (evaluates both branches)",
        severity="warning",
    ),
    LintRule(
        id="E003",
        name="ReDim(1 To 0)",
        pattern=r'ReDim\s+\w+\s*\(\s*1\s+To\s+0\s*\)',
        message="ReDim(1 To 0) crashes on empty arrays; use Erase or uninitialized Dim",
        severity="warning",
    ),
    LintRule(
        id="E004",
        name="Debug.Assert",
        pattern=r'\bDebug\.Assert\b',
        message="Use Err.Raise 5 for production error handling",
        severity="info",
    ),
    LintRule(
        id="E005",
        name="Bare On Error Resume Next",
        pattern=r'On\s+Error\s+Resume\s+Next',
        message="Ensure Err.Number is checked after Resume Next",
        severity="info",
    ),
]

# E006: Private Const inside procedure (multi-line check)
RE_PROC_START = re.compile(
    r'^\s*(Public\s+|Private\s+|Friend\s+)?(Static\s+)?(Function|Sub|Property\s+(Get|Let|Set))\s+',
    re.IGNORECASE
)
RE_PROC_END = re.compile(r'^\s*End\s+(Function|Sub|Property)\b', re.IGNORECASE)
RE_PRIVATE_CONST = re.compile(r'^\s*Private\s+Const\s+', re.IGNORECASE)

# E007: Internal Variant abuse (Private Function returning As Variant)
RE_PRIVATE_VARIANT = re.compile(
    r'^\s*Private\s+(?:Static\s+)?Function\s+\w+\([^)]*\)\s+As\s+Variant\b',
    re.IGNORECASE
)

# E008: ReDim Preserve inside loop (O(n^2) anti-pattern)
RE_LOOP_START = re.compile(r'^\s*(For\s|Do\s)', re.IGNORECASE)
RE_LOOP_END = re.compile(r'^\s*(Next|Loop)\b', re.IGNORECASE)
RE_REDIM_PRESERVE = re.compile(r'\bReDim\s+Preserve\b', re.IGNORECASE)

# E010: UDF parameters must be As Variant (red-line rule)
RE_UDF_START = re.compile(
    r'^\s*Public\s+(?:Static\s+)?Function\s+UDF_\w+\s*\(', re.IGNORECASE
)
RE_NON_VARIANT_PARAM = re.compile(
    r'\bAs\s+(?:Long|Double|String|Boolean|Range|Worksheet|Workbook)\b', re.IGNORECASE
)


@dataclass
class LintFinding:
    file: str
    line: int
    rule_id: str
    rule_name: str
    severity: str
    message: str
    text: str


def lint_file(filepath: Path) -> List[LintFinding]:
    """Run all lint rules on a single file."""
    findings = []
    try:
        lines = filepath.read_text(encoding='utf-8', errors='replace').splitlines()
    except OSError as e:
        print(f"ERROR: Cannot read {filepath}: {e}", file=sys.stderr)
        return findings

    rel = str(filepath.relative_to(ROOT))

    # E009: Option Explicit required (file-level check)
    has_option_explicit = any(
        re.match(r"^\s*Option\s+Explicit\s*$", line, re.IGNORECASE)
        for line in lines[:20]  # Check first 20 lines (header area)
    )
    if not has_option_explicit:
        findings.append(LintFinding(
            file=rel, line=1,
            rule_id="E009", rule_name="Missing Option Explicit",
            severity="warning",
            message="Add 'Option Explicit' to force variable declaration",
            text="(file header)",
        ))

    in_procedure = False
    loop_depth = 0

    for i, line in enumerate(lines, 1):
        # Skip comment-only lines for most rules
        stripped = line.strip()
        is_comment = stripped.startswith("'") or stripped.startswith("Rem ")

        # Regex-based rules (skip comments)
        if not is_comment:
            for rule in RULES:
                if rule.regex.search(line):
                    # E005: On Error Resume Next is OK if followed by Err check
                    # We report as info only
                    findings.append(LintFinding(
                        file=rel, line=i,
                        rule_id=rule.id, rule_name=rule.name,
                        severity=rule.severity,
                        message=rule.message,
                        text=stripped[:100],
                    ))

        # E006: Private Const inside procedure
        if RE_PROC_START.match(line):
            in_procedure = True
        elif RE_PROC_END.match(line):
            in_procedure = False
        elif in_procedure and RE_PRIVATE_CONST.match(line):
            findings.append(LintFinding(
                file=rel, line=i,
                rule_id="E006", rule_name="Private Const in procedure",
                severity="warning",
                message="Move Private Const to module level (VBA scope trap)",
                text=stripped[:100],
            ))

        # E007: Internal Variant abuse (return type only)
        if not is_comment and RE_PRIVATE_VARIANT.match(line):
            findings.append(LintFinding(
                file=rel, line=i,
                rule_id="E007", rule_name="Internal Variant return",
                severity="info",
                message="Private Function returns Variant; consider concrete return type",
                text=stripped[:100],
            ))

        # E008: ReDim Preserve inside loop (O(n^2) anti-pattern)
        if not is_comment:
            if RE_LOOP_START.match(line):
                loop_depth += 1
            elif RE_LOOP_END.match(line):
                loop_depth = max(0, loop_depth - 1)
            elif loop_depth > 0 and RE_REDIM_PRESERVE.search(line):
                findings.append(LintFinding(
                    file=rel, line=i,
                    rule_id="E008", rule_name="ReDim Preserve in loop",
                    severity="info",
                    message="ReDim Preserve in loop is O(n^2); consider pre-alloc or ArrayList pattern",
                    text=stripped[:100],
                ))

        # E010: UDF parameters must be As Variant (red-line rule)
        if not is_comment and RE_UDF_START.match(line):
            # Collect full signature (may span multiple lines via " _")
            sig_lines = [line]
            j = i  # 0-based index of next line
            while stripped.endswith(" _") and j < len(lines):
                sig_lines.append(lines[j])
                stripped = lines[j].strip()
                j += 1
            full_sig = " ".join(sig_lines)
            non_variant = RE_NON_VARIANT_PARAM.findall(full_sig)
            if non_variant:
                findings.append(LintFinding(
                    file=rel, line=i,
                    rule_id="E010", rule_name="UDF non-Variant param",
                    severity="warning",
                    message=f"UDF param uses '{non_variant[0]}' — must be As Variant (Range input causes #VALUE!)",
                    text=stripped[:100],
                ))

    return findings


def collect_vba_files() -> List[Path]:
    """Collect all .bas and .cls files."""
    files = []
    if SRC.exists():
        files.extend(sorted(SRC.glob("*.bas")))
    if VBA_CORE.exists():
        files.extend(sorted(VBA_CORE.glob("*.cls")))
    return files


def main():
    args = sys.argv[1:]
    summary_only = "--summary" in args
    json_output = "--json" in args

    files = collect_vba_files()
    if not files:
        print("ERROR: No .bas/.cls files found in src/ or VBA-Core/", file=sys.stderr)
        sys.exit(2)

    all_findings: List[LintFinding] = []
    for f in files:
        all_findings.extend(lint_file(f))

    # Separate warnings from info
    warnings = [f for f in all_findings if f.severity == "warning"]
    infos = [f for f in all_findings if f.severity == "info"]

    if json_output:
        output = {
            "files_scanned": len(files),
            "warnings": len(warnings),
            "info": len(infos),
            "findings": [asdict(f) for f in all_findings],
        }
        print(json.dumps(output, indent=2, ensure_ascii=False))
    elif summary_only:
        print(f"Files scanned: {len(files)}")
        print(f"Warnings: {len(warnings)}")
        print(f"Info: {len(infos)}")
        if warnings:
            print("\nWarnings by rule:")
            from collections import Counter
            counts = Counter(f.rule_id for f in warnings)
            # Build lookup from both RULES list and structural rules
            rule_names = {r.id: r.name for r in RULES}
            rule_names.update({
                "E006": "Private Const in procedure",
                "E009": "Missing Option Explicit",
                "E010": "UDF non-Variant param",
            })
            for rule_id, count in sorted(counts.items()):
                name = rule_names.get(rule_id, rule_id)
                print(f"  {rule_id} ({name}): {count}")
    else:
        # Full output
        print("=" * 60)
        print("VBA Static Lint — Excel-VBA-Libraries")
        print("=" * 60)
        print(f"\nScanning {len(files)} files...\n")

        if not all_findings:
            print("  ✓ Zero findings. All clean!")
        else:
            for f in all_findings:
                icon = "⚠️" if f.severity == "warning" else "ℹ️"
                print(f"  {icon} {f.file}:{f.line} [{f.rule_id}] {f.message}")
                print(f"     {f.text}")

        print(f"\n{'─' * 60}")
        print(f"  Files: {len(files)} | Warnings: {len(warnings)} | Info: {len(infos)}")
        print(f"{'─' * 60}")

    # Exit code
    sys.exit(1 if warnings else 0)


if __name__ == "__main__":
    main()
