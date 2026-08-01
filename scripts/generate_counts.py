"""Generate function counts from VBA source and update documentation.

Scans src/*.bas for Public Function/Sub declarations and produces:
  - Per-module counts
  - Total count
  - Updates the count header in rules/api-reference.md (between markers)

Usage:
  python scripts/generate_counts.py           # print counts
  python scripts/generate_counts.py --update  # update docs in-place
"""

import re
import sys
from pathlib import Path
from collections import OrderedDict

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"
API_DOC = ROOT / "rules" / "api-reference.md"

RE_PUBLIC = re.compile(
    r'^\s*Public\s+(?:Static\s+)?(?:Function|Sub)\s+(\w+)',
    re.IGNORECASE
)

# Markers in api-reference.md for auto-update
MARKER_START = "<!-- AUTO_COUNTS_START -->"
MARKER_END = "<!-- AUTO_COUNTS_END -->"


def count_module(filepath: Path) -> tuple:
    """Count Public Functions and Subs in a .bas file."""
    functions = []
    subs = []
    try:
        for line in filepath.read_text(encoding='utf-8', errors='replace').splitlines():
            m = RE_PUBLIC.match(line)
            if m:
                name = m.group(1)
                if re.search(r'\bFunction\b', line, re.IGNORECASE):
                    functions.append(name)
                else:
                    subs.append(name)
    except OSError:
        pass
    return functions, subs


def main():
    update_mode = "--update" in sys.argv

    if not SRC.exists():
        print(f"ERROR: {SRC} not found", file=sys.stderr)
        sys.exit(2)

    modules = OrderedDict()
    total_funcs = 0
    total_subs = 0

    for bas in sorted(SRC.glob("*.bas")):
        funcs, subs = count_module(bas)
        modules[bas.stem] = (len(funcs), len(subs))
        total_funcs += len(funcs)
        total_subs += len(subs)

    # Print report
    print("=" * 50)
    print("Function Counts — Excel-VBA-Libraries")
    print("=" * 50)
    print(f"\n{'Module':<25} {'Functions':>10} {'Subs':>6} {'Total':>7}")
    print("-" * 50)
    for name, (fc, sc) in modules.items():
        print(f"{name:<25} {fc:>10} {sc:>6} {fc+sc:>7}")
    print("-" * 50)
    print(f"{'TOTAL':<25} {total_funcs:>10} {total_subs:>6} {total_funcs+total_subs:>7}")
    print(f"\nModules: {len(modules)}")

    # Update mode: inject into api-reference.md
    if update_mode:
        if not API_DOC.exists():
            print(f"\nWARNING: {API_DOC} not found, skipping update", file=sys.stderr)
            return

        content = API_DOC.read_text(encoding='utf-8')
        counts_text = f"**{len(modules)} 模块 | {total_funcs} Public Functions | {total_subs} Public Subs | 共 {total_funcs+total_subs} 个 Public 接口**"

        if MARKER_START in content and MARKER_END in content:
            new_section = f"{MARKER_START}\n{counts_text}\n{MARKER_END}"
            content = re.sub(
                re.escape(MARKER_START) + r'.*?' + re.escape(MARKER_END),
                new_section,
                content,
                flags=re.DOTALL
            )
            API_DOC.write_text(content, encoding='utf-8')
            print(f"\n[OK] Updated {API_DOC.relative_to(ROOT)}")
        else:
            print(f"\nNOTE: Markers not found in {API_DOC.name}. Add these markers to enable auto-update:")
            print(f"  {MARKER_START}")
            print(f"  {MARKER_END}")


if __name__ == "__main__":
    main()
