"""Unified documentation consistency validation entry point.

Runs all static checks (no Excel required) for pre-commit gate.
For full validation including runtime checks, see run_all_tests.py and run_all_crossval.py.

Usage:
  python tests/run_all_validation.py           # all checks
  python tests/run_all_validation.py --quick    # Layer 1 only (fastest)
"""

import sys
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
UTILS = ROOT / "tests" / "utils"


def run_step(name, cmd, capture=True):
    print(f"  [{name}] ", end="", flush=True)
    try:
        result = subprocess.run(cmd, cwd=str(ROOT), capture_output=capture, text=True)
        if result.returncode == 0:
            print("PASS")
            return True
        else:
            print("FAIL")
            if capture and result.stdout:
                for line in result.stdout.splitlines():
                    if "FAIL" in line or "fail" in line.lower() or "error" in line.lower():
                        print(f"    {line.strip()}")
            return False
    except FileNotFoundError:
        print("SKIP (not found)")
        return True


def main():
    quick = "--quick" in sys.argv

    print("=" * 54)
    print("Documentation Consistency Validation")
    print("=" * 54)
    print()

    all_ok = True

    print("--- Doc Consistency (validate_docs.py) ---")
    cmd = [sys.executable, str(UTILS / "validate_docs.py")]
    if quick:
        cmd.append("--layer")
        cmd.append("1")
    all_ok &= run_step("doc-consistency", cmd)
    print()

    if not quick:
        print("--- Anchor Validation (validate_manual_anchors.py) ---")
        all_ok &= run_step("manual-anchors",
                          [sys.executable, str(UTILS / "validate_manual_anchors.py")])
        print()

        print("--- Test Coverage Gap Analysis (analyze_test_gaps.py) ---")
        all_ok &= run_step("coverage-gaps",
                          [sys.executable, str(UTILS / "analyze_test_gaps.py"), "--ci"])
        print()

    print("=" * 54)
    if all_ok:
        print("ALL VALIDATION CHECKS PASSED")
    else:
        print("SOME CHECKS FAILED — review output above")
    print("=" * 54)
    print()
    print("For full runtime validation:")
    print("  python tests/run_all_tests.py       # VBA unit tests (requires Excel)")
    print("  python tests/run_all_crossval.py    # cross-validation (requires Excel + Python)")
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
