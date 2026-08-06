#!/usr/bin/env python3
"""Pre-commit gate: run all static checks before allowing a commit.

Usage:
    python scripts/pre_commit_check.py          # full check
    python scripts/pre_commit_check.py --quick  # Layer 1 only (<1s)

Exit codes:
    0 = all checks passed
    1 = at least one check failed
"""
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

CHECKS = [
    ("Static validation (4 layers)", [sys.executable, str(ROOT / "tests/run_all_validation.py")]),
    ("VBA lint (7 rules)", [sys.executable, str(ROOT / "scripts/vba_lint.py"), "--summary"]),
    ("Docs consistency", [sys.executable, str(ROOT / "tests/utils/validate_docs.py")]),
    ("Manual anchors", [sys.executable, str(ROOT / "tests/utils/validate_manual_anchors.py")]),
    ("API count header", [sys.executable, str(ROOT / "scripts/generate_counts.py"), "--check"]),
]

QUICK_CHECKS = [
    ("Static validation (Layer 1)", [sys.executable, str(ROOT / "tests/run_all_validation.py"), "--quick"]),
    ("VBA lint (7 rules)", [sys.executable, str(ROOT / "scripts/vba_lint.py"), "--summary"]),
    ("API count header", [sys.executable, str(ROOT / "scripts/generate_counts.py"), "--check"]),
]


def main():
    quick = "--quick" in sys.argv
    checks = QUICK_CHECKS if quick else CHECKS
    failed = []

    print("=" * 60)
    print("PRE-COMMIT GATE" + (" (quick mode)" if quick else ""))
    print("=" * 60)

    for name, cmd in checks:
        print(f"\n[RUN] {name} ...")
        result = subprocess.run(cmd, cwd=str(ROOT), capture_output=True, text=True)
        if result.returncode != 0:
            print(f"  [FAIL]")
            if result.stdout.strip():
                print(f"  {result.stdout.strip()[-500:]}")
            if result.stderr.strip():
                print(f"  {result.stderr.strip()[-300:]}")
            failed.append(name)
        else:
            print(f"  [PASS]")

    print("\n" + "=" * 60)
    if failed:
        print(f"BLOCKED: {len(failed)} check(s) failed: {', '.join(failed)}")
        print("Fix issues before committing.")
        sys.exit(1)
    else:
        print("ALL CHECKS PASSED — safe to commit.")
        sys.exit(0)


if __name__ == "__main__":
    main()
