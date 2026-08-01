"""Cross-validate all VBA modules against Python reference values.

Usage: python tests/run_all_crossval.py [module_name]

Delegates to individual build_*.py modules. Each build_<Module>.py defines
TEST_CASES with VBA function calls and Python reference computations.

When *module_name* is given (without ``build_`` or ``.py``), only that
module's cross-validation runs.

Exit: 0 = all pass, 1 = at least one failure.
"""

import importlib
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

BUILD_MODULES = [
    ("VBA-Core",        "tests.crossval.build_vba_core"),
    ("ArrayUtils",      "tests.crossval.build_ArrayUtils"),
    ("StringUtils",     "tests.crossval.build_StringUtils"),
    ("StatsUtils",      "tests.crossval.build_StatsUtils"),
    ("DictSetUtils",    "tests.crossval.build_DictSetUtils"),
    ("DateTimeUtils",   "tests.crossval.build_DateTimeUtils"),
    ("RangeUtils",      "tests.crossval.build_RangeUtils"),
    ("RegexUtils",      "tests.crossval.build_RegexUtils"),
    ("JsonUtils",       "tests.crossval.build_JsonUtils"),
    ("XmlUtils",        "tests.crossval.build_XmlUtils"),
    ("LinearUtils",     "tests.crossval.build_LinearUtils"),
    ("FileSystemUtils", "tests.crossval.build_FileSystemUtils"),
    ("PivotUtils",      "tests.crossval.build_PivotUtils"),
    ("SqlUtils",        "tests.crossval.build_SqlUtils"),
    ("PhyChemUtils",    "tests.crossval.build_PhyChemUtils"),
    ("RegressUtils",    "tests.crossval.build_RegressUtils"),
    ("ManualExamples",  "tests.crossval.build_manual_examples"),
    ("UDF-Range-xlsm",  "tests.crossval.build_udf_range"),
]


def main() -> int:
    cli_args = [a for a in sys.argv[1:] if not a.startswith("--")]
    target = cli_args[0] if cli_args else None

    if target:
        # Filter to matching module
        target_lower = target.lower()
        modules = [(l, ip) for l, ip in BUILD_MODULES
                   if target_lower in l.lower()
                   or target_lower in ip.lower().replace("tests.crossval.build_", "")]
        if not modules:
            print(f"ERROR: No build module matching '{target}' found.")
            print(f"Available: {', '.join(l for l, _ in BUILD_MODULES)}")
            return 1
    else:
        modules = BUILD_MODULES

    print("=" * 60)
    print("Excel VBA Libraries — Cross-Validation vs Python")
    print("=" * 60)
    print(f"Modules to test: {len(modules)}")

    grand_pass = 0
    grand_fail = 0
    failed_modules = []

    for label, import_path in modules:
        print(f"\n{'─' * 60}")
        print(f"  Module: {label}")
        print(f"{'─' * 60}")
        try:
            mod = importlib.import_module(import_path)
            rc = mod.main()
            if rc == 0:
                grand_pass += 1
            else:
                grand_fail += 1
                failed_modules.append(label)
        except Exception as exc:
            print(f"  ERROR importing {import_path}: {exc}")
            grand_fail += 1
            failed_modules.append(f"{label} (import error: {exc})")

    # Grand summary
    total = grand_pass + grand_fail
    print(f"\n{'=' * 60}")
    print(f"  CROSS-VALIDATION GRAND SUMMARY")
    print(f"{'=' * 60}")
    print(f"  Modules tested : {total}")
    print(f"  ALL PASS       : {grand_pass}")
    print(f"  FAILURES       : {grand_fail}")
    if total > 0:
        print(f"  Pass rate      : {100.0 * grand_pass / total:.1f}%")
    print(f"{'=' * 60}")

    if failed_modules:
        print(f"\n  Failed modules ({len(failed_modules)}):")
        for m in failed_modules:
            print(f"    - {m}")

    return 0 if grand_fail == 0 else 1


def main_with_target(target: str | None) -> int:
    """Run cross-validation, optionally filtered to *target* module name."""
    # Save and restore sys.argv to avoid side effects
    old_argv = sys.argv[:]
    if target:
        sys.argv = [sys.argv[0], target]
    else:
        sys.argv = [sys.argv[0]]  # no filter
    try:
        return main()
    finally:
        sys.argv = old_argv


if __name__ == "__main__":
    sys.exit(main())
