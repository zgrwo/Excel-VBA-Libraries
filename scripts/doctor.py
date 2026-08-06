#!/usr/bin/env python3
"""
Environment readiness check for Excel-VBA-Libraries.

Verifies that the development environment has the required tools and
dependencies to run the project's test suite and build pipeline.

Usage:
    python scripts/doctor.py

Exit codes:
    0 - All checks passed
    1 - One or more checks failed (see stderr for details)
"""

import os
import sys
import subprocess


def check_python_version(min_version=(3, 8)):
    """Check that Python version meets minimum requirement."""
    current = sys.version_info[:2]
    if current < min_version:
        return False, (
            f"Python {current[0]}.{current[1]} found, "
            f"but >={'.'.join(map(str, min_version))} required"
        )
    return True, f"Python {current[0]}.{current[1]} OK"


def check_dependency(package_name, import_name=None):
    """Check that a Python package is importable."""
    import_name = import_name or package_name
    try:
        __import__(import_name)
        return True, f"{package_name} installed"
    except ImportError:
        return False, f"{package_name} NOT FOUND — run: pip install {package_name}"


def check_directory(path, label):
    """Check that a required directory exists."""
    full_path = os.path.join(REPO_ROOT, path)
    if os.path.isdir(full_path):
        return True, f"{label}/ exists"
    return False, f"{label}/ MISSING at {full_path}"


def check_file(path, label):
    """Check that a required file exists."""
    full_path = os.path.join(REPO_ROOT, path)
    if os.path.isfile(full_path):
        return True, f"{label} exists"
    return False, f"{label} MISSING at {full_path}"


def check_git():
    """Check that git is available and we're in a git repo."""
    try:
        result = subprocess.run(
            ['git', 'rev-parse', '--is-inside-work-tree'],
            capture_output=True, text=True, timeout=5
        )
        if result.returncode == 0:
            return True, "git repository OK"
        return False, "not a git repository"
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False, "git not found or not responding"


def main():
    global REPO_ROOT
    REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    os.chdir(REPO_ROOT)

    checks = [
        ("Python version", lambda: check_python_version()),
        ("Git", lambda: check_git()),
        ("openpyxl", lambda: check_dependency("openpyxl")),
        ("pywin32 (Excel COM)", lambda: check_dependency("pywin32", "win32com")),
        ("src/", lambda: check_directory("src", "src")),
        ("tests/", lambda: check_directory("tests", "tests")),
        ("VBA-Core/", lambda: check_directory("VBA-Core", "VBA-Core")),
        ("AGENTS.md", lambda: check_file("AGENTS.md", "AGENTS.md")),
    ]

    passed = 0
    failed = 0

    print("Excel-VBA-Libraries Environment Check")
    print("=" * 50)
    print()

    for label, check_fn in checks:
        ok, message = check_fn()
        status = "PASS" if ok else "FAIL"
        print(f"  [{status}] {message}")
        if ok:
            passed += 1
        else:
            failed += 1

    print()
    print(f"Result: {passed} passed, {failed} failed")

    if failed > 0:
        print()
        print("Fix the FAIL items above before running tests.")
        print("See AGENTS.md for the full build and test command table.")
        return 1

    print()
    print("Environment is ready. You can run:")
    print("  python tests/run_all_validation.py --quick")
    return 0


if __name__ == '__main__':
    sys.exit(main())
