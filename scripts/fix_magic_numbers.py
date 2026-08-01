#!/usr/bin/env python3
"""Replace inline magic numbers with named constants in StatsUtils.bas and RegressUtils.bas."""
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def fix_stats_utils():
    f = ROOT / "src" / "StatsUtils.bas"
    lines = f.read_text(encoding="utf-8").splitlines()
    new_lines = []

    for i, line in enumerate(lines):
        # Skip const definitions (first ~100 lines)
        if i < 100:
            new_lines.append(line)
            continue

        # Skip comment-only lines
        stripped = line.strip()
        if stripped.startswith("'"):
            new_lines.append(line)
            continue

        new_line = line
        # Replace inline tolerances
        new_line = re.sub(r"< 1E-15\b", "< TOL_STRICT", new_line)
        new_line = re.sub(r"<= 1E-12\b", "<= TOL_DEFAULT", new_line)
        new_line = re.sub(r"< 1E-12\b", "< TOL_DEFAULT", new_line)
        new_line = re.sub(r"< 1E-14\b", "< TOL_BISECT", new_line)

        # Replace local Const EPS/FPMIN with module-level references
        if "Const EPS As Double = 1E-16" in new_line:
            new_line = new_line.replace("Const EPS As Double = 1E-16", "' EPS -> module-level NUM_EPS")
        if "Const FPMIN As Double = 1E-300" in new_line:
            new_line = new_line.replace("Const FPMIN As Double = 1E-300", "' FPMIN -> module-level FPMIN")

        new_lines.append(new_line)

    f.write_text("\n".join(new_lines), encoding="utf-8")
    print(f"[OK] StatsUtils.bas updated")


def fix_regress_utils():
    f = ROOT / "src" / "RegressUtils.bas"
    content = f.read_text(encoding="utf-8")
    lines = content.splitlines()

    # Find where to insert constants (after Option Explicit + header)
    insert_idx = 0
    for i, line in enumerate(lines):
        if line.strip().lower() == "option explicit":
            insert_idx = i + 1
            break

    # Check if constants already exist
    if "NUM_EPS" in content:
        print("[SKIP] RegressUtils.bas (constants exist)")
        return

    # Find first Private Const or first function
    for i in range(insert_idx, len(lines)):
        if "Private Const" in lines[i] or re.match(r"^\s*(Public|Private)\s+(Function|Sub)", lines[i]):
            insert_idx = i
            break

    # Insert constants
    consts = [
        "",
        "' Numerical tolerance constants",
        "Private Const NUM_EPS    As Double = 2.22044604925031E-16  ' Machine epsilon",
        "Private Const RANK_TOL As Double = 1E-15                 ' Rank determination tolerance",
        "",
    ]
    lines = lines[:insert_idx] + consts + lines[insert_idx:]

    # Replace inline magic numbers
    new_lines = []
    for i, line in enumerate(lines):
        if i < insert_idx + len(consts):
            new_lines.append(line)
            continue

        stripped = line.strip()
        if stripped.startswith("'"):
            new_lines.append(line)
            continue

        new_line = line
        new_line = new_line.replace("2.22044604925031E-16", "NUM_EPS")
        new_line = re.sub(r"< 1E-15\b", "< RANK_TOL", new_line)
        new_line = re.sub(r"= 1E-15\b", "= RANK_TOL", new_line)
        new_lines.append(new_line)

    f.write_text("\n".join(new_lines), encoding="utf-8")
    print(f"[OK] RegressUtils.bas updated")


if __name__ == "__main__":
    fix_stats_utils()
    fix_regress_utils()
