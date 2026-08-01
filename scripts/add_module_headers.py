#!/usr/bin/env python3
"""Insert standardized module documentation headers into all src/*.bas files.

Header format:
'==============================================================================
' Module:       <name>
' Purpose:      <one-line description>
' Layer:        <Data|Statistics|Text|DateTime|Excel|Science>
' Dependencies: <deps>
' Public:       <count> functions/subs
' Notes:        <optional>
'==============================================================================
"""
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"

MODULE_INFO = {
    "ArrayUtils": ("Array operations: sort, filter, slice, aggregate, lookup", "Data"),
    "DictSetUtils": ("Dictionary/Set operations: merge, intersect, difference, frequency", "Data"),
    "PivotUtils": ("Data reshaping: pivot, unpivot, group-by, cross-join", "Data"),
    "SqlUtils": ("SQL queries on Excel ranges via ADODB", "Data"),
    "LinearUtils": ("Linear algebra: SVD, QR, LU, Cholesky, PINV, eigenvalues", "Statistics"),
    "StatsUtils": ("Statistics: descriptive, inference, distribution, correlation", "Statistics"),
    "RegressUtils": ("Regression: OLS, WLS, Ridge, ANOVA, factor importance", "Statistics"),
    "StringUtils": ("String processing: encode, decode, distance, UUID, URL", "Text"),
    "RegexUtils": ("Regular expressions: match, replace, split, capture groups", "Text"),
    "JsonUtils": ("JSON: pure-VBA recursive descent parser and stringifier", "Text"),
    "XmlUtils": ("XML: MSXML2 XPath query and table conversion", "Text"),
    "DateTimeUtils": ("Date/Time: ISO week, workdays, age, Easter, timestamps", "DateTime"),
    "RangeUtils": ("Range export: HTML/JSON/MD/CSV, area operations, naming", "Excel"),
    "FileSystemUtils": ("File system: UTF-8 read/write, folder ops, drive info", "Excel"),
    "PhyChemUtils": ("Physical chemistry: molecular weight, unit conversion, gas laws", "Science"),
}

RE_PUBLIC = re.compile(r"^\s*Public\s+(?:Static\s+)?(?:Function|Sub)\s+(\w+)", re.IGNORECASE)


def main():
    for bas in sorted(SRC.glob("*.bas")):
        name = bas.stem
        purpose, layer = MODULE_INFO.get(name, ("", ""))
        lines = bas.read_text(encoding="utf-8", errors="replace").splitlines()

        # Count public functions
        pub_count = sum(1 for l in lines if RE_PUBLIC.match(l))

        # Determine dependencies
        deps = "VBA-Core (VariantKit, ArrayOps, DictProxy)"
        notes = ""
        if name == "RegressUtils":
            deps = "VBA-Core + LinearUtils + StatsUtils (import these first)"
            notes = "Requires LinearUtils and StatsUtils loaded before this module."
        elif name == "SqlUtils":
            notes = "Requires ADODB. 64-bit Office needs Access Database Engine 2016."
        elif name == "XmlUtils":
            notes = "Requires MSXML2 (Windows built-in). XXE protection enabled."
        elif name == "RegexUtils":
            notes = "Requires VBScript.RegExp (Windows built-in). No Unicode category support."

        # Build header
        header = [
            "'==============================================================================",
            f"' Module:       {name}",
            f"' Purpose:      {purpose}",
            f"' Layer:        {layer}",
            f"' Dependencies: {deps}",
            f"' Public:       {pub_count} functions/subs",
        ]
        if notes:
            header.append(f"' Notes:        {notes}")
        header.append("'==============================================================================")

        # Check if header already exists
        if any("Module:" in l for l in lines[:10]):
            print(f"SKIP {name} (header exists)")
            continue

        # Insert after Option Explicit line
        insert_idx = 0
        for i, l in enumerate(lines):
            if l.strip().lower() == "option explicit":
                insert_idx = i + 1
                break

        # Add blank line after header
        new_lines = lines[:insert_idx] + [""] + header + [""] + lines[insert_idx:]
        bas.write_text("\n".join(new_lines), encoding="utf-8")
        print(f"DONE {name} ({pub_count} public, layer={layer})")


if __name__ == "__main__":
    main()
