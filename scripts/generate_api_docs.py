"""Extract API signatures from VBA source code.

Scans src/*.bas for Public Function/Sub declarations and generates
a markdown API reference with full signatures.

Usage:
  python scripts/generate_api_docs.py              # print to stdout
  python scripts/generate_api_docs.py -o out.md    # write to file
"""

import re
import sys
from pathlib import Path
from datetime import date

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"

RE_PUBLIC_SIG = re.compile(
    r'^\s*Public\s+(?:Static\s+)?(Function|Sub)\s+(\w+)\s*\(([^)]*)\)(?:\s*As\s+(\w+))?',
    re.IGNORECASE
)

RE_COMMENT_LINE = re.compile(r"^\s*'\s*(.+)$")


def extract_signatures(filepath: Path):
    """Extract all Public signatures with preceding comment block."""
    lines = filepath.read_text(encoding='utf-8', errors='replace').splitlines()
    entries = []
    comment_buffer = []

    for line in lines:
        # Accumulate comments
        cm = RE_COMMENT_LINE.match(line)
        if cm:
            comment_buffer.append(cm.group(1).strip())
            continue

        m = RE_PUBLIC_SIG.match(line)
        if m:
            kind = m.group(1)  # Function or Sub
            name = m.group(2)
            params = m.group(3).strip()
            ret_type = m.group(4) or ""

            # Clean params
            if params:
                params = re.sub(r'\s+', ' ', params)

            # Build signature
            sig = f"Public {kind} {name}({params})"
            if ret_type:
                sig += f" As {ret_type}"

            # Get description from last comment block
            desc = ""
            if comment_buffer:
                # Use first non-empty comment line as description
                for c in comment_buffer:
                    if c and not c.startswith("---"):
                        desc = c
                        break

            entries.append({
                "name": name,
                "kind": kind,
                "signature": sig,
                "return_type": ret_type,
                "description": desc,
            })
            comment_buffer = []
        elif not line.strip():
            # Empty line doesn't clear comments (allow multi-line)
            pass
        elif not line.strip().startswith("'"):
            comment_buffer = []

    return entries


def main():
    output_file = None
    if "-o" in sys.argv:
        idx = sys.argv.index("-o")
        if idx + 1 < len(sys.argv):
            output_file = sys.argv[idx + 1]

    if not SRC.exists():
        print(f"ERROR: {SRC} not found", file=sys.stderr)
        sys.exit(2)

    all_modules = {}
    total = 0

    for bas in sorted(SRC.glob("*.bas")):
        entries = extract_signatures(bas)
        all_modules[bas.stem] = entries
        total += len(entries)

    # Generate markdown
    lines = [
        f"# API Reference — Auto-Generated",
        f"",
        f"> Generated: {date.today().isoformat()} | {len(all_modules)} modules | {total} Public interfaces",
        f"> Source: `scripts/generate_api_docs.py` — do not edit manually.",
        f"",
    ]

    for module_name, entries in all_modules.items():
        lines.append(f"## {module_name} ({len(entries)} functions)")
        lines.append("")
        lines.append("| Function | Signature | Description |")
        lines.append("|----------|-----------|-------------|")
        for e in entries:
            sig_short = f"`{e['name']}({e['signature'].split('(')[1]}`" if '(' in e['signature'] else f"`{e['name']}`"
            desc = e['description'][:60] if e['description'] else "—"
            lines.append(f"| {e['name']} | {sig_short} | {desc} |")
        lines.append("")

    content = "\n".join(lines)

    if output_file:
        Path(output_file).write_text(content, encoding='utf-8')
        print(f"✓ Written to {output_file} ({total} signatures)")
    else:
        print(content)


if __name__ == "__main__":
    main()
