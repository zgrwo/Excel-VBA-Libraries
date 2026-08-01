"""Validate anchor links in VBA_LIB_User_Manual.md."""
import re
import sys
from pathlib import Path


def _heading_to_anchor(heading, keep_cjk=False):
    """Convert a heading text to a markdown anchor.

    Args:
        heading: The heading text.
        keep_cjk: If True, preserve CJK/Unicode characters (GitHub style).
                   If False, strip them (VS Code style).
    """
    anchor = heading.lower()
    if keep_cjk:
        # GitHub: keep Unicode letters/digits, strip punctuation
        anchor = re.sub(r'[^\w\s\-]', '', anchor, flags=re.UNICODE)
    else:
        # VS Code: strip all non-ASCII characters
        anchor = re.sub(r'[^a-z0-9_ \-]', '', anchor)
    anchor = re.sub(r'\s+', '-', anchor)  # spaces→hyphens
    return anchor


def extract_headings(text):
    """Extract all valid anchor targets from markdown headings and <a id> tags.

    For combined headings like '#### ArrayMin / ArrayMax / ArraySum',
    each function name separated by '/' becomes an individual anchor.

    Generates both GitHub-style (CJK preserved) and VS Code-style
    (CJK stripped) anchors to handle both renderers.
    """
    anchors = set()
    # ATX headings: #### FuncName → funcname
    for m in re.finditer(r'^#{1,6}\s+(.+?)(?:\s*\{.*?\})?\s*$', text, re.MULTILINE):
        heading = m.group(1).strip()
        # Strip trailing '(Sub)' qualifier
        heading = re.sub(r'\s*\(Sub\)\s*$', '', heading).strip()
        # Split combined headings like 'ArrayMin / ArrayMax / ArraySum'
        parts = [p.strip() for p in heading.split('/')]
        for part in parts:
            part = part.strip()
            if part:
                # Generate both GitHub and VS Code anchor forms
                anchors.add(_heading_to_anchor(part, keep_cjk=True))
                anchors.add(_heading_to_anchor(part, keep_cjk=False))
    # <a id="..."> anchors
    for m in re.finditer(r'<a\s+id="([^"]+)"', text):
        anchors.add(m.group(1))
    return anchors


def extract_links(text):
    """Extract all [text](#anchor) links with line numbers."""
    links = []
    for i, line in enumerate(text.split('\n'), 1):
        for m in re.finditer(r'\[([^\]]*)\]\(#([^)]+)\)', line):
            links.append((m.group(1), m.group(2), i))
    return links


def main():
    manual = Path(__file__).parent.parent.parent / 'rules' / 'user-manual.md'
    if not manual.exists():
        print(f"ERROR: {manual} not found")
        sys.exit(2)

    text = manual.read_text(encoding='utf-8')
    anchors = extract_headings(text)
    links = extract_links(text)

    broken = 0
    for link_text, target, line in links:
        if target not in anchors:
            print(f"BROKEN [{link_text}](#{target}) at line {line}")
            broken += 1

    print(f"\nAnchors: {len(anchors)} | Links: {len(links)} | Broken: {broken}")
    if broken:
        print("FAIL: Broken links found")
        sys.exit(1)
    print("PASS: All links valid")
    sys.exit(0)


if __name__ == '__main__':
    main()
