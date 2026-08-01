"""Generate CHANGELOG entries from git log.

Produces keepachangelog-formatted entries from commit messages.

Usage:
  python scripts/changelog.py                    # Show unreleased changes
  python scripts/changelog.py --since v1.0.0    # Changes since tag
  python scripts/changelog.py --all             # Full changelog
"""

import re
import sys
import subprocess
from datetime import date
from collections import defaultdict

# Commit type mapping (conventional commits)
TYPE_MAP = {
    "feat": "Added",
    "fix": "Fixed",
    "perf": "Changed",
    "refactor": "Changed",
    "docs": "Documentation",
    "test": "Testing",
    "chore": "Maintenance",
    "ci": "Maintenance",
    "build": "Maintenance",
}

RE_CONVENTIONAL = re.compile(
    r'^(feat|fix|perf|refactor|docs|test|chore|ci|build)'
    r'(?:\(([^)]+)\))?!?:\s*(.+)$',
    re.IGNORECASE
)


def get_git_log(since_tag=None):
    """Get git log entries."""
    cmd = ["git", "log", "--oneline", "--no-merges", "--format=%s"]
    if since_tag:
        cmd.append(f"{since_tag}..HEAD")
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, encoding='utf-8', errors='replace'
        )
        if result.returncode != 0:
            return []
        return [line.strip() for line in result.stdout.splitlines() if line.strip()]
    except (subprocess.SubprocessError, FileNotFoundError):
        return []


def categorize(commits):
    """Categorize commits into changelog sections."""
    sections = defaultdict(list)

    for msg in commits:
        m = RE_CONVENTIONAL.match(msg)
        if m:
            ctype = m.group(1).lower()
            scope = m.group(2) or ""
            description = m.group(3)
            section = TYPE_MAP.get(ctype, "Maintenance")
            prefix = f"**{scope}**: " if scope else ""
            sections[section].append(f"{prefix}{description}")
        else:
            sections["Maintenance"].append(msg)

    return sections


def format_changelog(sections, version="Unreleased"):
    """Format sections into keepachangelog markdown."""
    lines = [f"## [{version}] - {date.today().isoformat()}", ""]

    # Standard section order
    order = ["Added", "Changed", "Fixed", "Documentation", "Testing", "Maintenance"]
    for section in order:
        if section in sections:
            lines.append(f"### {section}")
            lines.append("")
            for item in sections[section]:
                lines.append(f"- {item}")
            lines.append("")

    return "\n".join(lines)


def main():
    since_tag = None
    show_all = "--all" in sys.argv

    if "--since" in sys.argv:
        idx = sys.argv.index("--since")
        if idx + 1 < len(sys.argv):
            since_tag = sys.argv[idx + 1]

    commits = get_git_log(since_tag if not show_all else None)

    if not commits:
        print("No commits found (or git not available).")
        print("Ensure you're in a git repository with commits.")
        return

    sections = categorize(commits)
    version = "Unreleased" if not since_tag else since_tag.lstrip("v")
    output = format_changelog(sections, version)
    print(output)
    print(f"\n---\nTotal commits: {len(commits)}")


if __name__ == "__main__":
    main()
