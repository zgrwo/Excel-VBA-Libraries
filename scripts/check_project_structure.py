#!/usr/bin/env python3
"""
Check staged files against the project structure defined in agents.md.

Called by .git/hooks/pre-commit. Reads the list of staged files from git
diff --cached, compares against the project structure tree in agents.md,
and exits non-zero if any added file is not defined.
"""

import os
import re
import sys
import subprocess

# Fix encoding on Windows terminals (GBK → UTF-8)
if sys.platform == 'win32':
    try:
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except AttributeError:
        pass  # Python < 3.7


def parse_project_structure(manifest_path):
    """Parse the ASCII tree from rules/project-structure.md.

    Returns (exact_paths, wildcard_dirs) where:
      - exact_paths: set of exact relative file paths (e.g. 'src/ArrayUtils.bas')
      - wildcard_dirs: set of directory prefixes that allow any file under them
        (e.g. 'tests/' means any path starting with 'tests/' is allowed)
    """
    with open(manifest_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Extract the code block after heading "项目结构" + "## 结构树"
    match = re.search(r'## 结构树.*?\n```\n(.*?)\n```', content, re.DOTALL)
    if not match:
        raise SystemExit(
            "ERROR: Cannot find '## 结构树' code block in rules/project-structure.md"
        )

    tree_text = match.group(1)
    return _parse_tree(tree_text)


def _parse_tree(tree_text):
    """Parse ASCII tree lines into allowed paths.

    Algorithm:
    - Each line with '──' is a tree entry.
    - Depth is tracked by counting '│' characters in the line prefix.
    - Entries ending with '/' are directories (pushed onto a depth stack).
    - Entries without '/' are files (added as exact paths).
    - Directories that have NO children (never served as parent) become wildcards.
    - Section header lines (no '──') and the root line are skipped.
    """
    exact_paths = set()
    all_dirs = {}  # full_path -> has_children (bool)

    # Stack: list of (depth, dir_path) for active nesting levels
    stack = [(0, '')]  # synthetic root at depth 0

    for line in tree_text.split('\n'):
        if '──' not in line:
            continue

        # Find the tree connector marker
        marker = None
        for m in ['├── ', '└── ']:
            if m in line:
                marker = m
                break
        if marker is None:
            continue

        # Depth = character position of ├/└ divided by 4
        # (each nesting level is 4 chars: "│   " or "    ")
        connector_char = '├' if '├' in line else '└'
        depth = line.index(connector_char) // 4

        # Extract name: text between marker and comment/end
        name_raw = line.split(marker, 1)[1]
        name = name_raw.split('#')[0].strip()
        if not name:
            continue

        # Pop stack to the parent level (parent has depth < current depth)
        while len(stack) > 1 and stack[-1][0] >= depth:
            stack.pop()

        parent_path = stack[-1][1]
        full_path = parent_path + name

        if name.endswith('/'):
            # Directory entry — track children later
            all_dirs[full_path] = False
            stack.append((depth, full_path))
            # Mark parent as having children
            if parent_path in all_dirs:
                all_dirs[parent_path] = True
        else:
            # File entry — exact match
            exact_paths.add(full_path)
            # Mark parent as having children
            if parent_path in all_dirs:
                all_dirs[parent_path] = True

    # Directories with no children become wildcards (e.g. tests/)
    wildcard_dirs = {
        dir_path for dir_path, has_children in all_dirs.items()
        if not has_children
    }

    return exact_paths, wildcard_dirs


def is_allowed(file_path, exact_paths, wildcard_dirs):
    """Check if a file path is allowed by the project structure."""
    if file_path in exact_paths:
        return True
    for wc in wildcard_dirs:
        if file_path.startswith(wc):
            return True
    return False


def get_staged_files():
    """Return list of (status, path) for staged files from git diff --cached."""
    result = subprocess.run(
        ['git', 'diff', '--cached', '--name-status'],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        raise SystemExit("ERROR: git diff --cached failed")

    files = []
    for line in result.stdout.strip().split('\n'):
        if not line:
            continue
        parts = line.split('\t')
        if len(parts) >= 2:
            files.append((parts[0], parts[1]))
    return files


def main():
    repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    manifest = os.path.join(repo_root, 'rules', 'project-structure.md')

    if not os.path.exists(manifest):
        raise SystemExit("ERROR: rules/project-structure.md not found at " + manifest)

    exact_paths, wildcard_dirs = parse_project_structure(manifest)

    staged = get_staged_files()
    if not staged:
        return 0  # nothing staged, allow commit

    violations = []
    for status, path in staged:
        # Check newly added (A) and renamed (R) files; skip modified (M) and deleted (D)
        if status not in ('A', 'R'):
            continue

        # Always allow project-structure.md itself and scripts/ (bootstrapping)
        if path in ('rules/project-structure.md',) or path.startswith('scripts/'):
            continue

        if not is_allowed(path, exact_paths, wildcard_dirs):
            violations.append(path)

    if violations:
        print("=" * 60, file=sys.stderr)
        print("  COMMIT REJECTED: 新增文件未在项目结构中定义", file=sys.stderr)
        print("=" * 60, file=sys.stderr)
        print(file=sys.stderr)
        print("以下文件不在 rules/project-structure.md 结构树中：", file=sys.stderr)
        for v in violations:
            print(f"  ✗ {v}", file=sys.stderr)
        print(file=sys.stderr)
        print("解决方法：", file=sys.stderr)
        print("  1. 在 rules/project-structure.md 结构树中添加该文件", file=sys.stderr)
        print("  2. 或从暂存区移除: git reset HEAD <file>", file=sys.stderr)
        print(file=sys.stderr)
        print("已定义路径:", file=sys.stderr)
        for p in sorted(exact_paths):
            print(f"  ✓ {p}", file=sys.stderr)
        for w in sorted(wildcard_dirs):
            print(f"  ✓ {w}*", file=sys.stderr)
        return 1

    return 0


if __name__ == '__main__':
    sys.exit(main())
