#!/bin/bash
# Install git hooks from scripts/ into .git/hooks/
#
# Usage: bash scripts/install_hooks.sh
#
# Currently installs:
#   pre-commit — check new files against agents.md project structure

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
HOOKS_DIR="$REPO_ROOT/.git/hooks"

echo "Installing git hooks..."

# pre-commit
cp "$SCRIPT_DIR/pre-commit" "$HOOKS_DIR/pre-commit"
chmod +x "$HOOKS_DIR/pre-commit"
echo "  ✓ pre-commit"

echo "Done. Hooks installed to .git/hooks/"
