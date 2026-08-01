#!/bin/bash
# Push script — validate, review, confirm, and push
#
# Usage: bash scripts/push.sh
#
# Workflow:
#   1. Check working tree (warn if uncommitted changes)
#   2. Run quick validation (Layer 1 static checks)
#   3. Show commits + files to be pushed
#   4. Confirm (y/N)
#   5. Push to origin/<current-branch>

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$REPO_ROOT"

BRANCH=$(git branch --show-current)
REMOTE="origin/$BRANCH"

echo "============================================"
echo "  Push to $REMOTE"
echo "============================================"
echo ""

# ── 0. Check working tree ────────────────────────────────────────
DIRTY=0
if ! git diff --quiet 2>/dev/null; then
    echo "  ⚠ Working tree has unstaged changes"
    DIRTY=1
fi
STAGED=$(git diff --cached --name-only 2>/dev/null | wc -l)
if [ "$STAGED" -gt 0 ]; then
    echo "  ⚠ $STAGED file(s) staged but not committed"
    DIRTY=1
fi
if [ "$DIRTY" -eq 1 ]; then
    echo "  (only committed changes will be pushed)"
    echo ""
fi

# ── 1. Quick Validation ──────────────────────────────────────────
echo "[1/4] Running quick validation..."
python tests/run_all_validation.py --quick
echo ""

# ── 2. Check what's ahead ────────────────────────────────────────
echo "[2/4] Checking commits ahead of $REMOTE..."
echo ""

AHEAD=$(git rev-list --count $REMOTE..HEAD 2>/dev/null || echo "0")

if [ "$AHEAD" -eq 0 ]; then
    echo "  Nothing to push (already up to date)."
    exit 0
fi

# ── 3. Show push preview ─────────────────────────────────────────
echo "[3/4] $AHEAD commit(s) to push:"
echo ""
git log --oneline $REMOTE..HEAD
echo ""

echo "  Files:"
git diff --name-status $REMOTE..HEAD | while read status file; do
    [ -z "$file" ] && continue
    case "$status" in
        A) echo "    + $file";;
        M) echo "    ~ $file";;
        D) echo "    - $file";;
        *) echo "    ? $file";;
    esac
done
echo ""

# ── 4. Confirm and push ──────────────────────────────────────────
read -p "[4/4] Push to $REMOTE? [y/N] " CONFIRM
if [ "$CONFIRM" != "y" ] && [ "$CONFIRM" != "Y" ]; then
    echo "Cancelled."
    exit 0
fi

echo ""
git push origin "$BRANCH"
echo ""
echo "Done."
