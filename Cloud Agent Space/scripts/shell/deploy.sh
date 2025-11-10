#!/bin/bash
# Auto-deploy script: Push to Google Apps Script and Git

set -e

# Get the directory where this script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
# Change to project root (parent of scripts/)
PROJECT_ROOT="$( cd "$SCRIPT_DIR/.." && pwd )"
cd "$PROJECT_ROOT"

echo "🚀 Deploying to Google Apps Script and Git..."
echo ""

# Push to Apps Script (from apps-script folder)
echo "📤 Pushing to Apps Script..."
cd "$PROJECT_ROOT/apps-script" || exit 1
clasp push || {
    echo "❌ Failed to push to Apps Script"
    exit 1
}
cd "$PROJECT_ROOT"

# Git operations
echo ""
echo "📦 Committing and pushing to Git..."

# Check if there are changes
if git diff --quiet && git diff --cached --quiet; then
    echo "   No changes to commit"
else
    git add -A
    git commit -m "Auto-update: $(date +'%Y-%m-%d %H:%M:%S')" || echo "   No changes to commit"
fi

# Push to git (check which remote to use)
if git remote | grep -q "bob-performance"; then
    git push bob-performance main || git push bob-performance master || echo "   Git push skipped"
elif git remote | grep -q "origin"; then
    git push origin main || git push origin master || echo "   Git push skipped"
else
    echo "   No git remote configured"
fi

echo ""
echo "✅ Deployment complete!"

