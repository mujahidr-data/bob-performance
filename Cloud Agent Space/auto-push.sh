#!/bin/bash
# Auto-push script: Watches for file changes and automatically pushes to Apps Script and Git
# Usage: ./auto-push.sh [file1] [file2] ... or run without args to watch all files

set -e

echo "🔄 Auto-push enabled: Changes will be pushed to Apps Script and Git"
echo ""

# Function to push changes
push_changes() {
    local files="$@"
    
    echo "📝 Files changed: $files"
    echo ""
    
    # Check if Apps Script files changed
    if echo "$files" | grep -qE '\.(gs|js)$'; then
        echo "📤 Pushing to Apps Script..."
        clasp push || {
            echo "⚠️  Warning: Apps Script push failed"
        }
        echo ""
    fi
    
    # Git operations
    echo "📦 Committing and pushing to Git..."
    
    # Check if there are changes
    if git diff --quiet && git diff --cached --quiet; then
        echo "   No changes to commit"
    else
        git add -A
        git commit -m "Auto-update: $(date +'%Y-%m-%d %H:%M:%S')" || echo "   No changes to commit"
        
        # Push to git
        if git remote | grep -q "bob-performance"; then
            git push bob-performance main || git push bob-performance master || echo "   Git push skipped"
        elif git remote | grep -q "origin"; then
            git push origin main || git push origin master || echo "   Git push skipped"
        fi
    fi
    
    echo ""
    echo "✅ Auto-push complete!"
    echo ""
}

# If files are provided as arguments, push them
if [ $# -gt 0 ]; then
    push_changes "$@"
    exit 0
fi

# Otherwise, watch for file changes
echo "👀 Watching for file changes..."
echo "   Press Ctrl+C to stop"
echo ""

# Use fswatch if available, otherwise use inotifywait or fallback to polling
if command -v fswatch &> /dev/null; then
    # macOS
    fswatch -o . | while read f; do
        changed_files=$(git diff --name-only)
        if [ ! -z "$changed_files" ]; then
            push_changes "$changed_files"
        fi
    done
elif command -v inotifywait &> /dev/null; then
    # Linux
    inotifywait -m -r -e modify,create,delete . --format '%w%f' | while read file; do
        # Ignore git and node_modules
        if [[ "$file" != *".git"* ]] && [[ "$file" != *"node_modules"* ]]; then
            push_changes "$file"
        fi
    done
else
    echo "⚠️  File watcher not available. Install fswatch (macOS) or inotifywait (Linux)"
    echo "   Or use: npm run watch"
    exit 1
fi

