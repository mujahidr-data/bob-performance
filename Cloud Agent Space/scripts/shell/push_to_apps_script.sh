#!/bin/bash

# Push HiBob Report Handler to Google Apps Script using clasp
# 
# This script helps you push the Apps Script code to your Google Apps Script project

set -e  # Exit on error

# Get the directory where this script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
# Change to project root (parent of scripts/)
PROJECT_ROOT="$( cd "$SCRIPT_DIR/.." && pwd )"
# Change to apps-script directory
cd "$PROJECT_ROOT/apps-script" || exit 1

echo "=========================================="
echo "  Push to Google Apps Script with clasp"
echo "=========================================="
echo ""

# Check if clasp is installed
if ! command -v clasp &> /dev/null; then
    echo "❌ clasp is not installed."
    echo ""
    echo "To install clasp, run:"
    echo "  npm install -g @google/clasp"
    echo ""
    echo "If you get permission errors, try:"
    echo "  sudo npm install -g @google/clasp"
    echo ""
    exit 1
fi

echo "✅ clasp is installed"
echo ""

# Check if user is logged in to clasp
if ! clasp login --status &> /dev/null; then
    echo "🔐 You need to login to clasp first."
    echo ""
    echo "Running: clasp login"
    echo ""
    clasp login
    echo ""
fi

echo "✅ You are logged in to clasp"
echo ""

# Check if .clasp.json exists
if [ ! -f ".clasp.json" ]; then
    echo "❌ .clasp.json not found. Creating it..."
    cat > .clasp.json << 'EOF'
{
  "scriptId": "1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47",
  "rootDir": "."
}
EOF
    echo "✅ Created .clasp.json"
fi

# Push the code
echo "📤 Pushing code to Apps Script project..."
echo "   Script ID: 1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47"
echo ""

clasp push

echo ""
echo "✅ Successfully pushed code to Apps Script!"
echo ""
echo "You can open the project with:"
echo "  clasp open"
echo ""
echo "Or visit directly:"
echo "  https://script.google.com/home/projects/1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47"
echo ""

