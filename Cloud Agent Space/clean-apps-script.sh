#!/bin/bash
# Helper script to clean up Apps Script files
# Note: clasp doesn't have delete command, so this just ensures only consolidated file is pushed

echo "🧹 Cleaning up Apps Script files..."
echo ""

# Remove any duplicate .js files (Apps Script converts .gs to .js)
if [ -f "bob-consolidated.js" ]; then
    echo "Removing bob-consolidated.js (keeping .gs version)"
    rm -f bob-consolidated.js
fi

# Ensure only consolidated file exists
echo "📋 Current Apps Script files:"
ls -la *.gs 2>/dev/null || echo "No .gs files found"

echo ""
echo "📤 Pushing to Apps Script..."
npx @google/clasp push

echo ""
echo "✅ Done!"
echo ""
echo "⚠️  IMPORTANT: Clasp doesn't have a delete command."
echo "   You need to manually delete old files in Apps Script:"
echo ""
echo "   1. Open: https://script.google.com/home/projects/1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47"
echo "   2. Delete these files if they exist:"
echo "      - bob-salary-data.js (or .gs)"
echo "      - HiBobPerformanceMenu.gs"
echo "      - HiBobReportHandler.gs"
echo "      - UPDATE_EXISTING_ONOPEN.gs"
echo "   3. Keep only: bob-consolidated.gs (or .js) and appsscript.json"
echo ""

