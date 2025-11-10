#!/bin/bash
# Deployment script for Facets Cloud Platform

echo "=========================================="
echo "Facets Deployment Preparation"
echo "=========================================="
echo ""

# Check if service_account.json exists
if [ ! -f "service_account.json" ]; then
    echo "⚠️  Warning: service_account.json not found"
    echo "   Make sure to set SERVICE_ACCOUNT_JSON environment variable in Facets"
    echo ""
fi

# Check required files
echo "📋 Checking required files..."
required_files=(
    "facets_handler.py"
    "hibob_report_downloader.py"
    "facets_requirements.txt"
)

for file in "${required_files[@]}"; do
    if [ -f "$file" ]; then
        echo "   ✓ $file"
    else
        echo "   ❌ $file (missing)"
    fi
done

echo ""
echo "📦 Files to upload to Facets:"
echo "   1. facets_handler.py (entry point)"
echo "   2. hibob_report_downloader.py (main automation)"
echo "   3. facets_requirements.txt (dependencies)"
echo "   4. config.template.json (reference)"
echo ""
echo "🔐 Environment Variables to set in Facets:"
echo "   - HIBOB_EMAIL: Your HiBob email"
echo "   - HIBOB_PASSWORD: Your HiBob password"
echo "   - SERVICE_ACCOUNT_JSON: Base64-encoded service_account.json"
echo "   - GOOGLE_SHEET_ID: 1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA"
echo ""
echo "📖 Next steps:"
echo "   1. Log in to Facets dashboard"
echo "   2. Create new project/function"
echo "   3. Upload the files listed above"
echo "   4. Set environment variables"
echo "   5. Set handler to: facets_handler.handler"
echo "   6. Set timeout to: 30 minutes (browser automation takes time)"
echo "   7. Set memory to: 2-4 GB (Playwright needs memory)"
echo "   8. Deploy!"
echo ""
echo "💡 See FACETS_DEPLOYMENT.md for detailed instructions"

