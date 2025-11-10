#!/bin/bash
# Startup script for HiBob Performance Report Downloader Web Interface

echo "=========================================="
echo "HiBob Performance Report Downloader"
echo "Web Interface Startup"
echo "=========================================="
echo ""

# Get the directory where this script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# Check if Python 3 is available
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is not installed or not in PATH"
    echo "   Please install Python 3 and try again"
    exit 1
fi

echo "✓ Python 3 found: $(python3 --version)"
echo ""

# Check if required packages are installed
echo "📦 Checking dependencies..."
MISSING_DEPS=()

if ! python3 -c "import flask" 2>/dev/null; then
    MISSING_DEPS+=("flask")
fi

if ! python3 -c "import playwright" 2>/dev/null; then
    MISSING_DEPS+=("playwright")
fi

if ! python3 -c "import pandas" 2>/dev/null; then
    MISSING_DEPS+=("pandas")
fi

if ! python3 -c "import google.auth" 2>/dev/null; then
    MISSING_DEPS+=("google-auth")
fi

if [ ${#MISSING_DEPS[@]} -gt 0 ]; then
    echo "⚠️  Missing dependencies: ${MISSING_DEPS[*]}"
    echo ""
    read -p "Install missing dependencies? (y/n) " -n 1 -r
    echo ""
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        echo "📥 Installing dependencies..."
        pip3 install -r requirements.txt
        if [ $? -ne 0 ]; then
            echo "❌ Failed to install dependencies"
            exit 1
        fi
        echo "✓ Dependencies installed"
    else
        echo "❌ Cannot start without required dependencies"
        exit 1
    fi
else
    echo "✓ All dependencies installed"
fi

echo ""

# Check if Playwright browsers are installed
if ! python3 -c "from playwright.sync_api import sync_playwright; p = sync_playwright().start(); p.chromium.executable_path; p.stop()" 2>/dev/null; then
    echo "⚠️  Playwright browsers not installed"
    read -p "Install Playwright browsers? (y/n) " -n 1 -r
    echo ""
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        echo "📥 Installing Playwright browsers..."
        python3 -m playwright install chromium
        echo "✓ Playwright browsers installed"
    fi
    echo ""
fi

# Check if config.json exists
if [ ! -f "config.json" ]; then
    echo "⚠️  config.json not found"
    if [ -f "config.template.json" ]; then
        echo "   Creating config.json from template..."
        cp config.template.json config.json
        echo "✓ config.json created"
        echo ""
        echo "⚠️  IMPORTANT: Please edit config.json with your credentials:"
        echo "   - HiBob email"
        echo "   - HiBob password"
        echo ""
        read -p "Press Enter to continue (you can edit config.json later)..."
    else
        echo "❌ config.template.json not found"
        exit 1
    fi
fi

# Check if service_account.json exists (optional but recommended)
if [ ! -f "service_account.json" ]; then
    echo "⚠️  service_account.json not found"
    echo "   Google Sheets upload will use OAuth2 (requires browser)"
    echo "   For better automation, set up service_account.json (see GOOGLE_SHEETS_API_SETUP.md)"
    echo ""
fi

# Get port from command line or use default
PORT=${1:-5000}

# Check if port is already in use
if lsof -Pi :$PORT -sTCP:LISTEN -t >/dev/null 2>&1 ; then
    echo "⚠️  Port $PORT is already in use"
    read -p "Use a different port? (y/n) " -n 1 -r
    echo ""
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        PORT=$((PORT + 1))
        echo "Using port $PORT instead"
    else
        echo "❌ Cannot start - port $PORT is in use"
        exit 1
    fi
fi

echo ""
echo "🚀 Starting web interface..."
echo ""
echo "=========================================="
echo "  Web Interface Starting"
echo "=========================================="
echo ""
echo "📱 Open your browser and go to:"
echo "   http://localhost:$PORT"
echo "   or"
echo "   http://127.0.0.1:$PORT"
echo ""
echo "🛑 Press Ctrl+C to stop the server"
echo ""
echo "=========================================="
echo ""

# Try to open browser automatically (macOS)
if [[ "$OSTYPE" == "darwin"* ]]; then
    sleep 2
    open "http://localhost:$PORT" 2>/dev/null &
fi

# Start the web app
python3 web_app.py $PORT

