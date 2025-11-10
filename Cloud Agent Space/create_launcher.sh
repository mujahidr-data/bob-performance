#!/bin/bash
# Script to create a macOS Automator app launcher

APP_NAME="Bob Web Interface"
APP_PATH="$HOME/Desktop/${APP_NAME}.app"
SCRIPT_PATH="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
START_SCRIPT="$SCRIPT_PATH/scripts/shell/start_web_app.sh"

echo "Creating macOS launcher app..."
echo ""

# Create the app bundle structure
mkdir -p "${APP_PATH}/Contents/MacOS"
mkdir -p "${APP_PATH}/Contents/Resources"

# Get absolute project root path
PROJECT_ROOT_ABS="$( cd "$SCRIPT_PATH" && pwd )"

# Create the executable script with embedded project path
cat > "${APP_PATH}/Contents/MacOS/${APP_NAME}" << EOF
#!/bin/bash
# Project root path (embedded at app creation time)
PROJECT_ROOT="$PROJECT_ROOT_ABS"

# Change to project root
cd "\$PROJECT_ROOT" || {
    echo "❌ Error: Could not change to project directory: \$PROJECT_ROOT"
    echo ""
    echo "The project directory may have been moved."
    echo "Please recreate the launcher app or run the script manually:"
    echo "  cd \"\$PROJECT_ROOT\""
    echo "  ./scripts/shell/start_web_app.sh"
    echo ""
    echo "Press Enter to close..."
    read
    exit 1
}

# Check if start script exists
if [ ! -f "\$PROJECT_ROOT/scripts/shell/start_web_app.sh" ]; then
    echo "❌ Error: Start script not found: \$PROJECT_ROOT/scripts/shell/start_web_app.sh"
    echo ""
    echo "The project structure may have changed."
    echo "Press Enter to close..."
    read
    exit 1
fi

# Make script executable
chmod +x "\$PROJECT_ROOT/scripts/shell/start_web_app.sh"

# Run the start script
"\$PROJECT_ROOT/scripts/shell/start_web_app.sh" || {
    echo ""
    echo "❌ Error: Web app failed to start"
    echo "Press Enter to close..."
    read
    exit 1
}
EOF

chmod +x "${APP_PATH}/Contents/MacOS/${APP_NAME}"

# Create or find icon (use absolute path from project root)
# SCRIPT_PATH is the directory containing this script (project root)
PROJECT_ROOT="$SCRIPT_PATH"
ICON_PATH="$PROJECT_ROOT/assets/bob_icon.icns"
ICON_SCRIPT="$SCRIPT_PATH/create_icon.py"
if [ ! -f "$ICON_PATH" ]; then
    echo "📦 Creating icon..."
    if [ -f "$ICON_SCRIPT" ]; then
        python3 "$ICON_SCRIPT" 2>&1 | grep -v "Creating\|Created\|Icon" || true
    fi
    if [ -f "$ICON_PATH" ]; then
        echo "✅ Icon created"
    fi
fi

# Copy icon to app if it exists
if [ -f "$ICON_PATH" ]; then
    cp "$ICON_PATH" "${APP_PATH}/Contents/Resources/app.icns"
    ICON_SET="true"
    echo "✅ Icon added to app"
else
    ICON_SET="false"
    echo "⚠️  Icon not found (app will use default icon)"
fi

# Create Info.plist
if [ "$ICON_SET" = "true" ]; then
    ICON_ENTRY="    <key>CFBundleIconFile</key>
    <string>app.icns</string>
    <key>CFBundleIconName</key>
    <string>app.icns</string>"
else
    ICON_ENTRY=""
fi

cat > "${APP_PATH}/Contents/Info.plist" << EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>${APP_NAME}</string>
    <key>CFBundleIdentifier</key>
    <string>com.bobperformance.webinterface</string>
    <key>CFBundleName</key>
    <string>${APP_NAME}</string>
    <key>CFBundleVersion</key>
    <string>1.0</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>LSMinimumSystemVersion</key>
    <string>10.12</string>
${ICON_ENTRY}
</dict>
</plist>
EOF

echo "✅ Launcher app created at: ${APP_PATH}"
echo ""
echo "You can now double-click '${APP_NAME}' on your Desktop to start the web interface!"
echo ""
echo "To move it to Applications folder:"
echo "  mv '${APP_PATH}' /Applications/"

