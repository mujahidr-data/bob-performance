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

# Create the executable script
cat > "${APP_PATH}/Contents/MacOS/${APP_NAME}" << 'EOF'
#!/bin/bash
# Get the directory where this app is located
APP_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
PROJECT_ROOT="$( cd "$APP_DIR/../../.." && pwd )"
cd "$PROJECT_ROOT"
exec "$PROJECT_ROOT/scripts/shell/start_web_app.sh"
EOF

chmod +x "${APP_PATH}/Contents/MacOS/${APP_NAME}"

# Create Info.plist
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
</dict>
</plist>
EOF

echo "✅ Launcher app created at: ${APP_PATH}"
echo ""
echo "You can now double-click '${APP_NAME}' on your Desktop to start the web interface!"
echo ""
echo "To move it to Applications folder:"
echo "  mv '${APP_PATH}' /Applications/"

