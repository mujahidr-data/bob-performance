# Web Interface Quick Start

## Starting the Web Interface

### macOS / Linux
```bash
./start_web_app.sh
```

Or specify a custom port:
```bash
./start_web_app.sh 8080
```

### Windows
```cmd
start_web_app.bat
```

Or specify a custom port:
```cmd
start_web_app.bat 8080
```

## What the Script Does

1. ✅ Checks if Python 3 is installed
2. ✅ Verifies all required dependencies
3. ✅ Installs missing dependencies (with your permission)
4. ✅ Checks for Playwright browsers
5. ✅ Creates config.json from template if needed
6. ✅ Checks if port is available
7. ✅ Starts the web server
8. ✅ Opens browser automatically (macOS)

## Accessing the Interface

Once started, open your browser and go to:
- `http://localhost:5000` (default port)
- Or the port shown in the terminal

## Features

- 🎯 Visual progress tracking
- 📊 Progress bar with step indicators
- ☑️ Checkbox selection for reports
- 🔄 Real-time status updates
- 🛑 Stop button to cancel automation

## Troubleshooting

### Port Already in Use
If port 5000 is in use, the script will:
- Ask if you want to use a different port
- Or you can specify a port: `./start_web_app.sh 8080`

### Missing Dependencies
The script will detect and offer to install missing packages automatically.

### Browser Doesn't Open
Just manually navigate to `http://localhost:5000` in your browser.

## Stopping the Server

Press `Ctrl+C` in the terminal where the script is running.

