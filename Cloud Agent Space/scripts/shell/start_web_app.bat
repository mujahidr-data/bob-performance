@echo off
REM Startup script for HiBob Performance Report Downloader Web Interface (Windows)

echo ==========================================
echo HiBob Performance Report Downloader
echo Web Interface Startup
echo ==========================================
echo.

REM Get the directory where this script is located
cd /d "%~dp0"
REM Change to project root (parent of scripts/)
cd /d "%~dp0.."

REM Check if Python 3 is available
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python 3 is not installed or not in PATH
    echo    Please install Python 3 and try again
    pause
    exit /b 1
)

echo [OK] Python found
python --version
echo.

REM Check if required packages are installed
echo [INFO] Checking dependencies...
python -c "import flask" >nul 2>&1
if errorlevel 1 (
    echo [WARN] Missing dependency: flask
    set MISSING=1
)

python -c "import playwright" >nul 2>&1
if errorlevel 1 (
    echo [WARN] Missing dependency: playwright
    set MISSING=1
)

python -c "import pandas" >nul 2>&1
if errorlevel 1 (
    echo [WARN] Missing dependency: pandas
    set MISSING=1
)

if defined MISSING (
    echo.
    set /p INSTALL="Install missing dependencies? (y/n): "
    if /i "%INSTALL%"=="y" (
        echo [INFO] Installing dependencies...
        pip install -r scripts\python\requirements.txt
        if errorlevel 1 (
            echo [ERROR] Failed to install dependencies
            pause
            exit /b 1
        )
        echo [OK] Dependencies installed
    ) else (
        echo [ERROR] Cannot start without required dependencies
        pause
        exit /b 1
    )
) else (
    echo [OK] All dependencies installed
)

echo.

REM Check if config.json exists
if not exist "config\config.json" (
    echo [WARN] config.json not found
    if exist "config\config.template.json" (
        echo    Creating config.json from template...
        copy config\config.template.json config\config.json >nul
        echo [OK] config.json created
        echo.
        echo [WARN] IMPORTANT: Please edit config\config.json with your credentials:
        echo    - HiBob email
        echo    - HiBob password
        echo.
        pause
    ) else (
        echo [ERROR] config\config.template.json not found
        pause
        exit /b 1
    )
)

REM Get port from command line or use default
set PORT=%1
if "%PORT%"=="" set PORT=5000

echo.
echo [INFO] Starting web interface...
echo.
echo ==========================================
echo   Web Interface Starting
echo ==========================================
echo.
echo [INFO] Open your browser and go to:
echo    http://localhost:%PORT%
echo    or
echo    http://127.0.0.1:%PORT%
echo.
echo [INFO] Press Ctrl+C to stop the server
echo.
echo ==========================================
echo.

REM Try to open browser automatically
timeout /t 2 /nobreak >nul
start http://localhost:%PORT%

REM Start the web app
python web\web_app.py %PORT%

pause

