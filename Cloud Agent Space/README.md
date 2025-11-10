# Bob Performance Module

Automated tool to download performance cycle reports from HiBob and upload them to Google Sheets.

## 🚀 Quick Start

### Option 1: Desktop App (Easiest)
1. Double-click **`Bob Web Interface.app`** on your Desktop
2. The web interface will open at `http://localhost:5001`
3. Enter report name and start automation

### Option 2: Command File
1. Double-click **`Start Bob Web Interface.command`** in the project folder
2. The web interface will open automatically

### Option 3: Terminal
```bash
cd "/Users/mujahidreza/Cursor/Cloud Agent Space"
./scripts/shell/start_web_app.sh
```

Or use the alias (after adding to `~/.zshrc`):
```bash
bob-web
```

## 📋 Overview

This automation:
1. Logs into HiBob via JumpCloud SSO
2. Navigates to the Performance Cycles page
3. Searches for and downloads a specific report
4. Uploads the report data to Google Sheets automatically

## 🏗️ Architecture

- **Local Python Script**: Uses Playwright for browser automation
- **Web Interface**: Flask-based UI for easy operation
- **Google Apps Script**: Menu integration and credential management
- **Google Sheets API**: Direct data upload (no file upload needed)

## 📁 Project Structure

```
bob-performance/
├── apps-script/              # Google Apps Script files
│   ├── bob-performance-module.gs
│   └── .clasp.json           # Apps Script configuration
├── scripts/
│   ├── python/               # Python automation scripts
│   └── shell/                 # Shell scripts and launchers
├── web/                       # Flask web interface
├── config/                    # Configuration files
├── assets/                    # Icons and assets
├── docs/                      # Documentation
├── APPS_SCRIPT_ID.txt         # Apps Script project ID reference
└── create_launcher.sh        # Script to create Desktop app
```

## 🔧 Setup

### 1. Install Dependencies

```bash
pip3 install -r scripts/python/requirements.txt
playwright install chromium
```

### 2. Configure Credentials

Edit `config/config.json`:
```json
{
  "email": "your-email@example.com",
  "password": "your-password",
  "google_sheet_id": "1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA"
}
```

Or use the Google Sheets menu:
- **🤖 Bob Automation** → **Performance Reports** → **Set HiBob Credentials**

### 3. Google Sheets API Setup

See `docs/GOOGLE_SHEETS_API_SETUP.md` for:
- Service Account setup (recommended)
- OAuth2 setup (alternative)

## 🎯 Usage

### Web Interface (Recommended)

1. Start the web server (see Quick Start above)
2. Open `http://localhost:5001` in your browser
3. Enter the report name (e.g., "Q2/Q3")
4. Click "Start Automation"
5. Select a report from the list
6. Watch the progress bar

### From Google Sheets

1. Open the Google Sheet
2. Click **🤖 Bob Automation** → **Performance Reports** → **🚀 Launch Web Interface**
3. Follow the instructions in the dialog

### Terminal (Alternative)

```bash
python3 scripts/python/hibob_report_downloader.py
```

## 📚 Documentation

- **Web Interface**: `docs/README_WEB_INTERFACE.md`
- **Google Sheets Menu**: `docs/GOOGLE_SHEETS_MENU_GUIDE.md`
- **API Setup**: `docs/GOOGLE_SHEETS_API_SETUP.md`
- **Apps Script Deployment**: `docs/APPS_SCRIPT_DEPLOYMENT.md`
- **Troubleshooting**: `docs/TROUBLESHOOTING_WEB.md`

## 🔗 Important IDs

- **Apps Script Project ID**: `1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47`
  - See `APPS_SCRIPT_ID.txt` for details
- **Google Sheet ID**: `1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA`

## 🛠️ Development

### Push to Apps Script

```bash
cd apps-script
npx @google/clasp push
```

### Create Desktop Launcher

```bash
bash create_launcher.sh
```

This creates `Bob Web Interface.app` on your Desktop with a custom icon.

## 📝 License

UNLICENSED

## 👥 Author

Bob Performance Module Team

