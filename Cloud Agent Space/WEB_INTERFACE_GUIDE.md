# Web Interface Guide

Run the HiBob automation from a beautiful web interface instead of the terminal!

## Quick Start

### 1. Install Flask (if not already installed)

```bash
pip3 install flask werkzeug
```

Or install all requirements:
```bash
pip3 install -r requirements.txt
```

### 2. Start the Web Server

```bash
cd "/Users/mujahidreza/Cursor/Cloud Agent Space"
python3 web_app.py
```

You'll see:
```
============================================================
  HiBob Performance Report Downloader - Web Interface
============================================================

🌐 Starting web server...
📱 Open your browser and go to: http://localhost:5000
🛑 Press Ctrl+C to stop the server
```

### 3. Open in Browser

Open your browser and go to:
```
http://localhost:5000
```

## Features

### 🎨 Beautiful Web Interface
- Modern, responsive design
- Real-time status updates
- Progress tracking
- Easy report selection

### 📋 Report Selection
- See all matching reports in a numbered list
- Click to select a report
- Visual feedback for selected report
- Match scores shown for each report

### 📊 Real-time Updates
- Live status messages
- Progress log
- Success/error notifications
- Automatic refresh

## How to Use

1. **Enter Report Name**
   - Type the report name (or partial name) in the input field
   - Examples: "Q2/Q3", "Annual Review", "Q2"

2. **Start Automation**
   - Click "▶️ Start Automation" button
   - The automation will:
     - Start browser
     - Log in to HiBob
     - Navigate to Performance Cycles
     - Search for reports

3. **Select Report**
   - When reports are found, they'll appear in a list
   - Click on the report you want to download
   - The selected report will be highlighted

4. **Wait for Completion**
   - The automation will:
     - Download the selected report
     - Upload to Google Sheets
   - You'll see a success message when done

5. **Stop (if needed)**
   - Click "⏹️ Stop" button to cancel the automation

## Status Messages

- **Ready**: Waiting for you to start
- **Running**: Automation in progress
- **Selecting**: Waiting for report selection
- **Downloading**: Downloading the report
- **Uploading**: Uploading to Google Sheets
- **Success**: ✅ Completed successfully
- **Error**: ❌ Something went wrong

## Troubleshooting

### Port Already in Use

If port 5000 is already in use, edit `web_app.py` and change:
```python
app.run(host='0.0.0.0', port=5000, ...)
```
to a different port (e.g., `port=5001`).

### Browser Not Starting

The web interface uses the same browser automation as the terminal version. If you see browser errors:
- Check that Playwright is installed: `python3 -m playwright install chromium`
- Try running the terminal version first to verify setup

### Reports Not Showing

If reports don't appear:
- Check the browser console (F12) for errors
- Verify you're logged in successfully
- Try a different search term

## Advantages Over Terminal

✅ **Better UX**: Visual interface instead of text
✅ **Easier Selection**: Click to select instead of typing numbers
✅ **Real-time Updates**: See progress as it happens
✅ **No Terminal Needed**: Run from any browser
✅ **Shareable**: Others can use it if you share the URL (on same network)

## Security Note

The web interface runs on `localhost` by default, which means:
- ✅ Only accessible from your computer
- ✅ Safe for local use
- ⚠️ If you want to share with others on your network, change `host='0.0.0.0'` (already set)
- ⚠️ Don't expose to the internet without proper security

## Next Steps

1. Start the server: `python3 web_app.py`
2. Open browser: `http://localhost:5000`
3. Enter report name and start!

Enjoy the web interface! 🎉

