# Launch Web Interface from Google Sheets

## How to Use

1. **Open your Google Sheet**
   - Go to: https://docs.google.com/spreadsheets/d/1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA/edit

2. **Click the Menu**
   - Look for **"🤖 Bob Automation"** in the menu bar
   - Click **"Performance Reports"** submenu
   - Click **"🚀 Launch Web Interface"**

3. **Follow the Instructions**
   - A dialog will open with step-by-step instructions
   - **Step 1:** Start the web server on your local machine
     ```bash
     ./start_web_app.sh
     ```
   - **Step 2:** Click the button to open the web interface

4. **Use the Web Interface**
   - Enter the report name
   - Click "Start Automation"
   - Select a report using checkboxes
   - Watch the progress bar

## What the Menu Does

The **"🚀 Launch Web Interface"** menu item:
- ✅ Opens a helpful dialog with instructions
- ✅ Provides clickable buttons to open localhost:5000
- ✅ Shows how to start the web server
- ✅ Lists all the features of the web interface

## Important Notes

⚠️ **The web interface runs on your local machine**
- You need to start the server first: `./start_web_app.sh`
- The menu provides buttons to open the interface, but the server must be running
- If the server isn't running, the links won't work

## Menu Location

```
🤖 Bob Automation
  ├── Import Base Data
  ├── Import Bonus History
  ├── Import Compensation History
  ├── Import Full Comp History
  ├── Import All Data
  ├── Convert Tenure to Array Formula
  └── Performance Reports
      ├── 🚀 Launch Web Interface  ← NEW!
      ├── Set HiBob Credentials
      ├── View Credentials Status
      └── 📖 Instructions & Help
```

## Troubleshooting

**Menu item doesn't appear?**
- Refresh the Google Sheet page
- Make sure you have edit access to the sheet
- The `onOpen()` function should run automatically

**Button doesn't open the interface?**
- Make sure the web server is running: `./start_web_app.sh`
- Check if port 5000 is available
- Try the alternative URL: `http://127.0.0.1:5000`

**Need to update the menu?**
- The menu is updated automatically when you push to Apps Script
- Run: `npx clasp push`

