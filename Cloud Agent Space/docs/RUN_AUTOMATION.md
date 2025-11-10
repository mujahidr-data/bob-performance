# Run HiBob Performance Report Automation - Step by Step

Follow these steps to run the automation for the first time.

## Pre-Flight Checklist

Before running, verify:

- [ ] Python 3.7+ installed
- [ ] Dependencies installed (`pip3 install -r requirements.txt`)
- [ ] Playwright browser installed (`playwright install chromium`)
- [ ] `config.json` exists with credentials
- [ ] Apps Script deployed as web app
- [ ] Credentials set in Google Sheets (optional - can use config.json)

## Step 1: Verify Setup

### Check Python
```bash
python3 --version
# Should show: Python 3.9.6 or higher
```

### Check Dependencies
```bash
cd "/Users/mujahidreza/Cursor/Cloud Agent Space"
python3 -c "import playwright; import requests; print('✅ All dependencies installed')"
```

### Check Config
```bash
cat config.json
# Should show email, password, and apps_script_url
```

## Step 2: Set Credentials (Choose One Method)

### Method A: Use config.json (Already Set)
Your `config.json` already has:
- Email: `mujahid.r@commerceiq.ai`
- Password: `Reset@54321`
- Apps Script URL: Already configured

✅ **You're ready to go!**

### Method B: Use Google Sheets Menu (Alternative)
1. Open Google Sheet: https://docs.google.com/spreadsheets/d/1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA
2. Refresh page
3. Click: **Bob Salary Data > Performance Reports > Set HiBob Credentials**
4. Enter credentials
5. The Python script will automatically fetch from Apps Script

## Step 3: Run the Automation

### Navigate to Project
```bash
cd "/Users/mujahidreza/Cursor/Cloud Agent Space"
```

### Run the Script
```bash
python3 hibob_report_downloader.py
```

### What Happens:
1. **Script starts** - Shows welcome message
2. **Prompts for report name** - Enter the report name (or partial name)
3. **Browser opens** - You'll see it logging in
4. **Login process**:
   - Navigates to HiBob
   - Enters email
   - Redirects to JumpCloud SSO
   - Enters password
   - Logs in
5. **Navigation**:
   - Goes to Performance Cycles page
   - Searches for your report
6. **Report selection**:
   - If multiple matches: Shows list, you select one
   - If one match: Auto-selects
7. **Download**:
   - Opens report
   - Clicks Actions button
   - Clicks Download
   - Waits for download
8. **Upload**:
   - Uploads to Google Sheets via Apps Script
   - Shows success message

## Step 4: Verify Results

### Check Google Sheet
1. Open: https://docs.google.com/spreadsheets/d/1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA
2. Look for sheet: **"Bob Perf Report"**
3. Verify:
   - Data is imported
   - Headers are formatted (blue background, bold)
   - Columns are auto-sized
   - Header row is frozen

## Example Run

```bash
$ python3 hibob_report_downloader.py

============================================================
  HiBob Performance Report Downloader
============================================================

📝 Enter the report name (or partial name) to search for: Q4 Review

🌐 Starting browser...
🔐 Navigating to HiBob login page...
📧 Entering email address...
⏭️  Clicking continue...
🔄 Waiting for JumpCloud SSO redirect...
🔑 Entering JumpCloud credentials...
✅ Signing in...
⏳ Waiting for login to complete...
✅ Successfully logged in to HiBob!
📊 Navigating to Performance Cycles page...
✅ On Performance Cycles page
🔍 Searching for report: 'Q4 Review'
✅ Found 1 matching report: Q4 Performance Review 2024
📄 Opening report...
🔍 Looking for Actions button...
📥 Clicking Actions button...
🔍 Looking for Download option...
⬇️  Clicking Download...
✅ Report downloaded: downloads/q4-review-2024.csv
☁️  Uploading to Google Sheets...
✅ Successfully uploaded to Google Sheets!
   Successfully imported report to Google Sheet

🎉 Process completed successfully!

🧹 Cleaning up...
```

## Troubleshooting

### "Configuration file not found"
```bash
cp config.template.json config.json
# Then edit config.json with your credentials
```

### "Playwright not installed"
```bash
playwright install chromium
```

### "Module not found"
```bash
pip3 install -r requirements.txt
```

### Login Fails
- Check credentials in `config.json`
- Verify email/password are correct
- Check screenshot: `error_login.png` (if created)

### Report Not Found
- Try a shorter/partial name
- Check screenshot: `error_search.png`
- Verify you're on the correct page

### Upload Fails
- Verify Apps Script URL in `config.json`
- Check Apps Script is deployed as web app
- Check Apps Script execution logs

## Quick Start Command

```bash
cd "/Users/mujahidreza/Cursor/Cloud Agent Space" && python3 hibob_report_downloader.py
```

## What to Have Ready

Before running:
1. ✅ Know the report name (or partial name) you want to download
2. ✅ Have internet connection
3. ✅ Browser will open (don't close it manually)
4. ✅ Let the script complete (takes 1-3 minutes)

---

**Ready to run?** Execute the command above and follow the prompts!

