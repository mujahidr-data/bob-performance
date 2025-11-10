# Quick Setup Instructions

Follow these steps to get the HiBob Report Automation running.

## 1. Install Python Dependencies

```bash
pip install -r requirements.txt
playwright install chromium
```

## 2. Configure Your Credentials

```bash
# Copy the template
cp config.template.json config.json

# Edit config.json with your credentials
# Required fields:
# - email: Your HiBob email address
# - password: Your JumpCloud password
# - apps_script_url: Your Apps Script deployment URL (see step 3)
```

## 3. Deploy Google Apps Script

### A. Access the Apps Script Project

Visit: https://script.google.com/home/projects/1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47

### B. Add the Code

1. Click "+" to add new file
2. Name it `HiBobReportHandler`
3. Copy all contents from `HiBobReportHandler.gs`
4. Paste into the new file
5. Save (Ctrl+S / Cmd+S)

### C. Deploy as Web App

1. Click **Deploy** > **New deployment**
2. Click gear icon ⚙️ next to "Select type"
3. Choose **Web app**
4. Settings:
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Click **Deploy**
6. Authorize when prompted
7. **Copy the Web app URL**
8. Paste URL into `config.json` as `apps_script_url`

### D. Test (Optional)

1. Select function `testSheetAccess` from dropdown
2. Click Run
3. Check logs for "Test passed!"

## 4. Run the Automation

```bash
python hibob_report_downloader.py
```

When prompted, enter the report name to search for.

## Verification Checklist

- [ ] Python 3.7+ installed
- [ ] Dependencies installed (`pip install -r requirements.txt`)
- [ ] Playwright browsers installed (`playwright install chromium`)
- [ ] `config.json` created with valid credentials
- [ ] Apps Script code added to project
- [ ] Apps Script deployed as web app
- [ ] Web app URL added to `config.json`
- [ ] Can access Google Sheet: https://docs.google.com/spreadsheets/d/1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA

## Troubleshooting

### "Configuration file not found"
Create config.json: `cp config.template.json config.json`

### "Playwright not installed"
Run: `playwright install chromium`

### "Upload failed"
- Verify Apps Script URL in config.json
- Re-deploy Apps Script web app
- Run testSheetAccess() in Apps Script

## What Happens When You Run It

1. ✅ Browser opens automatically
2. ✅ Logs into HiBob via JumpCloud
3. ✅ Navigates to Performance Cycles
4. ✅ Searches for your report
5. ✅ Downloads the report
6. ✅ Uploads to Google Sheets
7. ✅ Shows success message

## Next Steps

After successful setup and first run:
- Report will appear in "Bob Perf Report" sheet
- You can customize report parsing/transformations
- Contact admin for additional features

---

For detailed documentation, see [README.md](README.md)

