# Quick Test Guide - From Google Sheets

## 🚀 Fast Test (5 Minutes)

### Step 1: Open Google Sheet
https://docs.google.com/spreadsheets/d/1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA

### Step 2: Set Credentials
1. Click menu: **Bob Salary Data > Performance Reports > Set HiBob Credentials**
2. Email: `mujahid.r@commerceiq.ai`
3. Password: `Reset@54321`

### Step 3: Verify
Click: **Bob Salary Data > Performance Reports > View Credentials Status**
- Should show email and "Password: Yes"

### Step 4: Run Python Script
```bash
cd "/Users/mujahidreza/Cursor/Cloud Agent Space"
python3 hibob_report_downloader.py
```

### Step 5: Enter Report Name
When prompted, type the report name (or partial name)

### Step 6: Check Results
- Go back to Google Sheet
- Look for sheet: **"Bob Perf Report"**
- Verify data is imported

## ✅ Success Indicators

- Menu appears in Google Sheet
- Credentials save successfully
- Python script runs without errors
- Browser opens and logs in
- Report downloads
- Data appears in Google Sheet

## ❌ Common Issues

**Menu not appearing?**
→ Refresh the Google Sheet (F5)

**Credentials not saving?**
→ Check you have edit access to Apps Script

**Python script fails?**
→ Check `config.json` has Apps Script URL
→ Verify credentials are set in Google Sheets

---

**That's it!** The automation should work end-to-end.

