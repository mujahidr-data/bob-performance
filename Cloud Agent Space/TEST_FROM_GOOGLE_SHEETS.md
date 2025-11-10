# Testing from Google Sheets - Step by Step Guide

This guide shows you how to test and execute the HiBob Performance Report automation from Google Sheets.

## Prerequisites

Before testing from Google Sheets, ensure:

1. ✅ Apps Script code is deployed
2. ✅ Menu appears in Google Sheets
3. ✅ Python script is set up locally
4. ✅ Credentials can be set via menu

## Step 1: Open Your Google Sheet

1. Open: https://docs.google.com/spreadsheets/d/1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA
2. **Refresh the page** (to trigger `onOpen()` function)
3. Look for the menu: **"Bob Salary Data"** (top menu bar)

## Step 2: Set Up Credentials (First Time)

1. Click menu: **Bob Salary Data > Performance Reports > Set HiBob Credentials**
2. Enter your email: `mujahid.r@commerceiq.ai`
3. Click **OK**
4. Enter your password: `Reset@54321`
5. Click **OK**
6. You should see: **"Success! Credentials saved successfully!"**

### Verify Credentials Are Set

1. Click menu: **Bob Salary Data > Performance Reports > View Credentials Status**
2. Should show:
   - Email: `mujahid.r@commerceiq.ai`
   - Password: `Yes`

## Step 3: Test Apps Script Functions

### Test 1: Check Sheet Access

1. Open Apps Script editor:
   - Click menu: **Extensions > Apps Script**
   - Or visit: https://script.google.com/home/projects/1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47
2. Select function: `testSheetAccess` (from dropdown)
3. Click **Run** (▶️)
4. Check execution log - should see: **"Test passed!"**
5. Go back to Google Sheet
6. Verify sheet **"Bob Perf Report"** was created

### Test 2: Test Credentials API

1. In Apps Script editor, select function: `getCredentialsAPI`
2. Click **Run**
3. Check execution log - should show credentials (for testing only)
4. This verifies credentials are stored correctly

## Step 4: Execute the Automation

### Option A: From Terminal (Recommended)

The Python script runs locally, so you need to run it from your terminal:

1. **Open Terminal** on your local machine
2. Navigate to project:
   ```bash
   cd "/Users/mujahidreza/Cursor/Cloud Agent Space"
   ```
3. **Run the script:**
   ```bash
   python3 hibob_report_downloader.py
   ```
4. **Enter report name** when prompted:
   ```
   Enter the report name (or partial name) to search for: [Type report name]
   ```
5. **Watch the automation:**
   - Browser opens
   - Logs into HiBob
   - Downloads report
   - Uploads to Google Sheets

### Option B: Get Instructions from Google Sheets

1. In Google Sheet, click: **Bob Salary Data > Performance Reports > Instructions**
2. Read the instructions shown
3. Follow the steps to run locally

### Option C: Trigger Download (Menu Item)

1. In Google Sheet, click: **Bob Salary Data > Performance Reports > Download Performance Report**
2. This shows instructions and credential status
3. Then run the Python script locally

## Step 5: Verify Results in Google Sheet

After the Python script completes:

1. **Refresh your Google Sheet**
2. **Check for sheet:** "Bob Perf Report"
3. **Verify data:**
   - Headers are formatted (blue background, bold)
   - Data is imported correctly
   - Columns are auto-sized
   - Header row is frozen

## Complete Test Workflow

### Full End-to-End Test:

1. **Set credentials in Sheets:**
   - Menu: **Bob Salary Data > Performance Reports > Set HiBob Credentials**
   - Enter email and password

2. **Verify credentials:**
   - Menu: **Bob Salary Data > Performance Reports > View Credentials Status**

3. **Run Python script:**
   ```bash
   cd "/Users/mujahidreza/Cursor/Cloud Agent Space"
   python3 hibob_report_downloader.py
   ```

4. **Enter report name** when prompted

5. **Watch automation** (browser will open)

6. **Check Google Sheet** for imported data

## Troubleshooting from Google Sheets

### Menu Doesn't Appear

**Solution:**
1. Refresh the Google Sheet (F5 or Cmd+R)
2. Check Apps Script has `onOpen()` function
3. Verify code is saved in Apps Script
4. Try: **Extensions > Apps Script** to check for errors

### "Credentials not set" Error

**Solution:**
1. Click: **Bob Salary Data > Performance Reports > Set HiBob Credentials**
2. Enter credentials again
3. Verify: **View Credentials Status**

### Python Script Can't Fetch Credentials

**Solution:**
1. Verify Apps Script is deployed as web app
2. Check Apps Script URL in `config.json`
3. Test credentials API:
   - Visit: `https://script.google.com/.../exec?action=getCredentials`
   - Should return JSON with credentials

### Report Not Appearing in Sheet

**Solution:**
1. Check Apps Script execution logs
2. Verify `doPost()` function received the file
3. Check for errors in execution log
4. Verify sheet "Bob Perf Report" exists

## Quick Test Checklist

- [ ] Menu "Bob Salary Data" appears in Google Sheet
- [ ] Submenu "Performance Reports" is visible
- [ ] Can set credentials via menu
- [ ] Can view credentials status
- [ ] Python script runs locally
- [ ] Python script fetches credentials from Apps Script
- [ ] Browser automation works
- [ ] Report downloads successfully
- [ ] Report appears in Google Sheet
- [ ] Data is formatted correctly

## Menu Reference

**Main Menu:** Bob Salary Data

**Performance Reports Submenu:**
- **Set HiBob Credentials** - Store login credentials
- **View Credentials Status** - Check if credentials are set
- **Download Performance Report** - Show instructions
- **Instructions** - Complete setup guide

## What Happens When You Click Menu Items

### "Set HiBob Credentials"
- Prompts for email
- Prompts for password
- Stores in Script Properties
- Shows success message

### "View Credentials Status"
- Shows email (visible)
- Shows password status (Yes/No, not the actual password)
- Confirms credentials are ready

### "Download Performance Report"
- Shows credential status
- Provides instructions to run Python script
- Explains the process

### "Instructions"
- Shows complete setup guide
- Explains how to run the automation
- Provides troubleshooting tips

## Next Steps After Testing

Once testing is successful:

1. ✅ Document any issues encountered
2. ✅ Note any report-specific requirements
3. ✅ Share with team (if applicable)
4. ✅ Set up scheduling (if needed)

---

**Remember:** The Python script must run locally. Google Sheets menu is for:
- Setting credentials
- Viewing status
- Getting instructions
- Managing configuration

The actual automation runs on your local machine via the Python script.

