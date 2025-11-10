# Integration Instructions - Adding Performance Reports to Existing Menu

This guide shows how to integrate the Performance Report functionality into your existing Google Apps Script project.

## Current Setup

Your Apps Script project (`1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47`) already has:
- ✅ `bob-salary-data.js` with menu and import functions
- ✅ `onOpen()` function that creates "Bob Salary Data" menu

## What We're Adding

1. **Credential Storage** - Store HiBob email/password in Apps Script Properties
2. **Menu Items** - Add "Performance Reports" submenu to existing menu
3. **API Endpoint** - Allow Python script to fetch credentials
4. **Python Integration** - Update Python script to use Apps Script credentials

## Step 1: Add Menu Functions to Apps Script

### Option A: Add New File (Recommended)

1. In Apps Script editor, click **"+"** to add new file
2. Name it: `HiBobPerformanceMenu`
3. Copy all code from `HiBobPerformanceMenu.gs` in this repository
4. Paste into the new file
5. Save

### Option B: Update Existing onOpen() Function

1. Open `bob-salary-data.js` in Apps Script
2. Find the `onOpen()` function (around line 38)
3. Replace it with the code from `UPDATE_EXISTING_ONOPEN.gs`
4. Save

## Step 2: Add Credential Functions

The `HiBobPerformanceMenu.gs` file includes all needed functions:
- `setHiBobCredentials()` - Store credentials
- `viewCredentialsStatus()` - Check if credentials are set
- `getHiBobCredentials()` - Internal function to retrieve credentials
- `getCredentialsAPI()` - API endpoint for Python script

## Step 3: Update HiBobReportHandler.gs

1. Open `HiBobReportHandler.gs` in Apps Script
2. Add the `doGet()` function (for credentials API)
3. Add the `getCredentialsAPI()` function
4. The updated code is already in `HiBobReportHandler.gs` in this repository

## Step 4: Set Initial Credentials

1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA
2. Refresh the page (to trigger `onOpen()`)
3. Click menu: **Bob Salary Data > Performance Reports > Set HiBob Credentials**
4. Enter:
   - Email: `mujahid.r@commerceiq.ai`
   - Password: `Reset@54321`
5. Click OK

## Step 5: Update Python Script

The Python script (`hibob_report_downloader.py`) has been updated to:
- ✅ Automatically fetch credentials from Apps Script if not in config.json
- ✅ Fall back to config.json if Apps Script fetch fails
- ✅ Show helpful error messages

## Step 6: Test the Integration

### Test 1: Menu Appears
1. Open Google Sheet
2. Refresh page
3. Verify menu: **Bob Salary Data > Performance Reports** appears

### Test 2: Set Credentials
1. Click: **Bob Salary Data > Performance Reports > Set HiBob Credentials**
2. Enter credentials
3. Verify success message

### Test 3: View Status
1. Click: **Bob Salary Data > Performance Reports > View Credentials Status**
2. Should show email and "Password: Yes"

### Test 4: Python Script Uses Credentials
1. Run: `python3 hibob_report_downloader.py`
2. Script should automatically fetch credentials from Apps Script
3. Should see: "✅ Credentials loaded from Apps Script"

## File Structure After Integration

```
Apps Script Project:
├── bob-salary-data.js          (existing - salary data imports)
├── HiBobReportHandler.gs       (updated - file upload handler)
└── HiBobPerformanceMenu.gs     (new - menu and credentials)
```

## How It Works

### Credential Flow:
1. **User sets credentials** via Google Sheets menu
2. **Stored in Script Properties** (secure, only accessible to project)
3. **Python script fetches** via GET request to Apps Script
4. **Used for login** to HiBob

### Menu Flow:
1. **User opens sheet** → `onOpen()` runs
2. **Menu appears** with "Performance Reports" submenu
3. **User clicks menu item** → Function runs
4. **Credentials stored/retrieved** as needed

## Security Notes

- ✅ Credentials stored in Script Properties (encrypted by Google)
- ✅ Only accessible to users with edit access to Apps Script
- ✅ API endpoint returns credentials (consider adding auth token for production)
- ✅ Password is masked in status view

## Troubleshooting

### Menu doesn't appear
- Refresh the Google Sheet
- Check that `onOpen()` function exists and is saved
- Check execution logs for errors

### Credentials not saving
- Verify you have edit access to Apps Script
- Check Script Properties in Project Settings
- Try setting credentials again

### Python can't fetch credentials
- Verify Apps Script is deployed as web app
- Check that `doGet()` function exists in `HiBobReportHandler.gs`
- Test URL: `https://script.google.com/.../exec?action=getCredentials`
- Check Apps Script execution logs

### Python falls back to config.json
- This is normal if Apps Script fetch fails
- Make sure `config.json` has credentials as backup
- Check Apps Script URL is correct

## Next Steps

After integration:
1. ✅ Test menu appears in Google Sheet
2. ✅ Set credentials via menu
3. ✅ Test Python script fetches credentials
4. ✅ Run full automation test
5. ✅ Document any customizations

## Quick Reference

**Set Credentials:**
```
Google Sheet > Bob Salary Data > Performance Reports > Set HiBob Credentials
```

**Check Status:**
```
Google Sheet > Bob Salary Data > Performance Reports > View Credentials Status
```

**Run Automation:**
```bash
python3 hibob_report_downloader.py
```

---

**All files are ready!** Follow the steps above to integrate into your existing Apps Script project.

