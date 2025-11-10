# Google Apps Script Deployment Guide

This guide walks you through deploying the HiBob Report Handler in Google Apps Script.

## Target Configuration

- **Apps Script Project ID**: `1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47`
- **Target Google Sheet ID**: `1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA`
- **Target Sheet Name**: `Bob Perf Report`

## Step-by-Step Deployment

### Step 1: Access Your Apps Script Project

Click this link to open your project directly:
```
https://script.google.com/home/projects/1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47
```

Or navigate manually:
1. Go to https://script.google.com
2. Open "My Projects"
3. Find project ID: `1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47`

### Step 2: Add New File

1. In the Apps Script editor, look for the **Files** section on the left
2. Click the **"+"** icon next to "Files"
3. Select **"Script"**
4. Name the new file: `HiBobReportHandler`
5. Click **"Create"**

### Step 3: Copy Code

1. Open the file `HiBobReportHandler.gs` from this repository
2. **Select all** the code (Ctrl+A / Cmd+A)
3. **Copy** (Ctrl+C / Cmd+C)
4. Go back to Apps Script editor
5. **Paste** the code into the `HiBobReportHandler.gs` file you just created
6. **Save** (Ctrl+S / Cmd+S)

### Step 4: Test the Code (Optional but Recommended)

Before deploying, test that the code can access your sheet:

1. In the top toolbar, find the function dropdown (says "Select function")
2. Select **`testSheetAccess`**
3. Click **"Run"** (▶️ play button)
4. If prompted to authorize:
   - Click **"Review permissions"**
   - Choose your Google account
   - Click **"Advanced"** > **"Go to [Project Name] (unsafe)"**
   - Click **"Allow"**
5. Check the **"Execution log"** (View > Logs or Ctrl+Enter)
6. Should see: `"Test passed!"`

### Step 5: Deploy as Web App

1. Click **"Deploy"** button (top right, next to "Run")
2. Select **"New deployment"**
3. Click the **gear icon** (⚙️) next to "Select type"
4. Choose **"Web app"**
5. Fill in the deployment configuration:

   **Configuration:**
   - **Description**: `HiBob Report Upload Handler`
   - **Web app**
     - **Execute as**: `Me (your-email@gmail.com)`
     - **Who has access**: `Anyone`

6. Click **"Deploy"**

### Step 6: Authorize (First Time Only)

If this is your first deployment:

1. Click **"Authorize access"**
2. Choose your Google account
3. You may see "Google hasn't verified this app"
   - Click **"Advanced"**
   - Click **"Go to [Project Name] (unsafe)"**
4. Click **"Allow"** to grant permissions:
   - View and manage spreadsheets
   - Connect to external service

### Step 7: Copy Deployment URL

1. After successful deployment, you'll see a success message
2. **Copy** the **"Web app" URL** (looks like):
   ```
   https://script.google.com/macros/s/AKfycby.../exec
   ```
3. **Save this URL** - you'll need it for `config.json`

### Step 8: Update config.json

In your local repository:

1. Open `config.json`
2. Find the `apps_script_url` field
3. Paste the deployment URL you just copied
4. Save the file

Example:
```json
{
  "email": "your-email@company.com",
  "password": "your-password",
  "apps_script_url": "https://script.google.com/macros/s/AKfycby.../exec"
}
```

## Updating the Deployment

If you need to update the code later:

### Option A: New Deployment (Recommended)

1. Make changes to the code in Apps Script
2. Save changes
3. Click **"Deploy"** > **"New deployment"**
4. Follow steps 5-7 above
5. Update the URL in `config.json`

### Option B: Manage Deployments

1. Click **"Deploy"** > **"Manage deployments"**
2. Find your active deployment
3. Click **pencil icon** to edit
4. Update version or configuration
5. Click **"Deploy"**
6. URL remains the same (no need to update config.json)

## Verification

After deployment, verify everything works:

### Test 1: Run testSheetAccess()

1. In Apps Script, select `testSheetAccess` function
2. Click Run
3. Check logs for success message
4. Open the Google Sheet and verify "Bob Perf Report" sheet exists

### Test 2: Check Sheet Access

Open the target sheet directly:
```
https://docs.google.com/spreadsheets/d/1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA
```

Verify:
- You can open the sheet
- You have edit permissions
- "Bob Perf Report" sheet exists (after running test)

### Test 3: Run the Python Script

```bash
python hibob_report_downloader.py
```

Watch for the upload step - should show:
```
☁️  Uploading to Google Sheets...
✅ Successfully uploaded to Google Sheets!
```

## Troubleshooting

### "Script not found" or "Permission denied"

**Cause**: Insufficient permissions or wrong project ID

**Solution**:
- Verify you're logged into the correct Google account
- Check you have edit access to Apps Script project
- Confirm project ID: `1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47`

### "Cannot access spreadsheet"

**Cause**: Apps Script can't access the target Google Sheet

**Solution**:
- Verify sheet ID in `HiBobReportHandler.gs`: `1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA`
- Confirm you have edit permissions on the sheet
- Try sharing the sheet with yourself explicitly

### "Deployment failed" or "Authorization required"

**Cause**: Haven't authorized the necessary permissions

**Solution**:
- Re-run deployment
- Click through all authorization prompts
- Grant all requested permissions

### "Upload fails" from Python script

**Cause**: Incorrect deployment URL or Apps Script not accessible

**Solution**:
- Re-copy the deployment URL (check for extra spaces)
- Verify deployment is set to "Who has access: Anyone"
- Try creating a new deployment
- Check Apps Script execution logs (View > Executions)

## Security Notes

- The web app is set to "Anyone" access but doesn't expose sensitive data
- It only accepts POST requests with file uploads
- All files are imported directly to the specified sheet
- No data is stored in Apps Script
- Credentials are only in your local `config.json`

## Apps Script Execution Logs

To debug issues:

1. In Apps Script editor, go to **View** > **Executions**
2. Or visit: https://script.google.com/home/executions
3. Find recent executions
4. Click to view detailed logs
5. Look for errors or success messages

## What the Script Does

When a file is uploaded from the Python script:

1. ✅ Receives the file via HTTP POST
2. ✅ Opens Google Sheet: `1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA`
3. ✅ Creates "Bob Perf Report" sheet (or clears if exists)
4. ✅ Parses CSV file content
5. ✅ Writes data to sheet
6. ✅ Formats header row (bold, blue background, white text)
7. ✅ Auto-resizes all columns
8. ✅ Freezes header row
9. ✅ Returns success response

---

For more information, see [README.md](README.md) or [SETUP_INSTRUCTIONS.md](SETUP_INSTRUCTIONS.md)

