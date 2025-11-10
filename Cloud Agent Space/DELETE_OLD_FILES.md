# Delete Old Apps Script Files

Since clasp doesn't have a delete command, you need to manually delete old files in Apps Script.

## Current Status

After pulling, Apps Script shows:
- ✅ `bob-consolidated.js` (or .gs) - This is the consolidated file
- ✅ `appsscript.json` - Manifest file

## Files to Delete in Apps Script

If these files still exist in Apps Script, delete them:

1. ❌ `bob-salary-data.js` (or .gs)
2. ❌ `HiBobPerformanceMenu.gs`
3. ❌ `HiBobReportHandler.gs`
4. ❌ `UPDATE_EXISTING_ONOPEN.gs`

## Steps to Delete in Apps Script

### Method 1: Using Apps Script UI

1. **Open Apps Script:**
   - Visit: https://script.google.com/home/projects/1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47
   - Or click: Extensions > Apps Script in your Google Sheet

2. **Delete Old Files:**
   - In the left sidebar, find the file list
   - For each old file:
     - Right-click on the file name
     - Select "Delete"
     - Confirm deletion

3. **Verify:**
   - You should only have:
     - `bob-consolidated.gs` (or .js)
     - `appsscript.json`

### Method 2: Using clasp (if files are tracked)

If the old files are still being tracked, you can:

1. **Check what clasp sees:**
   ```bash
   npx @google/clasp pull
   ```

2. **If old files are pulled, delete them locally and push:**
   ```bash
   # Delete old files locally (already done)
   npx @google/clasp push
   ```

   Note: This won't delete files from Apps Script, but will ensure only the consolidated file is pushed.

## Verification

After deleting:

1. **Check Apps Script:**
   - Should only see `bob-consolidated.gs` and `appsscript.json`

2. **Test Menu:**
   - Open Google Sheet
   - Refresh page
   - Menu should work correctly

3. **Test Functions:**
   - Try: Bob Salary Data > Import Base Data
   - Try: Bob Salary Data > Performance Reports > Set HiBob Credentials

## Quick Command Reference

```bash
# Pull current state from Apps Script
npx @google/clasp pull

# Push only consolidated file
npx @google/clasp push

# Open Apps Script in browser
npx @google/clasp open
```

---

**Note:** clasp doesn't have a delete command, so manual deletion in Apps Script UI is required.

