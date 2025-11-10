# Code Consolidation Guide

This guide explains how to replace the existing multiple Apps Script files with a single, optimized consolidated file.

## What Was Consolidated

**Before (3 separate files):**
- `bob-salary-data.js` - Salary data imports
- `HiBobPerformanceMenu.gs` - Performance report menu
- `HiBobReportHandler.gs` - File upload handler

**After (1 module file):**
- `bob-performance-module.gs` - All functionality in one optimized file

## Benefits of Consolidation

✅ **Single source of truth** - All code in one place
✅ **Easier maintenance** - No duplicate functions
✅ **Better organization** - Clear sections by functionality
✅ **Shared helpers** - Common functions reused
✅ **Optimized** - Removed duplicate code
✅ **Unified menu** - Single onOpen() function

## Deployment Steps

### Step 1: Backup Current Code (Optional but Recommended)

1. In Apps Script, select all files
2. Copy all code to a backup location
3. Or use clasp to pull current version:
   ```bash
   clasp pull
   ```

### Step 2: Replace Files in Apps Script

**Option A: Replace All Files (Recommended)**

1. Open Apps Script: https://script.google.com/home/projects/1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47
2. Delete existing files:
   - `bob-salary-data.js`
   - `HiBobPerformanceMenu.gs`
   - `HiBobReportHandler.gs`
   - `UPDATE_EXISTING_ONOPEN.gs` (if exists)
3. Create new file: `bob-performance-module.gs`
4. Copy entire contents of `bob-performance-module.gs` from this repository
5. Paste into the new file
6. Save (Ctrl+S / Cmd+S)

**Option B: Update Existing File**

1. Open `bob-salary-data.js` in Apps Script
2. Replace ALL contents with `bob-performance-module.gs`
3. Delete other files:
   - `HiBobPerformanceMenu.gs`
   - `HiBobReportHandler.gs`
4. Save

### Step 3: Verify Functions

1. Check that all functions are present:
   - `onOpen()` - Menu creation
   - `importBobDataSimpleWithLookup()` - Base data import
   - `importBobBonusHistoryLatest()` - Bonus history
   - `importBobCompHistoryLatest()` - Comp history
   - `importBobFullCompHistory()` - Full comp history
   - `setHiBobCredentials()` - Set credentials
   - `getCredentialsAPI()` - Credentials API
   - `doGet()` - GET handler
   - `doPost()` - POST handler
   - All helper functions

2. Test menu appears:
   - Open Google Sheet
   - Refresh page
   - Verify "Bob Salary Data" menu appears with all items

### Step 4: Test Key Functions

1. **Test Menu:**
   - Open Google Sheet
   - Menu should appear: "Bob Salary Data"
   - Should have "Performance Reports" submenu

2. **Test Credentials:**
   - Menu: Bob Salary Data > Performance Reports > Set HiBob Credentials
   - Set test credentials
   - Verify: View Credentials Status

3. **Test Import (if needed):**
   - Menu: Bob Salary Data > Import Base Data
   - Should work as before

4. **Test Web App:**
   - Deploy as web app (if not already deployed)
   - Test credentials API: `?action=getCredentials`
   - Test file upload via Python script

### Step 5: Update Local Repository

After consolidating in Apps Script:

1. **Pull from Apps Script:**
   ```bash
   clasp pull
   ```

2. **Update local files:**
   - Keep `bob-performance-module.gs` as the main file
   - Can delete old separate files (or keep as backup)

3. **Update .clasp.json if needed:**
   - Should point to correct script ID
   - Root dir should be correct

4. **Commit and push:**
   ```bash
   git add .
   git commit -m "Consolidate Apps Script code into single optimized file"
   git push bob-performance main
   ```

## File Structure After Consolidation

```
Apps Script Project:
└── bob-performance-module.gs    (Single file with all functionality)

Local Repository:
├── bob-performance-module.gs    (Main module file)
├── bob-salary-data.js     (Can be deleted or kept as backup)
├── HiBobPerformanceMenu.gs (Can be deleted or kept as backup)
└── HiBobReportHandler.gs  (Can be deleted or kept as backup)
```

## Code Organization in Consolidated File

The consolidated file is organized into clear sections:

1. **Constants** - All configuration values
2. **UI Functions** - Menu creation
3. **Salary Data Functions** - Import functions
4. **Performance Report Functions** - Credentials and instructions
5. **Web App Handlers** - doGet, doPost
6. **Helper Functions** - Shared utilities

## Rollback Plan

If you need to rollback:

1. Use clasp to pull previous version:
   ```bash
   clasp pull
   ```

2. Or restore from git:
   ```bash
   git checkout HEAD~1 -- bob-salary-data.js
   git checkout HEAD~1 -- HiBobPerformanceMenu.gs
   git checkout HEAD~1 -- HiBobReportHandler.gs
   ```

3. Push to Apps Script:
   ```bash
   clasp push
   ```

## Verification Checklist

After consolidation, verify:

- [ ] Menu appears in Google Sheet
- [ ] All menu items work
- [ ] Salary data imports work
- [ ] Performance report credentials can be set
- [ ] Credentials API works (`?action=getCredentials`)
- [ ] File upload works (doPost)
- [ ] No errors in execution logs
- [ ] All functions are accessible

## Next Steps

1. ✅ Deploy consolidated code to Apps Script
2. ✅ Test all functionality
3. ✅ Update local repository
4. ✅ Push to git
5. ✅ Document any issues

---

**The consolidated file is ready!** Follow the steps above to deploy it.

