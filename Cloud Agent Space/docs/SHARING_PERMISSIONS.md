# Sharing Google Sheet - Permissions Guide

This guide explains what happens when you share the Google Sheet with others.

## Current Setup

Your Apps Script is **bound** to the Google Sheet, which means:
- The script is attached to the sheet
- Menu functions are available to users with edit access
- Script Properties (credentials) are only accessible to Apps Script project editors

## What Shared Users CAN Do

### If They Have EDIT Access to Sheet:

✅ **See the Menu**
- "Bob Salary Data" menu appears when they open the sheet
- All menu items are visible

✅ **Run Salary Data Import Functions**
- Import Base Data
- Import Bonus History
- Import Compensation History
- Import Full Comp History
- Import All Data
- Convert Tenure to Array Formula

⚠️ **Run Performance Report Functions (with limitations)**
- Can see "Performance Reports" submenu
- Can click "Set HiBob Credentials" (but won't work - see below)
- Can click "View Credentials Status" (but won't see credentials)
- Can click "Download Performance Report" (instructions only)
- Can click "Instructions" (instructions only)

### If They Have VIEW ONLY Access:

❌ **Cannot see menu** - Menu only appears for users with edit access
❌ **Cannot run any functions**

## What Shared Users CANNOT Do

### Script Properties Access:

❌ **Cannot access stored credentials**
- Script Properties (HIBOB_EMAIL, HIBOB_PASSWORD) are only accessible to:
  - Users with **edit access to the Apps Script project itself**
  - Not just edit access to the Google Sheet

❌ **Cannot set credentials via menu**
- "Set HiBob Credentials" will fail with permission error
- They need Apps Script project access to modify Script Properties

❌ **Cannot modify Apps Script code**
- Apps Script project is separate from sheet sharing
- They need explicit access to the Apps Script project

### Web App Access:

⚠️ **Cannot directly call web app** (unless you share the URL)
- The web app URL is in your `config.json`
- If they have the URL, they can POST files (but it executes as "Me")
- This is generally not recommended for security

## How to Share Properly

### Option 1: Share Sheet Only (Recommended for Most Users)

**For users who just need to view/use the data:**

1. Share Google Sheet with **View** or **Edit** access
2. They can:
   - View the data
   - Use salary data import functions (if Edit access)
   - See menu items
3. They cannot:
   - Access Performance Report credentials
   - Run Performance Report automation (needs local Python script anyway)

**This is the default and recommended approach.**

### Option 2: Share Apps Script Project (For Admins)

**For users who need to manage credentials:**

1. Share Google Sheet with **Edit** access
2. Share Apps Script project:
   - Open: https://script.google.com/home/projects/1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47
   - Click **Share** (top right)
   - Add user with **Editor** access
3. They can now:
   - Set credentials via menu
   - View credentials status
   - Modify Apps Script code (if needed)

**Use this only for trusted administrators.**

### Option 3: No Sharing (Most Secure)

**For sensitive operations:**

- Don't share the sheet
- Only you run the automation
- Results are visible only to you

**Best for sensitive performance data.**

## Security Considerations

### Script Properties (Credentials)

- ✅ Stored securely in Google's encrypted storage
- ✅ Only accessible to Apps Script project editors
- ✅ Not visible in sheet sharing
- ⚠️ If you share Apps Script project, they can see/modify credentials

### Web App Deployment

- Current setting: "Execute as: Me" - Always runs as you
- Current setting: "Who has access: Anyone" - Anyone with URL can call it
- ⚠️ Consider adding authentication token for production use

### Menu Functions

- Run with the permissions of the user who clicks them
- First-time users need to authorize the script
- Some functions may fail if they lack permissions

## Best Practices

### For Regular Users:
1. ✅ Share sheet with **View** or **Edit** access
2. ✅ They can use salary data imports
3. ❌ Don't share Apps Script project
4. ❌ They cannot use Performance Report automation (needs local setup anyway)

### For Administrators:
1. ✅ Share sheet with **Edit** access
2. ✅ Share Apps Script project with **Editor** access
3. ✅ They can manage credentials
4. ⚠️ They can modify code - only share with trusted users

### For Performance Reports:
1. ✅ Only you (or trusted admins) run the Python script locally
2. ✅ Results appear in shared sheet
3. ✅ Others can view results but not trigger automation

## Testing Sharing

To test what a shared user sees:

1. **Share with test account:**
   - Share sheet with another Google account
   - Give **Edit** access

2. **Test as that user:**
   - Open sheet in incognito/private window
   - Log in as the test account
   - Check what menu items appear
   - Try clicking menu items

3. **Expected behavior:**
   - Menu appears ✅
   - Salary data imports work ✅
   - Performance Report credentials: "Permission denied" ❌
   - This is expected and secure ✅

## Summary

| User Type | Sheet Access | Apps Script Access | Can Use Menu | Can Set Credentials | Can Run Python Script |
|-----------|--------------|-------------------|--------------|---------------------|----------------------|
| **Viewer** | View | None | ❌ No menu | ❌ No | ❌ No |
| **Editor** | Edit | None | ✅ Yes | ❌ No | ❌ No (needs local setup) |
| **Admin** | Edit | Editor | ✅ Yes | ✅ Yes | ❌ No (needs local setup) |
| **You** | Owner | Owner | ✅ Yes | ✅ Yes | ✅ Yes |

## Recommendation

**For most use cases:**
- Share sheet with **Edit** access for users who need to import salary data
- Don't share Apps Script project (keeps credentials secure)
- Only you (or trusted admins) run the Performance Report automation
- Results are visible to all shared users

**This provides the right balance of functionality and security.**

