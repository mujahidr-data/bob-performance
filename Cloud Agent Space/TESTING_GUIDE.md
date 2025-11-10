# Testing Guide - HiBob Report Automation

Follow these steps to test the complete automation system.

## Prerequisites Checklist

Before testing, ensure you have:

- [ ] Python 3.7+ installed
- [ ] Node.js and npm installed
- [ ] clasp installed and logged in
- [ ] Access to HiBob account
- [ ] Access to Google Sheet: https://docs.google.com/spreadsheets/d/1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA

## Step 1: Install Python Dependencies

```bash
cd "/Users/mujahidreza/Cursor/Cloud Agent Space"
pip install -r requirements.txt
playwright install chromium
```

**Verify:**
```bash
python --version
playwright --version
```

## Step 2: Configure Credentials

```bash
# Copy template
cp config.template.json config.json

# Edit config.json with your credentials
# You'll need:
# - email: Your HiBob email
# - password: Your JumpCloud password
# - apps_script_url: (We'll get this in Step 3)
```

**Note:** Don't commit `config.json` - it's in `.gitignore`

## Step 3: Deploy Apps Script Code

### 3a. Push Code to Apps Script

```bash
# Make sure you're logged in to clasp
clasp login --status

# If not logged in:
clasp login

# Push the code
npm run push
```

**Expected output:**
```
📤 Pushing to Apps Script...
✅ Pushed successfully.
📦 Committing and pushing to Git...
✅ Deployment complete!
```

### 3b. Deploy as Web App

1. **Open Apps Script:**
   ```bash
   npm run open
   ```
   Or visit: https://script.google.com/home/projects/1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47

2. **Deploy as Web App:**
   - Click **Deploy** > **New deployment**
   - Click gear icon ⚙️ > Select **Web app**
   - Configure:
     - **Execute as**: Me
     - **Who has access**: Anyone
   - Click **Deploy**
   - **Copy the Web app URL**

3. **Update config.json:**
   ```json
   {
     "email": "your-email@company.com",
     "password": "your-password",
     "apps_script_url": "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec"
   }
   ```

### 3c. Test Apps Script (Optional but Recommended)

1. In Apps Script editor, select function `testSheetAccess`
2. Click **Run**
3. Check execution log - should see "Test passed!"
4. Verify sheet "Bob Perf Report" was created in your Google Sheet

## Step 4: Test Python Automation

### 4a. Quick Test - Verify Script Runs

```bash
python hibob_report_downloader.py
```

**Expected:**
- Script prompts for report name
- Browser opens (you'll see it)
- Login process starts

**If errors occur:**
- Check `config.json` has correct credentials
- Check error screenshots (error_*.png files)
- Verify Apps Script URL is correct

### 4b. Full End-to-End Test

1. **Run the script:**
   ```bash
   python hibob_report_downloader.py
   ```

2. **Enter report name when prompted:**
   ```
   Enter the report name (or partial name) to search for: [Type report name]
   ```

3. **Watch the automation:**
   - Browser will open
   - Login to HiBob
   - Navigate to Performance Cycles
   - Search for report
   - Download report
   - Upload to Google Sheets

4. **Expected output:**
   ```
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
   ```

5. **Verify in Google Sheets:**
   - Open: https://docs.google.com/spreadsheets/d/1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA
   - Check for sheet "Bob Perf Report"
   - Verify data is imported correctly
   - Check formatting (blue headers, frozen row)

## Step 5: Troubleshooting Common Issues

### Issue: "Configuration file not found"
**Solution:**
```bash
cp config.template.json config.json
# Edit config.json with your credentials
```

### Issue: "Playwright not installed"
**Solution:**
```bash
playwright install chromium
```

### Issue: Login fails
**Check:**
- Credentials in `config.json` are correct
- Check screenshot: `error_login.png`
- Try logging in manually to verify credentials work

### Issue: Report not found
**Try:**
- Use a shorter/partial report name
- Check screenshot: `error_search.png`
- Verify you're on the correct page

### Issue: Upload to Google Sheets fails
**Check:**
- Apps Script URL in `config.json` is correct
- Apps Script is deployed as web app
- Run `testSheetAccess()` in Apps Script
- Check Apps Script execution logs

### Issue: "clasp: command not found"
**Solution:**
```bash
npm install -g @google/clasp
# Or use npx:
npx @google/clasp push
```

## Step 6: Verify Complete Flow

### Test Checklist:

- [ ] ✅ Python script runs without errors
- [ ] ✅ Browser opens and navigates correctly
- [ ] ✅ Login to HiBob works
- [ ] ✅ JumpCloud SSO login works
- [ ] ✅ Navigation to Performance Cycles works
- [ ] ✅ Report search finds the report
- [ ] ✅ Report downloads successfully
- [ ] ✅ File uploads to Apps Script
- [ ] ✅ Data appears in Google Sheet
- [ ] ✅ Sheet formatting is correct (headers, colors)
- [ ] ✅ Apps Script code is in git
- [ ] ✅ Python code is in git

## Step 7: Production Readiness

Once testing is successful:

1. **Verify git repository:**
   ```bash
   git status
   git log
   ```

2. **Push to remote (if not already done):**
   ```bash
   git push bob-performance main
   ```

3. **Document any customizations:**
   - Note any report-specific requirements
   - Document any selector changes needed
   - Update README with findings

## Quick Test Commands Summary

```bash
# 1. Install dependencies
pip install -r requirements.txt && playwright install chromium

# 2. Configure
cp config.template.json config.json
# Edit config.json

# 3. Deploy Apps Script
npm run push
# Then deploy as web app in browser, copy URL to config.json

# 4. Test
python hibob_report_downloader.py

# 5. Verify
# Check Google Sheet for imported data
```

## Next Steps After Successful Test

1. ✅ Document any issues encountered
2. ✅ Note any report-specific requirements
3. ✅ Set up scheduling (if needed)
4. ✅ Share with team (if applicable)
5. ✅ Create backup of config.json (securely)

## Support

If you encounter issues:
1. Check error screenshots in project directory
2. Review Apps Script execution logs
3. Check console output for specific error messages
4. Refer to README.md for detailed troubleshooting

---

**Ready to test?** Start with Step 1 and work through each step sequentially!

