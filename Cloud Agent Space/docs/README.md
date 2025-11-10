# HiBob Performance Report Automation

Automated tool to download performance cycle reports from HiBob and upload them to Google Sheets.

> **Note**: See the main `README.md` in the project root for the most up-to-date quick start guide and launcher information.

## Overview

This automation:
1. Logs into HiBob via JumpCloud SSO
2. Navigates to the Performance Cycles page
3. Searches for and downloads a specific report
4. Uploads the report to a Google Sheet automatically

## Architecture

- **Local Python Script**: Uses Playwright for browser automation and downloads reports
- **Google Apps Script**: Receives uploaded files and imports them to Google Sheets
- **Communication**: Simple HTTP POST with minimal dependencies

## Prerequisites

- Python 3.7 or higher
- pip (Python package installer)
- Internet connection
- Google account with access to the target Google Sheet

## Installation

### 1. Clone the Repository

```bash
cd /path/to/your/workspace
git clone <repository-url>
cd bob-performance
```

### 2. Install Python Dependencies

```bash
pip install -r requirements.txt
```

### 3. Install Playwright Browsers

```bash
playwright install chromium
```

This downloads the Chromium browser needed for automation.

### 4. Configure Credentials

Create your configuration file from the template:

```bash
cp config.template.json config.json
```

Edit `config.json` with your credentials:

```json
{
  "email": "your-email@company.com",
  "password": "your-jumpcloud-password",
  "apps_script_url": "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec"
}
```

**Important**: `config.json` is in `.gitignore` and will not be committed to git for security.

### 5. Set Up Google Apps Script

#### Step 1: Access Your Apps Script Project

1. Open your Google Apps Script project: `1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47`
2. Go to: https://script.google.com/home/projects/1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47

#### Step 2: Add the Handler Code

1. In Apps Script, click the "+" icon to add a new file
2. Name it `HiBobReportHandler`
3. Copy the entire contents of `HiBobReportHandler.gs` from this repository
4. Paste it into the new file
5. Click "Save" (or Ctrl+S / Cmd+S)

#### Step 3: Deploy as Web App

1. Click "Deploy" > "New deployment"
2. Click the gear icon ⚙️ next to "Select type"
3. Choose "Web app"
4. Configure deployment settings:
   - **Description**: "HiBob Report Upload Handler"
   - **Execute as**: "Me (your-email@gmail.com)"
   - **Who has access**: "Anyone"
5. Click "Deploy"
6. Review and authorize permissions when prompted
7. Copy the "Web app URL" (looks like `https://script.google.com/macros/s/.../exec`)
8. Paste this URL into your `config.json` as the `apps_script_url`

#### Step 4: Test Apps Script Setup (Optional)

1. In Apps Script editor, select function `testSheetAccess` from dropdown
2. Click "Run"
3. Check "Execution log" - should see "Test passed!"
4. Verify that a new sheet "Bob Perf Report" was created in your Google Sheet: https://docs.google.com/spreadsheets/d/1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA

## Usage

### Running the Automation

1. Open a terminal
2. Navigate to the project directory
3. Run the script:

```bash
python hibob_report_downloader.py
```

4. When prompted, enter the report name (or partial name):

```
Enter the report name (or partial name) to search for: Q4 Performance Review
```

5. If multiple matching reports are found, you'll be asked to select one:

```
Found 3 matching reports:
  1. Q4 Performance Review 2024
  2. Q4 Performance Review 2023
  3. Q3-Q4 Performance Summary

Enter the number of the report to download (1-3): 1
```

6. The script will:
   - Launch a browser window
   - Login to HiBob (you'll see this happening)
   - Navigate to the Performance Cycles page
   - Find and download the report
   - Upload it to Google Sheets
   - Display success message

### Example Output

```
============================================================
  HiBob Performance Report Downloader
============================================================

📝 Enter the report name (or partial name) to search for: Annual Review

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
🔍 Searching for report: 'Annual Review'
✅ Found 1 matching report: Annual Review 2024
📄 Opening report...
🔍 Looking for Actions button...
📥 Clicking Actions button...
🔍 Looking for Download option...
⬇️  Clicking Download...
✅ Report downloaded: downloads/annual-review-2024.csv
☁️  Uploading to Google Sheets...
✅ Successfully uploaded to Google Sheets!
   Successfully imported report to Google Sheet

🎉 Process completed successfully!

🧹 Cleaning up...
```

## Troubleshooting

### Error: Configuration file not found

**Solution**: Create `config.json` from the template:
```bash
cp config.template.json config.json
```
Then edit with your credentials.

### Error: Playwright browser not installed

**Solution**: Install Playwright browsers:
```bash
playwright install chromium
```

### Login Fails

**Possible causes**:
- Incorrect email or password in `config.json`
- MFA/2FA enabled on your account (not supported yet)
- JumpCloud login page structure changed

**Solution**: 
- Verify credentials in `config.json`
- Check the screenshot saved as `error_login.png`
- Run with `headless=False` to watch the browser

### Report Not Found

**Possible causes**:
- Report name doesn't match (try a shorter/partial name)
- Report is in a different location
- Page structure changed

**Solution**:
- Try searching with just a few keywords
- Check the screenshot saved as `error_search.png`

### Download Fails

**Possible causes**:
- Actions button location changed
- Download button text changed
- Insufficient permissions

**Solution**:
- Check the screenshot saved as `error_download.png`
- Verify you have permission to download the report manually

### Upload to Google Sheets Fails

**Possible causes**:
- Incorrect `apps_script_url` in config.json
- Apps Script not deployed correctly
- Sheet permissions issue

**Solution**:
- Re-check the Apps Script deployment URL
- Run `testSheetAccess()` in Apps Script editor
- Verify you have edit access to the Google Sheet

## File Structure

```
bob-performance/
├── hibob_report_downloader.py   # Main Python automation script
├── HiBobReportHandler.gs         # Google Apps Script handler
├── requirements.txt              # Python dependencies
├── config.template.json          # Configuration template
├── config.json                   # Your credentials (gitignored)
├── .gitignore                    # Git ignore rules
├── README.md                     # This file
└── downloads/                    # Downloaded reports (gitignored)
```

## Security Notes

- **Never commit `config.json`** to git (it's in `.gitignore`)
- Store credentials securely
- The Apps Script web app accepts requests from "Anyone" but doesn't expose data
- Downloaded reports are stored locally in `downloads/` folder
- Consider using environment variables for extra security

## Advanced Configuration

### Running in Headless Mode

To run without showing the browser window, edit `hibob_report_downloader.py`:

```python
self.browser = playwright.chromium.launch(headless=True)  # Change False to True
```

### Changing Download Directory

Edit the `__init__` method in `hibob_report_downloader.py`:

```python
self.download_dir = Path("my-custom-folder")
```

### Timeout Adjustments

If pages load slowly, increase timeouts in the script (default is 30 seconds):

```python
self.page.goto(url, wait_until="networkidle", timeout=60000)  # 60 seconds
```

## Target Google Sheet

- **Sheet ID**: `1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA`
- **Target Sheet Name**: `Bob Perf Report`
- **Sheet URL**: https://docs.google.com/spreadsheets/d/1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA

Each time the automation runs:
- Creates the "Bob Perf Report" sheet if it doesn't exist
- Clears existing data if sheet exists
- Imports the downloaded report
- Formats headers (bold, blue background)
- Auto-resizes columns
- Freezes the header row

## Future Enhancements

Potential improvements to consider:
- Support for Excel file formats (.xlsx, .xls)
- Multi-factor authentication (MFA) support
- Scheduled automation (cron job)
- Email notifications on completion/failure
- Report history tracking
- Custom data transformations before upload

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Review error screenshots in the project directory
3. Check Apps Script execution logs
4. Review the automation execution output

## License

This is a private automation tool for internal use.

