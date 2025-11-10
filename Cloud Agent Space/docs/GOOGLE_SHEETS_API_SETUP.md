# Google Sheets API Setup Guide

Since Apps Script deployment isn't working, we're using Google Sheets API directly.

## Option 1: Service Account (RECOMMENDED for Automation)

Service accounts are better for automated scripts - no browser interaction needed!

### Step 1: Enable Google Sheets API

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the Google Sheets API:
   - Go to "APIs & Services" → "Library"
   - Search for "Google Sheets API"
   - Click "Enable"

### Step 2: Create Service Account

1. Go to "APIs & Services" → "Credentials"
2. Click "Create Credentials" → "Service account"
3. Fill in the details:
   - Service account name: "hibob-uploader"
   - Service account ID: (auto-generated)
   - Description: "Service account for HiBob report uploads"
   - Click "Create and Continue"
4. Grant access (optional - can skip):
   - Click "Continue"
   - Click "Done"
5. Create and download key:
   - Click on the service account you just created
   - Go to "Keys" tab
   - Click "Add Key" → "Create new key"
   - Choose "JSON"
   - Click "Create"
   - The JSON file will download automatically
6. Save the file:
   - Rename it to `service_account.json`
   - Save it in this project directory

### Step 3: Share Google Sheet with Service Account

1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA/edit
2. Click "Share" button (top right)
3. Get the service account email from `service_account.json`:
   - Open `service_account.json`
   - Find the `client_email` field (looks like: `hibob-uploader@your-project.iam.gserviceaccount.com`)
4. Share the sheet:
   - Paste the service account email in the "Add people" field
   - Set permission to "Editor"
   - Click "Send" (you can uncheck "Notify people")

### Step 4: Install Dependencies

```bash
pip3 install google-auth google-api-python-client
```

### Step 5: Test

Run the test script:
```bash
python3 test_upload.py
```

The script will automatically use the service account - no browser needed!

---

## Option 2: OAuth 2.0 (Alternative - Requires Browser)

If you prefer user credentials instead of service account:

### Step 1-2: Same as above (Enable API)

### Step 3: Create OAuth 2.0 Credentials

1. Go to "APIs & Services" → "Credentials"
2. Click "Create Credentials" → "OAuth client ID"
3. If prompted, configure the OAuth consent screen:
   - User Type: Internal (for Workspace) or External
   - App name: "HiBob Performance Report Uploader"
   - User support email: Your email
   - Developer contact: Your email
   - Click "Save and Continue"
   - Scopes: Add `https://www.googleapis.com/auth/spreadsheets`
   - Click "Save and Continue"
   - Test users: Add your email if using External
   - Click "Save and Continue"
4. Create OAuth client:
   - Application type: **Desktop app**
   - Name: "HiBob Uploader"
   - Click "Create"
5. Download the credentials:
   - Click the download icon next to your OAuth client
   - Save the file as `credentials.json` in this project directory

### Step 4: Install Dependencies

```bash
pip3 install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client
```

### Step 5: First Run Authentication

When you run the script for the first time:
1. A browser window will open
2. Sign in with your Google account (the one that has access to the sheet)
3. Click "Allow" to grant permissions
4. A `token.pickle` file will be created (saves your credentials)

---

## Troubleshooting

### Service Account Issues:
- **"service_account.json not found"**: Download service account key from Google Cloud Console
- **"Access denied"**: Share the Google Sheet with the service account email (from `client_email` in JSON)
- **"API not enabled"**: Enable Google Sheets API in Google Cloud Console

### OAuth2 Issues:
- **"credentials.json not found"**: Download OAuth credentials from Google Cloud Console
- **"Access denied"**: Make sure the Google account has edit access to the sheet
- **Token expired**: Delete `token.pickle` and re-authenticate

## Which Should I Use?

- **Service Account**: Best for automation, scripts, CI/CD - no user interaction needed
- **OAuth2**: Better if you want to use your personal Google account credentials

