# Using clasp to Push Apps Script Code

This guide explains how to use clasp (Command Line Apps Script Projects) to push your code to Google Apps Script.

## What is clasp?

clasp is Google's official tool for developing Apps Script projects locally and pushing them to Google's servers.

## Installation

### Install clasp globally

```bash
npm install -g @google/clasp
```

If you get permission errors on macOS/Linux:

```bash
sudo npm install -g @google/clasp
```

### Verify installation

```bash
clasp --version
```

## Authentication

### Login to clasp (one-time setup)

```bash
clasp login
```

This will:
1. Open a browser window
2. Ask you to authorize clasp
3. Save your credentials locally in `~/.clasprc.json`

### Check login status

```bash
clasp login --status
```

## Configuration Files

I've created these files for you:

### `.clasp.json`
```json
{
  "scriptId": "1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47",
  "rootDir": "."
}
```

This tells clasp which Apps Script project to push to.

### `appsscript.json`
```json
{
  "timeZone": "America/New_York",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```

This is the Apps Script manifest file.

## Pushing Code to Apps Script

### Option 1: Use the helper script (easiest)

```bash
./push_to_apps_script.sh
```

This script will:
- Check if clasp is installed
- Check if you're logged in
- Push the code to your Apps Script project

### Option 2: Manual clasp commands

```bash
# Push all files
clasp push

# Push and watch for changes (auto-push on save)
clasp push --watch
```

### Option 3: Push specific files

```bash
# Push only the HiBob handler
clasp push -f HiBobReportHandler.gs
```

## What Gets Pushed?

When you run `clasp push`, it will push:
- `HiBobReportHandler.gs` - Your Apps Script code
- `appsscript.json` - Project manifest

It will NOT push (automatically ignored by clasp):
- `.clasp.json` - Local config
- Python files (`.py`)
- `config.json` - Credentials
- Other local files

You can create a `.claspignore` file if you want to customize what gets pushed.

## Common Commands

### Open the project in browser
```bash
clasp open
```

### Pull code from Apps Script (download)
```bash
clasp pull
```

### View project info
```bash
clasp status
```

### Create a new version
```bash
clasp version "Description of changes"
```

### Deploy as web app
```bash
clasp deploy --description "HiBob Report Handler"
```

### List deployments
```bash
clasp deployments
```

## Troubleshooting

### "User has not enabled the Apps Script API"

1. Go to https://script.google.com/home/usersettings
2. Enable "Google Apps Script API"
3. Try again

### "Access Not Granted or Expired"

```bash
clasp login --creds credentials.json
```

Or logout and login again:
```bash
clasp logout
clasp login
```

### "Error: Could not read API credentials"

Make sure you're logged in:
```bash
clasp login
```

### "Error: Looks like you are offline"

Check your internet connection and try again.

### Permission Errors on Install

If `npm install -g @google/clasp` fails with EACCES:

**Option 1 (Recommended)**: Configure npm to use a different directory
```bash
mkdir ~/.npm-global
npm config set prefix '~/.npm-global'
echo 'export PATH=~/.npm-global/bin:$PATH' >> ~/.zshrc
source ~/.zshrc
npm install -g @google/clasp
```

**Option 2**: Use sudo (less secure)
```bash
sudo npm install -g @google/clasp
```

## Workflow

### Initial Setup
1. Install clasp: `npm install -g @google/clasp`
2. Login: `clasp login`
3. Verify config: check `.clasp.json` has correct scriptId

### Making Changes
1. Edit `HiBobReportHandler.gs` locally
2. Push to Apps Script: `clasp push` or `./push_to_apps_script.sh`
3. Test in Apps Script editor or via web app

### Continuous Development
```bash
# Watch for changes and auto-push
clasp push --watch
```

Now you can edit locally and changes push automatically!

## Alternative: Manual Copy-Paste

If clasp doesn't work for you, you can still manually:
1. Open the Apps Script project in browser
2. Copy code from `HiBobReportHandler.gs`
3. Paste into the Apps Script editor
4. Save

## Links

- **Your Apps Script Project**: https://script.google.com/home/projects/1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47
- **clasp Documentation**: https://github.com/google/clasp
- **Apps Script API Settings**: https://script.google.com/home/usersettings

---

## Quick Start Summary

```bash
# Install clasp
npm install -g @google/clasp

# Login
clasp login

# Push code (from project directory)
cd "/Users/mujahidreza/Cursor/Cloud Agent Space"
clasp push

# Or use the helper script
./push_to_apps_script.sh
```

That's it! Your code is now in Google Apps Script.

