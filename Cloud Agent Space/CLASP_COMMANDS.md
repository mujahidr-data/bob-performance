# Clasp Commands Quick Reference

This project is set up with npm scripts similar to your other bob projects. Here are the available commands:

## NPM Scripts

### Push to Apps Script
```bash
npm run push
```
Pushes your code to Google Apps Script using clasp.

### Pull from Apps Script
```bash
npm run pull
```
Downloads code from Google Apps Script to your local files.

### Open Apps Script in Browser
```bash
npm run open
```
Opens your Apps Script project in the browser.

### Deploy (Push + Git)
```bash
npm run deploy
```
Pushes to Apps Script AND commits/pushes to git automatically.

### Watch Mode (Auto-deploy on save)
```bash
npm run watch
```
Watches for file changes and automatically deploys. Requires `nodemon`:
```bash
npm install -g nodemon
```

### Sync (Using deploy.sh)
```bash
npm run sync
# or directly:
./deploy.sh
```
Runs the deploy script that pushes to Apps Script and git.

## Direct Clasp Commands

You can also use clasp directly:

```bash
# Push code
clasp push

# Pull code
clasp pull

# Open in browser
clasp open

# Check status
clasp status

# View logs
clasp logs

# Create version
clasp version "Description"

# List deployments
clasp deployments
```

## Quick Start

1. **First time setup** (if not already done):
   ```bash
   clasp login
   ```

2. **Push your code**:
   ```bash
   npm run push
   ```

3. **Or use the deploy script** (pushes to Apps Script + Git):
   ```bash
   ./deploy.sh
   ```

## Project Configuration

- **Script ID**: `1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47`
- **Config File**: `.clasp.json`
- **Main File**: `HiBobReportHandler.gs`

## Troubleshooting

### "clasp: command not found"
Make sure clasp is installed:
```bash
npm install -g @google/clasp
```

Or use npx:
```bash
npx @google/clasp push
```

### "User has not enabled the Apps Script API"
1. Go to https://script.google.com/home/usersettings
2. Enable "Google Apps Script API"
3. Try again

### "Access Not Granted or Expired"
```bash
clasp logout
clasp login
```

---

For more details, see [CLASP_SETUP.md](CLASP_SETUP.md)

