# Browser Crash Troubleshooting

If you're getting "Target page, context or browser has been closed" error, try these solutions:

## Quick Fixes

### Option 1: Try Headless Mode

Run the headless version:
```bash
python3 hibob_report_downloader_headless.py
```

This runs the browser in the background (no visible window) which is more stable on macOS.

### Option 2: Reinstall Playwright Browser

The browser binaries might be corrupted:
```bash
playwright install chromium --force
```

### Option 3: Update Playwright

Update to the latest version:
```bash
pip3 install --upgrade playwright
playwright install chromium
```

### Option 4: Check macOS Permissions

macOS might be blocking the browser. Check:
1. System Settings > Privacy & Security
2. Look for any blocked applications
3. Allow Chromium if it's blocked

### Option 5: Try Different Browser

Edit `hibob_report_downloader.py` and change:
```python
self.browser = self.playwright.chromium.launch(...)
```
to:
```python
self.browser = self.playwright.firefox.launch(...)
```

Then install Firefox:
```bash
playwright install firefox
```

## Common Causes

1. **macOS Security**: System blocking browser launch
2. **Corrupted Browser**: Browser binaries need reinstall
3. **Memory Issues**: System running out of resources
4. **Playwright Version**: Incompatible version

## Debug Steps

1. **Check Playwright version:**
   ```bash
   python3 -c "import playwright; print(playwright.__version__)"
   ```

2. **Test browser launch:**
   ```bash
   python3 -c "from playwright.sync_api import sync_playwright; p = sync_playwright().start(); b = p.chromium.launch(); print('OK'); b.close(); p.stop()"
   ```

3. **Check system logs:**
   - Open Console.app
   - Look for Chromium errors

## If Nothing Works

Use headless mode - it's more stable:
```bash
python3 hibob_report_downloader_headless.py
```

The automation will work the same, just without showing the browser window.

