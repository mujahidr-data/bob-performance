# Auto-Push Setup Guide

Your project is now configured for automatic pushing to both Apps Script and Git!

## How Auto-Push Works

### 1. Git Hooks (Automatic)

**Pre-commit Hook** (`.git/hooks/pre-commit`):
- Automatically pushes Apps Script files to Google before committing
- Only runs if `.gs` or `.js` files are being committed
- Ensures Apps Script is always in sync with git

**Post-commit Hook** (`.git/hooks/post-commit`):
- Automatically pushes commits to git remote
- Runs after every `git commit`
- Pushes to `bob-performance` remote (or `origin` if not available)

### 2. Manual Push Commands

**Quick Push (Apps Script + Git):**
```bash
npm run push
```
Pushes to Apps Script, then commits and pushes to git.

**Apps Script Only:**
```bash
npm run clasp:push
```

**Git Only:**
```bash
npm run git:push
```

**Quick Update (Fast):**
```bash
npm run quick-push
```
Pushes to Apps Script and git with minimal commit message.

### 3. File Watcher (Auto-push on Save)

**Using nodemon (Recommended):**
```bash
npm run watch
```
Watches for changes to `.gs` and `.js` files and auto-pushes.

**Using custom script:**
```bash
./auto-push.sh
```
Watches all files and auto-pushes on changes.

**Manual trigger:**
```bash
./auto-push.sh file1.gs file2.js
```
Pushes specific files immediately.

### 4. Deploy Script

**Full deployment:**
```bash
./deploy.sh
```
Or:
```bash
npm run sync
```

## Usage Examples

### Scenario 1: Edit Apps Script File

1. Edit `HiBobReportHandler.gs`
2. Save file
3. Run: `npm run push`
   - ✅ Pushes to Apps Script
   - ✅ Commits to git
   - ✅ Pushes to remote

### Scenario 2: Auto-watch Mode

1. Run: `npm run watch`
2. Edit any `.gs` or `.js` file
3. Save file
4. ✅ Automatically pushes to Apps Script and git

### Scenario 3: Git Commit (Automatic)

1. Edit files
2. Run: `git add . && git commit -m "My changes"`
3. ✅ Pre-commit hook pushes to Apps Script
4. ✅ Post-commit hook pushes to git remote

### Scenario 4: Quick Update

1. Make small change
2. Run: `npm run quick-push`
3. ✅ Fast push to both Apps Script and git

## Configuration

### Git Remotes

The scripts prioritize:
1. `bob-performance` remote
2. `origin` remote (fallback)

Check your remotes:
```bash
git remote -v
```

### Apps Script

Configured in `.clasp.json`:
```json
{
  "scriptId": "1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47"
}
```

## Troubleshooting

### "clasp: command not found"
```bash
npm install -g @google/clasp
# Or use npx:
npx clasp push
```

### "Git push failed"
- Check network connection
- Verify remote is configured: `git remote -v`
- Check permissions: `git push bob-performance main`

### "Apps Script push failed"
- Verify you're logged in: `clasp login --status`
- Check `.clasp.json` has correct scriptId
- Try: `clasp push` manually

### Hooks not running
```bash
# Make sure hooks are executable
chmod +x .git/hooks/pre-commit
chmod +x .git/hooks/post-commit
```

### File watcher not working
```bash
# Install nodemon (for npm run watch)
npm install -g nodemon

# Or install fswatch (macOS) for auto-push.sh
brew install fswatch
```

## Workflow Recommendations

### Daily Development

1. **Start watching:**
   ```bash
   npm run watch
   ```

2. **Edit files** - changes auto-push

3. **Stop watching:** Ctrl+C

### Manual Workflow

1. **Edit files**

2. **Push when ready:**
   ```bash
   npm run push
   ```

### Quick Updates

1. **Make small change**

2. **Quick push:**
   ```bash
   npm run quick-push
   ```

## What Gets Pushed

### To Apps Script:
- `*.gs` files (Google Apps Script)
- `*.js` files (if any)
- `appsscript.json` (manifest)

### To Git:
- All tracked files
- Commits with timestamp
- Pushes to `bob-performance` remote

## Security Notes

- ✅ Git hooks run locally only
- ✅ Credentials not committed (in `.gitignore`)
- ✅ Apps Script credentials stored in Script Properties
- ⚠️  Post-commit hook pushes automatically (can disable if needed)

## Disable Auto-Push (If Needed)

### Disable Git Hooks:
```bash
chmod -x .git/hooks/pre-commit
chmod -x .git/hooks/post-commit
```

### Re-enable:
```bash
chmod +x .git/hooks/pre-commit
chmod +x .git/hooks/post-commit
```

## Summary

✅ **Pre-commit hook** → Auto-pushes Apps Script before commit
✅ **Post-commit hook** → Auto-pushes git after commit
✅ **npm run push** → Manual push to both
✅ **npm run watch** → Auto-push on file save
✅ **npm run quick-push** → Fast push option

Everything is set up for automatic pushing! 🚀

