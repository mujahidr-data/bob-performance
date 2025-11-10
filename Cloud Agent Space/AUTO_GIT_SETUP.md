# Automatic Git Upload Setup

Your project is now configured to automatically push to Git whenever you deploy to Apps Script!

## How It Works

### 1. **Automatic Git Push with `npm run push`**

Now when you run:
```bash
npm run push
```

It will:
1. ✅ Push to Google Apps Script
2. ✅ Add all changes to git
3. ✅ Commit with timestamp
4. ✅ Push to `bob-performance` remote

### 2. **Git Hook (post-commit)**

A git hook is installed that automatically pushes to the remote after every commit. This means:
- Any manual `git commit` will also trigger a push
- Works silently in the background
- Won't fail if network is unavailable

### 3. **Deploy Script**

The `deploy.sh` script now:
- Pushes to Apps Script first
- Then handles git operations
- Uses the correct remote (`bob-performance`)

## Available Commands

### Main Commands (Auto Git)

```bash
# Push to Apps Script + Auto Git (RECOMMENDED)
npm run push

# Deploy (same as push)
npm run deploy

# Use deploy script
./deploy.sh
```

### Apps Script Only (No Git)

```bash
# Push to Apps Script only (no git)
npm run clasp:push
```

### Git Only

```bash
# Check git status
npm run git:status

# Manual git push (if needed)
npm run git:push
```

## Git Remotes

Your project is configured with:
- **Primary**: `bob-performance` → https://github.com/mujahidr-data/bob-performance.git
- **Secondary**: `origin` → https://github.com/mujahidr-data/Bob-Automation.git

The scripts prioritize `bob-performance` remote.

## Workflow Examples

### Standard Workflow
```bash
# Make changes to HiBobReportHandler.gs
# Then push (automatically commits and pushes to git)
npm run push
```

### Watch Mode (Auto-deploy on save)
```bash
# Install nodemon if not already installed
npm install -g nodemon

# Watch for changes and auto-deploy
npm run watch
```

### Manual Git Workflow
```bash
# Make changes
git add .
git commit -m "Your message"
# Git hook automatically pushes to remote
```

## What Gets Committed

The auto-commit includes:
- ✅ All modified files
- ✅ All new files
- ✅ Uses timestamp: `Auto-update: YYYY-MM-DD HH:MM:SS`

## Disabling Auto-Git (If Needed)

If you want to push to Apps Script without git:

```bash
# Use the clasp-only command
npm run clasp:push
```

Or temporarily disable the git hook:
```bash
chmod -x .git/hooks/post-commit
```

Re-enable it:
```bash
chmod +x .git/hooks/post-commit
```

## Troubleshooting

### "Git push failed or no changes"
This is normal if:
- There are no changes to commit
- You're already up to date
- Network issue (will retry next time)

### "No git remote configured"
Make sure you have a remote:
```bash
git remote -v
```

If missing, add it:
```bash
git remote add bob-performance https://github.com/mujahidr-data/bob-performance.git
```

### Git Hook Not Working
Check if it's executable:
```bash
ls -la .git/hooks/post-commit
chmod +x .git/hooks/post-commit
```

## Summary

✅ **`npm run push`** → Apps Script + Git (automatic)
✅ **Git hook** → Auto-push after commits
✅ **`deploy.sh`** → Full deployment script
✅ **Remote priority** → `bob-performance` > `origin`

Everything is set up for automatic git uploads! 🚀

