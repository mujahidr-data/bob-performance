# Web Interface Troubleshooting

## "Access to localhost was denied" Error

### Quick Fixes

#### 1. Try 127.0.0.1 instead of localhost

Instead of:
```
http://localhost:5000
```

Try:
```
http://127.0.0.1:5000
```

#### 2. Check if Server is Running

Make sure the server started successfully. You should see:
```
🌐 Starting web server...
📱 Open your browser and go to: http://127.0.0.1:5000
```

#### 3. Port Already in Use

If port 5000 is busy, use a different port:

```bash
python3 web_app.py 5001
```

Then open: `http://127.0.0.1:5001`

#### 4. Check Firewall

macOS might be blocking the connection. Check:
- System Settings > Network > Firewall
- Allow Python/Flask if prompted

#### 5. Try Different Browser

Sometimes browser security settings block localhost:
- Try Chrome, Firefox, or Safari
- Check browser console (F12) for errors

### Common Issues

#### Issue: "Address already in use"

**Solution:**
```bash
# Find what's using port 5000
lsof -i :5000

# Kill the process (replace PID with actual process ID)
kill -9 PID

# Or use a different port
python3 web_app.py 5001
```

#### Issue: "Connection refused"

**Causes:**
- Server not running
- Wrong port number
- Firewall blocking

**Solution:**
1. Make sure server is running
2. Check the port number in terminal output
3. Try `127.0.0.1` instead of `localhost`

#### Issue: Browser shows "This site can't be reached"

**Solution:**
1. Verify server is running (check terminal)
2. Try `http://127.0.0.1:5000` instead
3. Check if port is correct
4. Try disabling browser extensions

### Step-by-Step Debug

1. **Start the server:**
   ```bash
   python3 web_app.py
   ```

2. **Check for errors:**
   - Look for "Starting web server..." message
   - Check for any error messages

3. **Try the URL:**
   - First try: `http://127.0.0.1:5000`
   - If that doesn't work, try: `http://localhost:5000`

4. **Check terminal output:**
   - You should see Flask debug messages when you access the page
   - If you don't see anything, the request isn't reaching the server

5. **Test with curl:**
   ```bash
   curl http://127.0.0.1:5000
   ```
   - If this works, the server is fine, it's a browser issue
   - If this fails, the server isn't running properly

### Alternative: Use Different Port

If port 5000 has issues, use port 8080:

```bash
python3 web_app.py 8080
```

Then open: `http://127.0.0.1:8080`

### Still Not Working?

1. **Check Python/Flask installation:**
   ```bash
   python3 -c "import flask; print(flask.__version__)"
   ```

2. **Check if port is accessible:**
   ```bash
   netstat -an | grep 5000
   ```

3. **Try running with verbose output:**
   - The server should show Flask debug messages
   - Check terminal for any errors

4. **Check macOS security:**
   - System Settings > Privacy & Security
   - Allow Python if prompted

### Quick Test

Run this to test if Flask works:

```bash
python3 -c "from flask import Flask; app = Flask(__name__); app.run(port=5000)"
```

Then try accessing `http://127.0.0.1:5000` in browser.

If this works, the issue is with `web_app.py`. If it doesn't, it's a Flask/system issue.

