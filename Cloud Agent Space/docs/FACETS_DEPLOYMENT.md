# Deploying to Facets Cloud Platform

This guide explains how to deploy the HiBob automation to Facets for cloud-hosted execution.

## Prerequisites

1. **Facets Account**: Sign up at [facets.cloud](https://facets.cloud)
2. **Service Account JSON**: Already set up (`service_account.json`)
3. **Git Repository**: Code is already in Git

## Key Modifications Needed

### 1. Remove Interactive Input
The current script uses `input()` which won't work in serverless. We need to:
- Accept report name as function parameter
- Accept credentials via environment variables or function parameters

### 2. Handle Browser Automation
Playwright in serverless has challenges:
- Large binary size (may need custom buildpack)
- Cold start delays
- Memory limits

### 3. File Handling
- Downloads need to be in memory or temporary storage
- Can't rely on local filesystem

## Step-by-Step Deployment

### Step 1: Create Facets-Compatible Function

We need to create a wrapper function that:
- Accepts parameters (report_name, etc.)
- Runs the automation
- Returns results

### Step 2: Prepare Dependencies

**requirements.txt** (already exists, but verify):
```
playwright==1.40.0
requests==2.31.0
pandas==2.1.4
openpyxl==3.1.2
google-auth==2.25.2
google-api-python-client==2.108.0
```

**Note**: Playwright requires browser binaries. Facets may need:
- Custom buildpack with Chromium
- Or use Playwright's bundled browsers

### Step 3: Environment Variables

Set these in Facets project settings:
- `HIBOB_EMAIL`: Your HiBob email
- `HIBOB_PASSWORD`: Your HiBob password (or use secrets management)
- `GOOGLE_SHEET_ID`: `1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA`
- `SERVICE_ACCOUNT_JSON`: Base64-encoded service_account.json (or use Facets secrets)

### Step 4: Create Deployment Function

Create `facets_handler.py`:

```python
import os
import json
import base64
from pathlib import Path
from hibob_report_downloader import HiBobReportDownloader

def handler(event, context):
    """
    Facets function handler
    
    Expected event format:
    {
        "report_name": "Q2&Q3 Performance Check In",
        "config": {
            "email": "user@example.com",
            "password": "password",
            "google_sheet_id": "1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA"
        }
    }
    """
    try:
        # Parse event
        report_name = event.get('report_name')
        if not report_name:
            return {
                'statusCode': 400,
                'body': json.dumps({'error': 'report_name is required'})
            }
        
        # Load config from environment or event
        config = event.get('config', {})
        if not config.get('email'):
            config['email'] = os.environ.get('HIBOB_EMAIL')
        if not config.get('password'):
            config['password'] = os.environ.get('HIBOB_PASSWORD')
        
        # Handle service account
        service_account_json = os.environ.get('SERVICE_ACCOUNT_JSON')
        if service_account_json:
            # Decode if base64
            try:
                service_account_data = json.loads(base64.b64decode(service_account_json))
            except:
                service_account_data = json.loads(service_account_json)
            
            # Write to temporary file
            service_account_path = Path('/tmp/service_account.json')
            with open(service_account_path, 'w') as f:
                json.dump(service_account_data, f)
        
        # Initialize downloader
        downloader = HiBobReportDownloader(config)
        
        # Run automation
        result = downloader.run(report_name)
        
        if result:
            return {
                'statusCode': 200,
                'body': json.dumps({
                    'success': True,
                    'message': f'Report "{report_name}" downloaded and uploaded successfully'
                })
            }
        else:
            return {
                'statusCode': 500,
                'body': json.dumps({
                    'success': False,
                    'message': 'Automation failed'
                })
            }
            
    except Exception as e:
        return {
            'statusCode': 500,
            'body': json.dumps({
                'success': False,
                'error': str(e)
            })
        }
```

### Step 5: Modify HiBobReportDownloader for Serverless

We need to update `hibob_report_downloader.py` to:
1. Accept report_name as parameter (not input)
2. Use headless browser
3. Handle temporary file storage

Key changes needed:
- Remove `input()` calls
- Use `headless=True` for browser
- Use `/tmp` for downloads (serverless-friendly)

### Step 6: Deploy to Facets

1. **Create Project in Facets**:
   - Log in to Facets dashboard
   - Create new project
   - Select Python runtime

2. **Upload Code**:
   - Upload all Python files
   - Upload `requirements.txt`
   - Upload `service_account.json` (or set as secret)

3. **Configure Environment Variables**:
   - Set `HIBOB_EMAIL`
   - Set `HIBOB_PASSWORD` (use secrets)
   - Set `GOOGLE_SHEET_ID`
   - Set `SERVICE_ACCOUNT_JSON` (if not uploading file)

4. **Set Function Handler**:
   - Entry point: `facets_handler.handler`
   - Timeout: 15-30 minutes (browser automation takes time)
   - Memory: 2-4 GB (Playwright needs memory)

5. **Install Playwright Browsers**:
   - May need custom build step:
   ```bash
   playwright install chromium
   playwright install-deps chromium
   ```

### Step 7: Configure Triggers

Facets supports various triggers:
- **HTTP Trigger**: Call via POST request
- **Scheduled Trigger**: Run on cron schedule
- **Event Trigger**: Trigger from other services

**HTTP Trigger Example**:
```bash
curl -X POST https://your-function.facets.cloud \
  -H "Content-Type: application/json" \
  -d '{
    "report_name": "Q2&Q3 Performance Check In"
  }'
```

## Challenges & Solutions

### Challenge 1: Playwright Browser Binaries
**Problem**: Playwright requires Chromium (~170MB)
**Solution**: 
- Use Facets custom buildpack
- Or use Playwright's bundled browser installation
- Consider using Playwright's Docker image

### Challenge 2: Cold Start Times
**Problem**: First invocation can be slow (30+ seconds)
**Solution**:
- Use provisioned concurrency (if available)
- Keep function warm with scheduled pings
- Consider dedicated instance instead of serverless

### Challenge 3: Execution Time Limits
**Problem**: Browser automation can take 5-15 minutes
**Solution**:
- Request extended timeout from Facets
- Optimize automation (reduce waits)
- Consider breaking into smaller functions

### Challenge 4: Memory Limits
**Problem**: Playwright + browser needs 1-2GB RAM
**Solution**:
- Request higher memory tier
- Use headless mode (reduces memory)
- Close browser immediately after use

## Alternative: Use Existing Web App

If Facets doesn't support Playwright well, consider:

1. **Deploy web_app.py to Facets**:
   - Flask app is easier to host
   - Can trigger automation via API
   - Still needs Playwright, but more control

2. **Hybrid Approach**:
   - Keep browser automation on cloud VM (Railway, Render)
   - Use Facets for lightweight API endpoints
   - Facets calls VM endpoint

## Recommended Approach

### Option A: Full Facets Deployment (If Supported)
1. Create `facets_handler.py`
2. Modify `hibob_report_downloader.py` for serverless
3. Deploy with custom Playwright buildpack
4. Test with small reports first

### Option B: Hybrid (Recommended)
1. Deploy `web_app.py` to cloud VM (Railway/Render)
2. Create lightweight Facets function that calls VM
3. Facets function handles authentication/triggering
4. Actual automation runs on VM

### Option C: Use Facets for API Only
1. Deploy simple API to Facets
2. API triggers automation on separate service
3. Best of both worlds

## Testing Locally

Before deploying, test the handler locally:

```python
# test_facets_handler.py
from facets_handler import handler

event = {
    'report_name': 'Q2&Q3 Performance Check In',
    'config': {
        'email': 'your-email@example.com',
        'password': 'your-password'
    }
}

result = handler(event, None)
print(result)
```

## Next Steps

1. **Check Facets Documentation**: Verify Playwright support
2. **Create facets_handler.py**: Function wrapper
3. **Modify hibob_report_downloader.py**: Remove interactive input
4. **Test Locally**: Ensure it works without input()
5. **Deploy to Facets**: Follow Facets deployment guide
6. **Monitor**: Check logs and performance

## Questions to Ask Facets Support

1. Does Facets support Playwright/Chromium?
2. What's the maximum execution time?
3. What's the maximum memory allocation?
4. Can we use custom buildpacks?
5. How to handle large dependencies (browser binaries)?
6. Is there a Docker/container option?

---

**Note**: If Facets doesn't support Playwright well, consider Railway.app, Render.com, or Google Cloud Run which have better support for browser automation.

