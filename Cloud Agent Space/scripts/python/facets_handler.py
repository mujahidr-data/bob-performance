#!/usr/bin/env python3
"""
Facets Cloud Platform Handler
Wrapper function for HiBob automation to work on Facets serverless platform
"""

import os
import json
import base64
from pathlib import Path
import sys

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from hibob_report_downloader import HiBobReportDownloader


def handler(event, context=None):
    """
    Facets function handler
    
    Expected event format:
    {
        "report_name": "Q2&Q3 Performance Check In",
        "config": {
            "email": "user@example.com",  # Optional if in env vars
            "password": "password"         # Optional if in env vars
        }
    }
    
    Returns:
    {
        "statusCode": 200,
        "body": {
            "success": true,
            "message": "Report downloaded and uploaded successfully"
        }
    }
    """
    try:
        # Parse event (could be dict or string)
        if isinstance(event, str):
            event = json.loads(event)
        
        # Get report name
        report_name = event.get('report_name')
        if not report_name:
            return {
                'statusCode': 400,
                'body': json.dumps({
                    'success': False,
                    'error': 'report_name is required in event'
                })
            }
        
        # Build config from environment variables or event
        config = event.get('config', {})
        
        # Get credentials from environment or event
        email = config.get('email') or os.environ.get('HIBOB_EMAIL')
        password = config.get('password') or os.environ.get('HIBOB_PASSWORD')
        
        if not email or not password:
            return {
                'statusCode': 400,
                'body': json.dumps({
                    'success': False,
                    'error': 'HiBob credentials (email/password) required. Set HIBOB_EMAIL and HIBOB_PASSWORD environment variables or provide in event.config'
                })
            }
        
        # Handle service account JSON
        # Option 1: From environment variable (base64 or JSON string)
        service_account_json = os.environ.get('SERVICE_ACCOUNT_JSON')
        service_account_path = None
        
        if service_account_json:
            try:
                # Try to decode as base64 first
                try:
                    decoded = base64.b64decode(service_account_json)
                    service_account_data = json.loads(decoded)
                except:
                    # If not base64, try as JSON string
                    service_account_data = json.loads(service_account_json)
                
                # Write to temporary file (serverless-friendly location)
                service_account_path = Path('/tmp/service_account.json')
                with open(service_account_path, 'w') as f:
                    json.dump(service_account_data, f)
                
                print(f"✓ Service account loaded from environment variable")
            except Exception as e:
                print(f"⚠️  Error loading service account from env: {str(e)}")
        
        # Option 2: Check if service_account.json exists in working directory
        if not service_account_path or not service_account_path.exists():
            local_service_account = Path('service_account.json')
            if local_service_account.exists():
                service_account_path = local_service_account
                print(f"✓ Using local service_account.json")
        
        # Build config dict for HiBobReportDownloader
        hibob_config = {
            'email': email,
            'password': password,
            'apps_script_url': os.environ.get('APPS_SCRIPT_URL', ''),  # Not needed for Google Sheets API
            'service_account_path': str(service_account_path) if service_account_path else None
        }
        
        # Create config file for downloader (it expects a file)
        # Write to temp location
        config_path = Path('/tmp/config.json')
        with open(config_path, 'w') as f:
            json.dump(hibob_config, f)
        
        # Initialize downloader
        print(f"🚀 Starting HiBob automation for report: {report_name}")
        downloader = HiBobReportDownloader(str(config_path))
        
        # Override service account path if provided
        if service_account_path and service_account_path.exists():
            # Temporarily copy service account to expected location
            import shutil
            target_path = Path('service_account.json')
            shutil.copy(service_account_path, target_path)
        
        # Run automation
        result = downloader.run(report_name)
        
        # Cleanup
        if service_account_path and Path('service_account.json').exists():
            # Only delete if we created it
            if service_account_path != Path('service_account.json'):
                try:
                    Path('service_account.json').unlink()
                except:
                    pass
        
        if result:
            return {
                'statusCode': 200,
                'body': json.dumps({
                    'success': True,
                    'message': f'Report "{report_name}" downloaded and uploaded to Google Sheets successfully',
                    'report_name': report_name
                })
            }
        else:
            return {
                'statusCode': 500,
                'body': json.dumps({
                    'success': False,
                    'error': 'Automation failed. Check logs for details.',
                    'report_name': report_name
                })
            }
            
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        print(f"❌ Error in handler: {str(e)}")
        print(error_trace)
        
        return {
            'statusCode': 500,
            'body': json.dumps({
                'success': False,
                'error': str(e),
                'traceback': error_trace
            })
        }


# For local testing
if __name__ == "__main__":
    # Test event
    test_event = {
        'report_name': 'Q2&Q3 Performance Check In',
        'config': {
            'email': os.environ.get('HIBOB_EMAIL', ''),
            'password': os.environ.get('HIBOB_PASSWORD', '')
        }
    }
    
    result = handler(test_event)
    print(json.dumps(result, indent=2))

