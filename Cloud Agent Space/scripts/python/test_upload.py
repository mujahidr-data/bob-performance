#!/usr/bin/env python3
"""
Test script to upload a downloaded Excel file to Google Sheets using Google Sheets API
"""

import json
import os
import sys
from pathlib import Path
import pandas as pd

def test_upload(filepath):
    """Test uploading an Excel file to Google Sheets using Google Sheets API"""
    
    print("=" * 60)
    print("TEST UPLOAD TO GOOGLE SHEETS")
    print("=" * 60)
    
    # Verify file exists
    if not os.path.exists(filepath):
        print(f"❌ File not found: {filepath}")
        return False
    
    # Get file size
    file_size = os.path.getsize(filepath)
    filename = os.path.basename(filepath)
    print(f"\n📄 File: {filename}")
    print(f"   Size: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)")
    
    # Read Excel file
    print(f"\n📖 Reading Excel file...")
    try:
        df = pd.read_excel(filepath, engine='openpyxl')
        print(f"   ✓ Read {len(df)} rows, {len(df.columns)} columns")
        print(f"   Columns: {', '.join(df.columns[:5].tolist())}{'...' if len(df.columns) > 5 else ''}")
    except Exception as e:
        print(f"❌ Error reading Excel file: {str(e)}")
        import traceback
        traceback.print_exc()
        return False
    
    # Convert to data array
    print(f"\n🔄 Converting to data array...")
    df = df.fillna('')
    data = [df.columns.tolist()] + df.values.tolist()
    print(f"   ✓ Prepared {len(data)} rows, {len(data[0]) if data else 0} columns")
    
    # Check payload size
    print(f"\n📦 Checking payload size...")
    payload_json = json.dumps({'data': data})
    payload_size = len(payload_json)
    print(f"   Payload size: {payload_size:,} bytes ({payload_size / 1024 / 1024:.2f} MB)")
    
    if payload_size > 50 * 1024 * 1024:
        print("   ⚠️  WARNING: Payload exceeds 50MB, may fail!")
    elif payload_size > 10 * 1024 * 1024:
        print("   ⚠️  WARNING: Payload is large (>10MB), may be slow")
    else:
        print("   ✓ Payload size is reasonable")
    
    # Use Google Sheets API
    print(f"\n☁️  Uploading to Google Sheets using Google Sheets API...")
    
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
    except ImportError:
        print(f"\n❌ Google API libraries not installed")
        print(f"   Install with: pip3 install google-auth google-api-python-client")
        return False
    
    # Google Sheets API scope
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    
    # Sheet ID from the URL
    SPREADSHEET_ID = '1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA'
    SHEET_NAME = 'Bob Perf Report'
    
    # Try service account first (better for automation)
    service_account_file = Path('service_account.json')
    creds = None
    
    if service_account_file.exists():
        print(f"   🔑 Using service account credentials...")
        try:
            creds = service_account.Credentials.from_service_account_file(
                str(service_account_file),
                scopes=SCOPES
            )
            print(f"   ✓ Service account credentials loaded")
        except Exception as e:
            print(f"   ⚠️  Error loading service account: {str(e)}")
            print(f"   💡 Make sure service_account.json is valid and the sheet is shared with the service account email")
            return False
    else:
        # Fallback to OAuth2 (for user credentials)
        print(f"   ⚠️  service_account.json not found, trying OAuth2...")
        try:
            from google.oauth2.credentials import Credentials
            from google_auth_oauthlib.flow import InstalledAppFlow
            from google.auth.transport.requests import Request
            import pickle
            
            token_file = Path('token.pickle')
            
            # Load existing token
            if token_file.exists():
                try:
                    with open(token_file, 'rb') as token:
                        creds = pickle.load(token)
                    print(f"   ✓ Loaded existing OAuth credentials")
                except:
                    print(f"   ⚠️  Could not load existing credentials")
            
            # If no valid credentials, get new ones
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    print(f"   🔄 Refreshing OAuth credentials...")
                    creds.refresh(Request())
                else:
                    print(f"   🔑 Requesting new OAuth credentials...")
                    print(f"   📋 A browser window will open for authentication")
                    
                    # Check for credentials file
                    creds_file = Path('credentials.json')
                    if not creds_file.exists():
                        print(f"\n   ❌ credentials.json not found!")
                        print(f"   💡 To set up Google Sheets API:")
                        print(f"      Option 1 (Recommended): Use Service Account")
                        print(f"      1. Go to: https://console.cloud.google.com/")
                        print(f"      2. Create a project or select existing")
                        print(f"      3. Enable Google Sheets API")
                        print(f"      4. Create Service Account")
                        print(f"      5. Download JSON key as service_account.json")
                        print(f"      6. Share the Google Sheet with the service account email")
                        print(f"      Option 2: Use OAuth2 (requires browser)")
                        print(f"      1. Create OAuth 2.0 credentials (Desktop app)")
                        print(f"      2. Download as credentials.json")
                        print(f"   📖 See GOOGLE_SHEETS_API_SETUP.md for detailed instructions")
                        return False
                    
                    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                    creds = flow.run_local_server(port=0)
                
                # Save credentials for next time
                with open(token_file, 'wb') as token:
                    pickle.dump(creds, token)
                print(f"   ✓ OAuth credentials saved")
        except ImportError:
            print(f"   ❌ OAuth2 libraries not available")
            print(f"   💡 Install with: pip install google-auth-oauthlib")
            return False
    
    # Build the service
    print(f"   🔗 Connecting to Google Sheets...")
    service = build('sheets', 'v4', credentials=creds)
    
    # Get or create the sheet
    print(f"   📋 Accessing sheet: {SHEET_NAME}")
    spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheet_names = [sheet['properties']['title'] for sheet in spreadsheet.get('sheets', [])]
    
    if SHEET_NAME not in sheet_names:
        print(f"   ➕ Creating new sheet: {SHEET_NAME}")
        request_body = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': SHEET_NAME
                    }
                }
            }]
        }
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body=request_body
        ).execute()
    else:
        print(f"   ✓ Found existing sheet: {SHEET_NAME}")
        # Clear existing data
        range_name = f"{SHEET_NAME}!A1:Z10000"
        service.spreadsheets().values().clear(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name
        ).execute()
    
    # Write data to sheet
    print(f"   ✍️  Writing {len(data)} rows to sheet...")
    range_name = f"{SHEET_NAME}!A1"
    
    body = {
        'values': data
    }
    
    result = service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_name,
        valueInputOption='RAW',
        body=body
    ).execute()
    
    updated_cells = result.get('updatedCells', 0)
    updated_rows = result.get('updatedRows', 0)
    
    # Format header row
    print(f"   🎨 Formatting header row...")
    sheet_id = None
    for sheet in spreadsheet.get('sheets', []):
        if sheet['properties']['title'] == SHEET_NAME:
            sheet_id = sheet['properties']['sheetId']
            break
    
    if sheet_id is not None:
        requests = [
            {
                'repeatCell': {
                    'range': {
                        'sheetId': sheet_id,
                        'startRowIndex': 0,
                        'endRowIndex': 1
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'backgroundColor': {'red': 0.26, 'green': 0.52, 'blue': 0.96},
                            'textFormat': {
                                'foregroundColor': {'red': 1.0, 'green': 1.0, 'blue': 1.0},
                                'bold': True
                            }
                        }
                    },
                    'fields': 'userEnteredFormat(backgroundColor,textFormat)'
                }
            },
            {
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': sheet_id,
                        'gridProperties': {
                            'frozenRowCount': 1
                        }
                    },
                    'fields': 'gridProperties.frozenRowCount'
                }
            }
        ]
        
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={'requests': requests}
        ).execute()
    
    print(f"\n✅ SUCCESS!")
    print(f"   📊 Updated {updated_rows} rows, {updated_cells} cells")
    print(f"   🔗 Sheet: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit#gid={sheet_id if sheet_id else 0}")
    
    return True

if __name__ == "__main__":
    # Find the most recent downloaded file
    downloads_dir = Path("downloads")
    if not downloads_dir.exists():
        print("❌ downloads directory not found")
        sys.exit(1)
    
    # Get all Excel files
    excel_files = list(downloads_dir.glob("*.xlsx"))
    if not excel_files:
        print("❌ No Excel files found in downloads directory")
        sys.exit(1)
    
    # Use the most recent file
    latest_file = max(excel_files, key=os.path.getctime)
    print(f"Using most recent file: {latest_file.name}\n")
    
    # Test upload
    try:
        success = test_upload(str(latest_file))
        
        if success:
            print("\n" + "=" * 60)
            print("✅ TEST PASSED - Upload successful!")
            print("=" * 60)
            sys.exit(0)
        else:
            print("\n" + "=" * 60)
            print("❌ TEST FAILED - Upload unsuccessful")
            print("=" * 60)
            sys.exit(1)
    except Exception as e:
        print(f"\n❌ TEST ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
