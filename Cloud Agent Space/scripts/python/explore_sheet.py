#!/usr/bin/env python3
"""
Explore Google Sheet structure to understand existing worksheets and data
"""

import json
from pathlib import Path
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Sheet ID
SPREADSHEET_ID = '1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA'

def explore_sheet():
    """Explore the Google Sheet structure"""
    try:
        # Load service account
        service_account_file = Path('service_account.json')
        if not service_account_file.exists():
            print("❌ service_account.json not found")
            print("   Please ensure service_account.json is in the current directory")
            return
        
        print("🔐 Authenticating with Google Sheets API...")
        creds = service_account.Credentials.from_service_account_file(
            str(service_account_file),
            scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
        )
        
        print("🔗 Connecting to Google Sheets...")
        service = build('sheets', 'v4', credentials=creds)
        
        # Get spreadsheet metadata
        print(f"\n📊 Exploring Sheet: {SPREADSHEET_ID}")
        print("=" * 60)
        
        spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        
        # Get all sheets
        sheets = spreadsheet.get('sheets', [])
        print(f"\n📋 Found {len(sheets)} worksheet(s):\n")
        
        for i, sheet in enumerate(sheets, 1):
            props = sheet['properties']
            sheet_id = props['sheetId']
            title = props['title']
            row_count = props.get('gridProperties', {}).get('rowCount', 0)
            col_count = props.get('gridProperties', {}).get('columnCount', 0)
            
            print(f"{i}. {title}")
            print(f"   Sheet ID: {sheet_id}")
            print(f"   Dimensions: {row_count} rows × {col_count} columns")
            
            # Get a sample of data (first 10 rows)
            try:
                result = service.spreadsheets().values().get(
                    spreadsheetId=SPREADSHEET_ID,
                    range=f"{title}!A1:Z10"
                ).execute()
                
                values = result.get('values', [])
                if values:
                    print(f"   Sample data ({len(values)} rows):")
                    # Show headers
                    if len(values) > 0:
                        headers = values[0]
                        print(f"   Headers: {', '.join(headers[:10])}{'...' if len(headers) > 10 else ''}")
                    # Show a few data rows
                    for row_idx, row in enumerate(values[1:4], 1):
                        if row:
                            print(f"   Row {row_idx}: {', '.join(str(cell)[:20] for cell in row[:5])}{'...' if len(row) > 5 else ''}")
                else:
                    print("   (No data found)")
            except Exception as e:
                print(f"   ⚠️  Could not read data: {str(e)}")
            
            print()
        
        # Get full sheet structure
        print("\n" + "=" * 60)
        print("📝 Full Sheet Structure:")
        print("=" * 60)
        
        sheet_structure = {
            'spreadsheet_id': SPREADSHEET_ID,
            'spreadsheet_title': spreadsheet.get('properties', {}).get('title', 'Unknown'),
            'sheets': []
        }
        
        for sheet in sheets:
            props = sheet['properties']
            sheet_info = {
                'title': props['title'],
                'sheet_id': props['sheetId'],
                'index': props['index'],
                'row_count': props.get('gridProperties', {}).get('rowCount', 0),
                'column_count': props.get('gridProperties', {}).get('columnCount', 0),
                'hidden': props.get('hidden', False),
                'tab_color': props.get('tabColor', {})
            }
            
            # Try to get column headers
            try:
                result = service.spreadsheets().values().get(
                    spreadsheetId=SPREADSHEET_ID,
                    range=f"{props['title']}!A1:Z1"
                ).execute()
                values = result.get('values', [])
                if values and len(values) > 0:
                    sheet_info['headers'] = values[0]
            except:
                pass
            
            sheet_structure['sheets'].append(sheet_info)
        
        # Save structure to file
        output_file = Path('sheet_structure.json')
        with open(output_file, 'w') as f:
            json.dump(sheet_structure, f, indent=2)
        
        print(f"\n✅ Sheet structure saved to: {output_file}")
        print("\n📋 Summary:")
        print(f"   Total worksheets: {len(sheets)}")
        print(f"   Sheet titles: {', '.join([s['properties']['title'] for s in sheets])}")
        
        return sheet_structure
        
    except Exception as e:
        print(f"\n❌ Error exploring sheet: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    explore_sheet()

