#!/usr/bin/env python3
"""
Verify Google Sheet structure matches expected structure from Git/codebase
"""

import json
from pathlib import Path
from google.oauth2 import service_account
from googleapiclient.discovery import build

SPREADSHEET_ID = '1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA'

# Expected sheet names from codebase (bob-performance-module.gs)
EXPECTED_SHEETS = {
    'Base Data': 'BASE_DATA',
    'Bonus History': 'BONUS_HISTORY',
    'Comp History': 'COMP_HISTORY',
    'Full Comp History': 'FULL_COMP_HISTORY',
    'Bob Perf Report': 'PERF_REPORT',
    'Uploader': 'UPLOADER',
    'Bob Fields Meta Data': 'BOB_FIELDS_META',
    'Bob Lists': 'BOB_LISTS',
    'Employees': 'EMPLOYEES',
    'History Uploader': 'HISTORY_UPLOADER',
    'Bob Updater Guide': 'GUIDE'
}

def verify_sheet_structure():
    """Compare current sheet structure with expected structure"""
    try:
        # Load service account
        service_account_file = Path('service_account.json')
        if not service_account_file.exists():
            print("❌ service_account.json not found")
            return
        
        creds = service_account.Credentials.from_service_account_file(
            str(service_account_file),
            scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
        )
        
        service = build('sheets', 'v4', credentials=creds)
        
        print("=" * 70)
        print("📋 VERIFYING GOOGLE SHEET STRUCTURE")
        print("=" * 70)
        print()
        
        # Get current sheets
        spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        current_sheets = {}
        for sheet in spreadsheet.get('sheets', []):
            title = sheet['properties']['title']
            current_sheets[title] = {
                'sheet_id': sheet['properties']['sheetId'],
                'hidden': sheet['properties'].get('hidden', False),
                'index': sheet['properties']['index']
            }
        
        print(f"📊 Current Sheet Count: {len(current_sheets)}")
        print(f"📊 Expected Sheet Count: {len(EXPECTED_SHEETS)}")
        print()
        
        # Compare
        print("=" * 70)
        print("✅ SHEETS FOUND (Expected):")
        print("=" * 70)
        found_sheets = []
        for sheet_name, code_name in EXPECTED_SHEETS.items():
            if sheet_name in current_sheets:
                status = "✅"
                hidden = " (hidden)" if current_sheets[sheet_name]['hidden'] else ""
                print(f"{status} {sheet_name:30} [{code_name:20}]{hidden}")
                found_sheets.append(sheet_name)
            else:
                print(f"❌ {sheet_name:30} [{code_name:20}] - MISSING")
        
        print()
        print("=" * 70)
        print("⚠️  SHEETS FOUND (Not in Code):")
        print("=" * 70)
        unexpected_sheets = []
        for sheet_name in sorted(current_sheets.keys()):
            if sheet_name not in EXPECTED_SHEETS:
                hidden = " (hidden)" if current_sheets[sheet_name]['hidden'] else ""
                print(f"   {sheet_name:30}{hidden}")
                unexpected_sheets.append(sheet_name)
        
        if not unexpected_sheets:
            print("   (None - all sheets are expected)")
        
        print()
        print("=" * 70)
        print("📊 SUMMARY:")
        print("=" * 70)
        print(f"   Total sheets in Google Sheet: {len(current_sheets)}")
        print(f"   Expected sheets found: {len(found_sheets)}/{len(EXPECTED_SHEETS)}")
        print(f"   Unexpected sheets: {len(unexpected_sheets)}")
        print(f"   Missing expected sheets: {len(EXPECTED_SHEETS) - len(found_sheets)}")
        
        # List missing sheets
        missing_sheets = [name for name in EXPECTED_SHEETS.keys() if name not in found_sheets]
        if missing_sheets:
            print()
            print("❌ MISSING EXPECTED SHEETS:")
            for sheet in missing_sheets:
                print(f"   - {sheet} ({EXPECTED_SHEETS[sheet]})")
        
        # Additional analysis
        print()
        print("=" * 70)
        print("💡 ANALYSIS:")
        print("=" * 70)
        
        if missing_sheets:
            print("⚠️  Some expected sheets are missing.")
            print("   These may be created automatically when needed, or")
            print("   may have been deleted/renamed in another project.")
        else:
            print("✅ All expected sheets are present!")
        
        if unexpected_sheets:
            print()
            print("ℹ️  Additional sheets found that aren't in the code:")
            print("   These are likely manual work sheets or from other projects.")
            print("   They won't affect the automation, but you may want to:")
            print("   - Keep them if they're needed for manual processes")
            print("   - Document them if they're part of your workflow")
            print("   - Consider automating them if they're frequently updated")
        
        # Save comparison report
        comparison = {
            'spreadsheet_id': SPREADSHEET_ID,
            'verification_date': str(Path('sheet_structure.json').stat().st_mtime) if Path('sheet_structure.json').exists() else 'unknown',
            'current_sheets': list(current_sheets.keys()),
            'expected_sheets': list(EXPECTED_SHEETS.keys()),
            'found_sheets': found_sheets,
            'missing_sheets': missing_sheets,
            'unexpected_sheets': unexpected_sheets,
            'status': 'match' if not missing_sheets else 'mismatch'
        }
        
        output_file = Path('sheet_verification.json')
        with open(output_file, 'w') as f:
            json.dump(comparison, f, indent=2)
        
        print()
        print(f"📄 Verification report saved to: {output_file}")
        
        return comparison
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    verify_sheet_structure()

