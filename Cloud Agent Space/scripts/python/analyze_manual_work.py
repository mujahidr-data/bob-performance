#!/usr/bin/env python3
"""
Analyze Google Sheet to understand manual processes that can be automated
"""

import json
from pathlib import Path
from google.oauth2 import service_account
from googleapiclient.discovery import build
import pandas as pd

SPREADSHEET_ID = '1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA'

def analyze_sheet():
    """Analyze sheet structure and identify automation opportunities"""
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
        
        print("📊 Analyzing Google Sheet for Automation Opportunities")
        print("=" * 70)
        
        # Key sheets to analyze
        key_sheets = {
            'Bob Perf Report': 'Current automated upload destination',
            'Paste Bob Perf Report Here': 'Manual staging area (can be automated)',
            'Summary': 'Aggregated data (can be auto-generated)',
            'H1 2025': 'Review cycle data (can be auto-populated)',
            'AYR 2024': 'Annual review data (can be auto-populated)',
            'Lookup': 'Reference tables (can be auto-updated)'
        }
        
        analysis = {}
        
        for sheet_name in key_sheets.keys():
            print(f"\n📋 Analyzing: {sheet_name}")
            print("-" * 70)
            
            try:
                # Get all data
                result = service.spreadsheets().values().get(
                    spreadsheetId=SPREADSHEET_ID,
                    range=f"{sheet_name}!A1:Z1000"
                ).execute()
                
                values = result.get('values', [])
                if not values:
                    print(f"   ⚠️  No data found")
                    continue
                
                # Convert to DataFrame for analysis
                headers = values[0]
                data_rows = values[1:] if len(values) > 1 else []
                
                print(f"   Headers ({len(headers)}): {', '.join(headers[:10])}{'...' if len(headers) > 10 else ''}")
                print(f"   Data rows: {len(data_rows)}")
                
                # Analyze structure
                analysis[sheet_name] = {
                    'headers': headers,
                    'row_count': len(data_rows),
                    'column_count': len(headers),
                    'sample_rows': data_rows[:3] if len(data_rows) >= 3 else data_rows
                }
                
                # Identify key columns
                key_columns = []
                for col in headers:
                    col_lower = col.lower()
                    if any(keyword in col_lower for keyword in ['email', 'name', 'id', 'rating', 'manager', 'department', 'site']):
                        key_columns.append(col)
                
                if key_columns:
                    print(f"   Key columns: {', '.join(key_columns)}")
                
            except Exception as e:
                print(f"   ❌ Error: {str(e)}")
        
        # Save analysis
        output_file = Path('sheet_analysis.json')
        with open(output_file, 'w') as f:
            json.dump(analysis, f, indent=2, default=str)
        
        print("\n" + "=" * 70)
        print("✅ Analysis complete!")
        print(f"📄 Detailed analysis saved to: {output_file}")
        
        # Suggest automation opportunities
        print("\n" + "=" * 70)
        print("💡 Automation Opportunities:")
        print("=" * 70)
        
        suggestions = [
            "1. Auto-populate 'Summary' sheet from 'Bob Perf Report' data",
            "2. Auto-update 'H1 2025' and 'AYR 2024' sheets with latest performance data",
            "3. Auto-sync 'Lookup' tables based on performance ratings",
            "4. Eliminate 'Paste Bob Perf Report Here' (use 'Bob Perf Report' directly)",
            "5. Auto-generate reports/analytics from performance data",
            "6. Auto-update manager hierarchies and department mappings",
            "7. Auto-calculate metrics (promotion readiness, rating distributions, etc.)"
        ]
        
        for suggestion in suggestions:
            print(f"   {suggestion}")
        
        return analysis
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    analyze_sheet()

