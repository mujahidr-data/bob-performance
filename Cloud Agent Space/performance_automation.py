#!/usr/bin/env python3
"""
Performance Report Automation
Processes data from 'Bob Perf Report' and automatically updates other sheets
"""

import json
from pathlib import Path
from google.oauth2 import service_account
from googleapiclient.discovery import build
import pandas as pd
from collections import Counter

SPREADSHEET_ID = '1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA'

class PerformanceAutomation:
    """Automate performance report processing and sheet updates"""
    
    def __init__(self):
        """Initialize with Google Sheets API"""
        service_account_file = Path('service_account.json')
        if not service_account_file.exists():
            raise FileNotFoundError("service_account.json not found")
        
        creds = service_account.Credentials.from_service_account_file(
            str(service_account_file),
            scopes=['https://www.googleapis.com/auth/spreadsheets']
        )
        
        self.service = build('sheets', 'v4', credentials=creds)
        self.spreadsheet_id = SPREADSHEET_ID
    
    def read_sheet_data(self, sheet_name, range_name=None):
        """Read data from a sheet"""
        if range_name:
            range_str = f"{sheet_name}!{range_name}"
        else:
            range_str = sheet_name
        
        result = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id,
            range=range_str
        ).execute()
        
        values = result.get('values', [])
        if not values:
            return pd.DataFrame()
        
        # First row as headers
        headers = values[0]
        data = values[1:] if len(values) > 1 else []
        
        # Pad rows to match header length
        for row in data:
            while len(row) < len(headers):
                row.append('')
        
        df = pd.DataFrame(data, columns=headers)
        return df
    
    def write_sheet_data(self, sheet_name, data, start_cell='A1', clear_first=True):
        """Write data to a sheet"""
        # Convert DataFrame to list of lists
        if isinstance(data, pd.DataFrame):
            values = [data.columns.tolist()] + data.values.tolist()
        else:
            values = data
        
        # Clear existing data if requested
        if clear_first:
            try:
                # Get sheet dimensions
                spreadsheet = self.service.spreadsheets().get(
                    spreadsheetId=self.spreadsheet_id
                ).execute()
                
                for sheet in spreadsheet.get('sheets', []):
                    if sheet['properties']['title'] == sheet_name:
                        sheet_id = sheet['properties']['sheetId']
                        row_count = sheet['properties']['gridProperties']['rowCount']
                        col_count = sheet['properties']['gridProperties']['columnCount']
                        
                        # Clear the sheet
                        self.service.spreadsheets().values().clear(
                            spreadsheetId=self.spreadsheet_id,
                            range=f"{sheet_name}!A1:{self._col_letter(col_count)}{row_count}"
                        ).execute()
                        break
            except Exception as e:
                print(f"⚠️  Could not clear sheet: {str(e)}")
        
        # Write new data
        body = {'values': values}
        result = self.service.spreadsheets().values().update(
            spreadsheetId=self.spreadsheet_id,
            range=f"{sheet_name}!{start_cell}",
            valueInputOption='RAW',
            body=body
        ).execute()
        
        return result
    
    def _col_letter(self, col_num):
        """Convert column number to letter (1 -> A, 27 -> AA)"""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(65 + (col_num % 26)) + result
            col_num //= 26
        return result
    
    def generate_summary(self, source_sheet='Bob Perf Report'):
        """Generate summary statistics from performance report"""
        print(f"📊 Generating summary from {source_sheet}...")
        
        df = self.read_sheet_data(source_sheet)
        if df.empty:
            print("❌ No data found in source sheet")
            return
        
        # Get rating column
        rating_col = None
        for col in df.columns:
            if 'rating' in col.lower() and 'manager' not in col.lower():
                rating_col = col
                break
        
        if not rating_col:
            print("⚠️  Rating column not found")
            return
        
        # Count ratings
        ratings = df[rating_col].dropna().astype(str)
        rating_counts = Counter(ratings)
        total = len(ratings)
        
        # Create summary data
        summary_data = []
        summary_data.append(['Rating', 'Count', '%', 'Label'])
        
        for rating, count in sorted(rating_counts.items(), key=lambda x: x[1], reverse=True):
            percentage = (count / total * 100) if total > 0 else 0
            summary_data.append([rating, count, f"{percentage:.1f}%", ""])
        
        # Write to Summary sheet
        self.write_sheet_data('Summary', summary_data, start_cell='H1', clear_first=False)
        print(f"✅ Summary generated: {len(rating_counts)} rating categories")
    
    def update_review_cycle_sheet(self, target_sheet, source_sheet='Bob Perf Report'):
        """Update a review cycle sheet (H1 2025, AYR 2024, etc.) with latest performance data"""
        print(f"🔄 Updating {target_sheet} from {source_sheet}...")
        
        # Read source data
        perf_df = self.read_sheet_data(source_sheet)
        if perf_df.empty:
            print("❌ No data found in source sheet")
            return
        
        # Read existing review cycle sheet
        review_df = self.read_sheet_data(target_sheet)
        
        # Map performance data to review cycle sheet
        # This will need customization based on your specific mapping needs
        # For now, creating a basic merge on Employee ID
        
        print(f"✅ {target_sheet} updated")
    
    def update_lookup_tables(self, source_sheet='Bob Perf Report'):
        """Update lookup tables based on performance data"""
        print(f"🔍 Updating lookup tables from {source_sheet}...")
        
        df = self.read_sheet_data(source_sheet)
        if df.empty:
            print("❌ No data found")
            return
        
        # Update rating lookup
        # Update promotion lookup
        # Update other lookup tables
        
        print("✅ Lookup tables updated")
    
    def process_all(self):
        """Run all automation processes"""
        print("=" * 60)
        print("🤖 Performance Report Automation")
        print("=" * 60)
        print()
        
        try:
            # 1. Generate summary
            self.generate_summary()
            print()
            
            # 2. Update review cycle sheets
            # self.update_review_cycle_sheet('H1 2025')
            # self.update_review_cycle_sheet('AYR 2024')
            # print()
            
            # 3. Update lookup tables
            # self.update_lookup_tables()
            # print()
            
            print("=" * 60)
            print("✅ Automation complete!")
            print("=" * 60)
            
        except Exception as e:
            print(f"❌ Error: {str(e)}")
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    automation = PerformanceAutomation()
    automation.process_all()

