#!/usr/bin/env python3
"""
Analyze Manager Blurb Generation Failures
Helps identify why blurbs are being flagged for manual review
"""

import json
import sys
import os
from pathlib import Path

# Add scripts/python to path
PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "scripts" / "python"))

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
except ImportError:
    print("❌ Missing Google API libraries")
    print("   Install with: pip install google-auth google-api-python-client")
    sys.exit(1)


def analyze_failures():
    """Analyze the Manager Blurbs sheet to understand failure patterns"""
    
    # Load config
    config_path = PROJECT_ROOT / "config" / "config.json"
    with open(config_path, 'r') as f:
        config = json.load(f)
    
    spreadsheet_id = config.get('spreadsheet_id')
    
    # Authenticate
    service_account_paths = [
        PROJECT_ROOT / 'config' / 'service_account.json',
    ]
    
    service_account_file = None
    for path in service_account_paths:
        if path.exists():
            service_account_file = str(path)
            break
    
    if not service_account_file:
        print("❌ service_account.json not found")
        sys.exit(1)
    
    creds = service_account.Credentials.from_service_account_file(
        service_account_file,
        scopes=['https://www.googleapis.com/auth/spreadsheets']
    )
    
    service = build('sheets', 'v4', credentials=creds)
    
    print("📊 Analyzing Manager Blurbs Sheet...")
    print()
    
    # Read Manager Blurbs sheet
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range="'Manager Blurbs'!A:B"
    ).execute()
    
    data = result.get('values', [])
    
    if len(data) < 2:
        print("❌ No data found in Manager Blurbs sheet")
        return
    
    headers = data[0]
    rows = data[1:]
    
    # Categorize blurbs
    valid_blurbs = []
    manual_review = []
    no_feedback = []
    
    for row in rows:
        if len(row) < 2:
            continue
        
        emp_id = row[0]
        blurb = row[1]
        
        if "manual review" in blurb.lower():
            manual_review.append((emp_id, blurb))
        elif "no meaningful feedback" in blurb.lower() or "no feedback available" in blurb.lower():
            no_feedback.append((emp_id, blurb))
        else:
            valid_blurbs.append((emp_id, blurb))
    
    total = len(rows)
    
    # Print summary
    print(f"📈 Summary Statistics")
    print(f"{'='*60}")
    print(f"Total Employees:           {total}")
    print(f"✅ Valid Blurbs:           {len(valid_blurbs):3d} ({len(valid_blurbs)/total*100:.1f}%)")
    print(f"⚠️  Manual Review:          {len(manual_review):3d} ({len(manual_review)/total*100:.1f}%)")
    print(f"ℹ️  No Feedback:            {len(no_feedback):3d} ({len(no_feedback)/total*100:.1f}%)")
    print()
    
    # Now read Bob Perf Report to see what feedback they have
    print("🔍 Analyzing Feedback Quality in Bob Perf Report...")
    print()
    
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range="'Bob Perf Report'!A1:Z1000"
    ).execute()
    
    perf_data = result.get('values', [])
    
    if not perf_data:
        print("❌ No data in Bob Perf Report")
        return
    
    perf_headers = perf_data[0]
    perf_rows = perf_data[1:]
    
    # Find feedback columns
    feedback_cols = {}
    for i, header in enumerate(perf_headers):
        header_lower = header.lower()
        if 'leadership principle' in header_lower and 'exemplified' in header_lower:
            feedback_cols['leadership_strength'] = i
        elif 'leadership principle' in header_lower and 'improvement' in header_lower:
            feedback_cols['leadership_improvement'] = i
        elif 'leveraged ai' in header_lower or 'ai, automation' in header_lower:
            feedback_cols['ai_leverage'] = i
        elif 'ai has not yet' in header_lower or 'openness' in header_lower:
            feedback_cols['ai_readiness'] = i
        elif 'support, coaching' in header_lower:
            feedback_cols['support_needed'] = i
        elif 'comment' in header_lower and 'performed against' in header_lower:
            feedback_cols['performance_comment'] = i
        elif 'comment' in header_lower and 'potential' in header_lower:
            feedback_cols['potential_comment'] = i
        elif 'comment' in header_lower and 'promoted' in header_lower:
            feedback_cols['promotion_comment'] = i
        elif header_lower == 'employee' or 'employee id' in header_lower:
            feedback_cols['employee_id'] = i
    
    print(f"Found {len(feedback_cols)} feedback columns:")
    for key, idx in feedback_cols.items():
        if key != 'employee_id':
            print(f"  - {key}: Column {idx} ({perf_headers[idx][:50]}...)")
    print()
    
    # Sample employees flagged for manual review
    print(f"🔎 Sample Employees Flagged for Manual Review (first 10):")
    print(f"{'='*60}")
    
    # Create emp_id to row mapping
    emp_id_col = feedback_cols.get('employee_id', 0)
    perf_map = {}
    for row in perf_rows:
        if len(row) > emp_id_col:
            perf_map[str(row[emp_id_col]).strip()] = row
    
    for i, (emp_id, blurb) in enumerate(manual_review[:10]):
        print(f"\n{i+1}. Employee ID: {emp_id}")
        
        if emp_id in perf_map:
            perf_row = perf_map[emp_id]
            print(f"   Feedback found in Bob Perf Report:")
            
            for col_name, col_idx in feedback_cols.items():
                if col_name == 'employee_id':
                    continue
                
                if col_idx < len(perf_row):
                    value = str(perf_row[col_idx]).strip()
                    if value and value not in ['', 'N/A', 'n/a', '-', 'None']:
                        print(f"   - {col_name}: {value[:80]}{'...' if len(value) > 80 else ''}")
        else:
            print(f"   ❌ No data found in Bob Perf Report for this employee")
    
    print()
    print(f"{'='*60}")
    print()
    
    # Analyze feedback completeness
    print("📊 Feedback Completeness Analysis:")
    print(f"{'='*60}")
    
    feedback_stats = {col: 0 for col in feedback_cols.keys() if col != 'employee_id'}
    total_checked = 0
    
    for emp_id, blurb in manual_review[:50]:  # Check first 50
        if emp_id in perf_map:
            total_checked += 1
            perf_row = perf_map[emp_id]
            
            for col_name, col_idx in feedback_cols.items():
                if col_name == 'employee_id':
                    continue
                
                if col_idx < len(perf_row):
                    value = str(perf_row[col_idx]).strip()
                    if value and value not in ['', 'N/A', 'n/a', '-', 'None']:
                        feedback_stats[col_name] += 1
    
    if total_checked > 0:
        print(f"Analyzed {total_checked} employees flagged for manual review:\n")
        for col_name, count in sorted(feedback_stats.items(), key=lambda x: x[1], reverse=True):
            pct = count / total_checked * 100
            print(f"  {col_name:25s}: {count:3d}/{total_checked:3d} ({pct:5.1f}%) have data")
    
    print()
    print("💡 Common Issues:")
    print(f"{'='*60}")
    print("1. Sparse feedback: Only 1-2 fields populated per employee")
    print("2. Short responses: Feedback is too brief (<50 words total)")
    print("3. Generic phrases: Not enough performance-specific content")
    print("4. Off-topic content: Detected as news/random by semantic AI")
    print()
    print("🔧 Recommendations:")
    print(f"{'='*60}")
    print("1. Encourage managers to provide more detailed feedback")
    print("2. Lower semantic confidence threshold from 50% to 40%")
    print("3. Accept blurbs with 25-30 words if they have performance keywords")
    print("4. Add more flexibility for short but valid feedback")
    print()


if __name__ == '__main__':
    analyze_failures()

