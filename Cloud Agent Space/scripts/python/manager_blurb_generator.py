#!/usr/bin/env python3
"""
Manager Blurb Generator
Generates concise, AI-powered performance review summaries from manager feedback
"""

import os
import sys
import json
import re
from typing import Dict, List, Optional

# Check for required libraries
try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except ImportError:
    print("❌ Missing required Google API libraries")
    print("   Install with: pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client")
    sys.exit(1)

try:
    from transformers import pipeline
    import torch
except ImportError:
    print("❌ Missing transformers library")
    print("   Install with: pip install transformers torch")
    sys.exit(1)


# Configuration
SPREADSHEET_ID = None  # Will be loaded from config
SOURCE_SHEET = "Bob Perf Report"
TARGET_SHEET = "Manager Blurbs"

# Leadership idioms to remove
LEADERSHIP_IDIOMS = [
    "get shit done", "gsd", "win as a team", "deep dive", "think big",
    "bias for action", "dive deep", "deliver results", "customer obsession",
    "earn trust", "have backbone", "insist on highest standards",
    "learn and be curious", "hire and develop", "ownership", "invent and simplify",
    "are right a lot", "frugality", "think and act like an owner"
]

# Filler words to minimize
FILLER_WORDS = [
    "very", "really", "quite", "just", "basically", "literally",
    "actually", "honestly", "obviously", "clearly", "definitely"
]


class ManagerBlurbGenerator:
    """Generate AI-powered manager blurbs from performance review feedback"""
    
    def __init__(self, config_path: str):
        """Initialize with config file"""
        self.config = self._load_config(config_path)
        self.spreadsheet_id = self.config.get('spreadsheet_id')
        self.service = self._authenticate()
        self.summarizer = None
        
    def _load_config(self, config_path: str) -> Dict:
        """Load configuration from JSON file"""
        try:
            with open(config_path, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"❌ Config file not found: {config_path}")
            sys.exit(1)
    
    def _authenticate(self):
        """Authenticate with Google Sheets API"""
        print("🔐 Authenticating with Google Sheets...")
        
        # Look for service account file
        service_account_paths = [
            os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'config', 'service_account.json'),
            os.path.join(os.path.dirname(__file__), 'service_account.json'),
            'config/service_account.json',
            'service_account.json'
        ]
        
        service_account_file = None
        for path in service_account_paths:
            if os.path.exists(path):
                service_account_file = path
                print(f"   ✓ Found service account: {path}")
                break
        
        if not service_account_file:
            print("❌ service_account.json not found")
            sys.exit(1)
        
        creds = service_account.Credentials.from_service_account_file(
            service_account_file,
            scopes=['https://www.googleapis.com/auth/spreadsheets']
        )
        
        return build('sheets', 'v4', credentials=creds)
    
    def _init_summarizer(self):
        """Initialize the summarization model (lazy loading)"""
        if self.summarizer is None:
            print("🤖 Loading BART summarization model...")
            print("   (This may take a minute on first run)")
            
            # Use facebook/bart-large-cnn for summarization
            # Falls back to smaller model if memory constrained
            try:
                self.summarizer = pipeline(
                    "summarization",
                    model="facebook/bart-large-cnn",
                    device=0 if torch.cuda.is_available() else -1  # GPU if available
                )
                print("   ✓ Loaded BART-large model")
            except Exception as e:
                print(f"   ⚠️  Failed to load BART-large, trying smaller model: {e}")
                try:
                    self.summarizer = pipeline(
                        "summarization",
                        model="sshleifer/distilbart-cnn-12-6",
                        device=-1  # CPU only
                    )
                    print("   ✓ Loaded DistilBART model")
                except Exception as e2:
                    print(f"❌ Failed to load any model: {e2}")
                    sys.exit(1)
    
    def _clean_text(self, text: str) -> str:
        """Clean and normalize text"""
        if not text or text.strip() == "":
            return ""
        
        # Remove leadership idioms
        text_lower = text.lower()
        for idiom in LEADERSHIP_IDIOMS:
            text_lower = re.sub(r'\b' + re.escape(idiom) + r'\b', '', text_lower, flags=re.IGNORECASE)
        
        # Remove filler words
        for filler in FILLER_WORDS:
            text_lower = re.sub(r'\b' + filler + r'\b', '', text_lower, flags=re.IGNORECASE)
        
        # Remove extra whitespace
        text_lower = re.sub(r'\s+', ' ', text_lower).strip()
        
        # Remove incomplete sentences at the end
        sentences = re.split(r'[.!?]+', text_lower)
        complete_sentences = [s.strip() for s in sentences if len(s.strip()) > 10]
        
        return '. '.join(complete_sentences)
    
    def _extract_feedback_text(self, row: List, col_indices: Dict) -> str:
        """Extract and combine all manager feedback text"""
        feedback_parts = []
        
        # Column mappings from Bob Perf Report
        feedback_columns = [
            'leadership_strength',
            'leadership_improvement',
            'ai_leverage',
            'ai_readiness',
            'support_needed',
            'performance_comment',
            'potential_comment',
            'promotion_comment'
        ]
        
        for col_key in feedback_columns:
            if col_key in col_indices and col_indices[col_key] < len(row):
                text = str(row[col_indices[col_key]]).strip()
                if text and text not in ['', 'N/A', 'n/a', '-', 'None']:
                    feedback_parts.append(text)
        
        # Combine all feedback
        combined = ' '.join(feedback_parts)
        return self._clean_text(combined)
    
    def _generate_blurb(self, feedback_text: str, emp_name: str = "") -> str:
        """Generate a concise manager blurb using AI"""
        
        if not feedback_text or len(feedback_text) < 20:
            return "No feedback available"
        
        # Initialize summarizer if needed
        self._init_summarizer()
        
        # Prepare prompt for performance-review style
        # BART works best with clear text, so we'll post-process the output
        
        try:
            # BART summarization
            max_length = 80  # Words for input to model (will be trimmed to 60)
            min_length = 40
            
            summary = self.summarizer(
                feedback_text,
                max_length=max_length,
                min_length=min_length,
                do_sample=False,
                truncation=True
            )
            
            blurb = summary[0]['summary_text']
            
            # Post-process to ensure performance-review tone
            blurb = self._format_performance_tone(blurb)
            
            # Limit to ~60 words
            words = blurb.split()
            if len(words) > 60:
                blurb = ' '.join(words[:60]) + '.'
            
            return blurb
            
        except Exception as e:
            print(f"   ⚠️  Summarization failed for {emp_name}: {e}")
            # Fallback: simple extraction
            return self._simple_extract(feedback_text)
    
    def _format_performance_tone(self, text: str) -> str:
        """Format text to performance-review tone"""
        
        # Capitalize first letter
        if text:
            text = text[0].upper() + text[1:]
        
        # Ensure ends with period
        if not text.endswith('.'):
            text += '.'
        
        # Fix common grammar issues
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'\s([.,;!?])', r'\1', text)
        
        return text
    
    def _simple_extract(self, text: str) -> str:
        """Simple fallback extraction if AI fails"""
        sentences = re.split(r'[.!?]+', text)
        
        # Take first 2-3 sentences
        selected = []
        word_count = 0
        
        for sentence in sentences:
            sentence = sentence.strip()
            if not sentence:
                continue
            
            words_in_sentence = len(sentence.split())
            if word_count + words_in_sentence <= 60:
                selected.append(sentence)
                word_count += words_in_sentence
            else:
                break
        
        result = '. '.join(selected)
        if result:
            result += '.'
        
        return self._format_performance_tone(result)
    
    def _get_column_indices(self, headers: List[str]) -> Dict:
        """Map column names to indices"""
        indices = {}
        
        # Map expected columns
        for i, header in enumerate(headers):
            header_lower = header.lower().strip()
            
            if 'employee' in header_lower and 'id' not in header_lower:
                indices['employee_name'] = i
            elif 'employee id' in header_lower or header_lower == 'employee':
                indices['employee_id'] = i
            elif 'leadership principle' in header_lower and 'exemplified' in header_lower:
                indices['leadership_strength'] = i
            elif 'leadership principle' in header_lower and 'improvement' in header_lower:
                indices['leadership_improvement'] = i
            elif 'leveraged ai' in header_lower or 'ai, automation' in header_lower:
                indices['ai_leverage'] = i
            elif 'ai has not yet' in header_lower or 'openness' in header_lower:
                indices['ai_readiness'] = i
            elif 'support, coaching' in header_lower:
                indices['support_needed'] = i
            elif 'comment' in header_lower and 'performed against' in header_lower:
                indices['performance_comment'] = i
            elif 'comment' in header_lower and 'potential' in header_lower:
                indices['potential_comment'] = i
            elif 'comment' in header_lower and 'promoted' in header_lower:
                indices['promotion_comment'] = i
        
        return indices
    
    def generate_all_blurbs(self):
        """Generate blurbs for all employees and write to hidden sheet"""
        print(f"📊 Reading data from '{SOURCE_SHEET}'...")
        
        try:
            # Read Bob Perf Report
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{SOURCE_SHEET}'!A1:Z10000"
            ).execute()
            
            data = result.get('values', [])
            
            if not data:
                print("❌ No data found in Bob Perf Report")
                return False
            
            headers = data[0]
            rows = data[1:]
            
            print(f"   ✓ Found {len(rows)} employees")
            
            # Get column indices
            col_indices = self._get_column_indices(headers)
            print(f"   ✓ Mapped columns: {list(col_indices.keys())}")
            
            # Generate blurbs
            blurbs_data = [['Emp ID', 'Manager Blurb']]  # Header
            
            print("🤖 Generating manager blurbs...")
            for i, row in enumerate(rows):
                if len(row) == 0:
                    continue
                
                # Get employee ID and name
                emp_id = str(row[col_indices['employee_id']]).strip() if 'employee_id' in col_indices else ''
                emp_name = str(row[col_indices['employee_name']]).strip() if 'employee_name' in col_indices else ''
                
                if not emp_id:
                    continue
                
                print(f"   [{i+1}/{len(rows)}] Processing {emp_name} ({emp_id})...")
                
                # Extract feedback
                feedback = self._extract_feedback_text(row, col_indices)
                
                # Generate blurb
                blurb = self._generate_blurb(feedback, emp_name)
                
                blurbs_data.append([emp_id, blurb])
                print(f"       ✓ {blurb[:80]}...")
            
            # Write to Google Sheets (create/update hidden sheet)
            print(f"\n📝 Writing blurbs to '{TARGET_SHEET}'...")
            self._write_to_sheet(blurbs_data)
            
            print(f"✅ Successfully generated {len(blurbs_data)-1} manager blurbs!")
            return True
            
        except HttpError as e:
            print(f"❌ Google Sheets API error: {e}")
            return False
        except Exception as e:
            print(f"❌ Error: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _write_to_sheet(self, data: List[List]):
        """Write blurbs to hidden sheet"""
        
        try:
            # Check if sheet exists
            spreadsheet = self.service.spreadsheets().get(
                spreadsheetId=self.spreadsheet_id
            ).execute()
            
            sheets = spreadsheet.get('sheets', [])
            sheet_names = [s['properties']['title'] for s in sheets]
            sheet_id = None
            
            if TARGET_SHEET not in sheet_names:
                print(f"   Creating sheet '{TARGET_SHEET}'...")
                request = {
                    'requests': [{
                        'addSheet': {
                            'properties': {
                                'title': TARGET_SHEET,
                                'hidden': True  # Hide the sheet
                            }
                        }
                    }]
                }
                response = self.service.spreadsheets().batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body=request
                ).execute()
                sheet_id = response['replies'][0]['addSheet']['properties']['sheetId']
                print(f"   ✓ Created hidden sheet")
            else:
                # Get sheet ID
                for sheet in sheets:
                    if sheet['properties']['title'] == TARGET_SHEET:
                        sheet_id = sheet['properties']['sheetId']
                        break
                print(f"   ✓ Found existing sheet")
            
            # Write data
            range_name = f"'{TARGET_SHEET}'!A1"
            body = {'values': data}
            
            self.service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id,
                range=range_name,
                valueInputOption='RAW',
                body=body
            ).execute()
            
            # Format header row
            if sheet_id is not None:
                requests = [{
                    'repeatCell': {
                        'range': {
                            'sheetId': sheet_id,
                            'startRowIndex': 0,
                            'endRowIndex': 1
                        },
                        'cell': {
                            'userEnteredFormat': {
                                'backgroundColor': {'red': 0.6, 'green': 0.15, 'blue': 1.0},
                                'textFormat': {
                                    'foregroundColor': {'red': 1.0, 'green': 1.0, 'blue': 1.0},
                                    'bold': True,
                                    'fontFamily': 'Roboto',
                                    'fontSize': 10
                                }
                            }
                        },
                        'fields': 'userEnteredFormat(backgroundColor,textFormat)'
                    }
                }]
                
                self.service.spreadsheets().batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body={'requests': requests}
                ).execute()
            
            print(f"   ✓ Wrote {len(data)} rows to sheet")
            
        except Exception as e:
            print(f"❌ Failed to write to sheet: {e}")
            raise


def main():
    """Main entry point"""
    print("🚀 Manager Blurb Generator\n")
    
    # Get config path
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(os.path.dirname(os.path.dirname(script_dir)), 'config', 'config.json')
    
    if not os.path.exists(config_path):
        print(f"❌ Config file not found: {config_path}")
        sys.exit(1)
    
    # Initialize and run
    generator = ManagerBlurbGenerator(config_path)
    success = generator.generate_all_blurbs()
    
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()

