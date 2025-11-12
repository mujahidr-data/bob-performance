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
        self.spreadsheet_id = self.config.get('spreadsheet_id') or self.config.get('SPREADSHEET_ID')
        
        if not self.spreadsheet_id:
            print("❌ spreadsheet_id not found in config.json")
            print("   Please add 'spreadsheet_id' or 'SPREADSHEET_ID' to config/config.json")
            sys.exit(1)
        
        self.service = self._authenticate()
        self.summarizer = None
        self.semantic_validator = None
        
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
    
    def _init_semantic_validator(self):
        """Initialize semantic validation model (lazy loading)"""
        if self.semantic_validator is None:
            print("🔍 Loading semantic validation model...")
            print("   (This will verify blurbs are coherent performance reviews)")
            
            try:
                # Use zero-shot classification for semantic validation
                self.semantic_validator = pipeline(
                    "zero-shot-classification",
                    model="facebook/bart-large-mnli",
                    device=0 if torch.cuda.is_available() else -1
                )
                print("   ✓ Loaded BART-MNLI validator")
            except Exception as e:
                print(f"   ⚠️  Failed to load BART-MNLI, trying DistilBART: {e}")
                try:
                    self.semantic_validator = pipeline(
                        "zero-shot-classification",
                        model="typeform/distilbert-base-uncased-mnli",
                        device=-1
                    )
                    print("   ✓ Loaded DistilBERT-MNLI validator")
                except Exception as e2:
                    print(f"   ⚠️  Semantic validation unavailable: {e2}")
                    print("   ℹ️  Continuing with rule-based QA only")
                    self.semantic_validator = None
    
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
            try:
                if col_key in col_indices and col_indices[col_key] < len(row):
                    text = str(row[col_indices[col_key]]).strip()
                    if text and text not in ['', 'N/A', 'n/a', '-', 'None', 'None.']:
                        feedback_parts.append(text)
            except (IndexError, KeyError):
                continue
        
        # Combine all feedback
        combined = ' '.join(feedback_parts)
        return self._clean_text(combined)
    
    def _generate_blurb(self, feedback_text: str, emp_name: str = "") -> str:
        """Generate a concise manager blurb using AI with QA validation"""
        
        if not feedback_text or len(feedback_text) < 20:
            return "No meaningful feedback available"
        
        # Initialize summarizer if needed
        self._init_summarizer()
        
        # Try AI summarization first
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
            
            # QA validation pass (rule-based)
            is_valid_rules, reason = self._validate_blurb_quality(blurb, emp_name)
            
            if not is_valid_rules:
                # Failed rule-based QA - try fallback
                print(f"       ⚠️  Rule-based QA failed ({reason}), using fallback")
                fallback = self._simple_extract(feedback_text)
                
                # Validate fallback with rules
                is_valid_fallback, _ = self._validate_blurb_quality(fallback, emp_name)
                if is_valid_fallback:
                    blurb = fallback
                else:
                    # Both failed - return generic message
                    return "Performance feedback available; requires manual review for summary"
            
            # Semantic validation pass (AI-based)
            is_valid_semantic, semantic_reason, confidence = self._validate_semantic_coherence(blurb, emp_name)
            
            if not is_valid_semantic:
                print(f"       ⚠️  Semantic QA failed ({semantic_reason}), using fallback")
                fallback = self._simple_extract(feedback_text)
                
                # Validate fallback semantically
                is_valid_fallback_semantic, _, _ = self._validate_semantic_coherence(fallback, emp_name)
                is_valid_fallback_rules, _ = self._validate_blurb_quality(fallback, emp_name)
                
                if is_valid_fallback_semantic and is_valid_fallback_rules:
                    return fallback
                else:
                    return "Performance feedback available; requires manual review for summary"
            
            # Both validations passed
            return blurb
            
        except Exception as e:
            print(f"       ⚠️  Summarization failed for {emp_name}: {e}")
            # Fallback: simple extraction
            fallback = self._simple_extract(feedback_text)
            
            # Validate fallback (both rules and semantic)
            is_valid_rules, _ = self._validate_blurb_quality(fallback, emp_name)
            is_valid_semantic, _, _ = self._validate_semantic_coherence(fallback, emp_name)
            
            if is_valid_rules and is_valid_semantic:
                return fallback
            else:
                return "Performance feedback available; requires manual review for summary"
    
    def _format_performance_tone(self, text: str) -> str:
        """Format text to performance-review tone"""
        
        if not text or len(text.strip()) == 0:
            return ""
        
        text = text.strip()
        
        # Capitalize first letter
        if text:
            text = text[0].upper() + text[1:]
        
        # Ensure ends with period
        if not text.endswith(('.', '!', '?')):
            text += '.'
        
        # Fix common grammar issues
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'\s([.,;!?])', r'\1', text)
        
        return text
    
    def _validate_blurb_quality(self, blurb: str, emp_name: str = "") -> tuple[bool, str]:
        """
        QA validation pass for generated blurbs
        Returns: (is_valid, reason)
        """
        
        if not blurb or len(blurb.strip()) < 20:
            return False, "Too short"
        
        blurb_clean = blurb.strip()
        
        # Check 1: Proper capitalization at start
        if not blurb_clean[0].isupper():
            return False, "Missing capitalization"
        
        # Check 2: Proper ending punctuation
        if not blurb_clean.endswith(('.', '!', '?')):
            return False, "Missing ending punctuation"
        
        # Check 3: Check for obviously irrelevant content
        irrelevant_patterns = [
            r'cnn\.com',
            r'ireporter',
            r'travel snapshots',
            r'gallery\.',
            r'visi$',  # truncated
            r'\.com has been hacked',
            r'http',
            r'www\.',
            r'click here',
            r'subscribe',
            r'newsletter'
        ]
        
        blurb_lower = blurb_clean.lower()
        for pattern in irrelevant_patterns:
            if re.search(pattern, blurb_lower):
                return False, f"Irrelevant content detected: {pattern}"
        
        # Check 4: Must contain at least one performance-related keyword
        performance_keywords = [
            'delivered', 'developed', 'demonstrated', 'led', 'managed', 'improved',
            'needs', 'should', 'opportunity', 'focus', 'strengthen', 'expand',
            'performance', 'growth', 'impact', 'ownership', 'collaboration',
            'technical', 'leadership', 'team', 'project', 'product', 'customer',
            'ready', 'readiness', 'potential', 'promotion', 'ai', 'automation',
            'thinking', 'execution', 'delivery', 'quality', 'skill', 'ability'
        ]
        
        has_keyword = any(keyword in blurb_lower for keyword in performance_keywords)
        if not has_keyword:
            return False, "No performance-related keywords found"
        
        # Check 5: Word count should be reasonable (30-80 words)
        word_count = len(blurb_clean.split())
        if word_count < 30:
            return False, f"Too short ({word_count} words)"
        if word_count > 85:
            return False, f"Too long ({word_count} words)"
        
        # Check 6: Must have at least 2 complete sentences
        sentences = re.split(r'[.!?]+', blurb_clean)
        complete_sentences = [s.strip() for s in sentences if len(s.strip()) > 10]
        if len(complete_sentences) < 2:
            return False, "Incomplete or single sentence"
        
        # Check 7: No obvious mid-sentence truncation (ends with incomplete word)
        if blurb_clean.endswith('...'):
            return False, "Truncated with ellipsis"
        
        # Check 8: Should not be all lowercase or all uppercase
        if blurb_clean.islower() or blurb_clean.isupper():
            return False, "Improper casing"
        
        # Check 9: Check for gibberish (too many consecutive consonants or vowels)
        words = blurb_clean.split()
        for word in words[:5]:  # Check first 5 words
            if len(word) > 15 or re.search(r'[bcdfghjklmnpqrstvwxyz]{7,}', word.lower()):
                return False, f"Potential gibberish detected: {word}"
        
        return True, "Valid"
    
    def _validate_semantic_coherence(self, blurb: str, emp_name: str = "") -> tuple[bool, str, float]:
        """
        Semantic validation using zero-shot classification
        Returns: (is_valid, reason, confidence_score)
        """
        
        # Initialize semantic validator if needed
        self._init_semantic_validator()
        
        # If validator not available, skip semantic check
        if self.semantic_validator is None:
            return True, "Semantic validation skipped", 1.0
        
        try:
            # Define candidate labels for zero-shot classification
            candidate_labels = [
                "employee performance review",
                "professional development feedback",
                "random unrelated text",
                "gibberish or nonsense",
                "news article or website content"
            ]
            
            # Run zero-shot classification
            result = self.semantic_validator(
                blurb,
                candidate_labels,
                multi_label=False
            )
            
            # Get top prediction
            top_label = result['labels'][0]
            top_score = result['scores'][0]
            
            # Check if top prediction is performance-related
            performance_labels = [
                "employee performance review",
                "professional development feedback"
            ]
            
            if top_label in performance_labels:
                if top_score >= 0.5:  # High confidence it's a performance review
                    return True, f"Semantic: {top_label}", top_score
                else:
                    return False, f"Low confidence ({top_score:.2f}) for performance review", top_score
            else:
                # Top prediction is NOT performance-related
                return False, f"Semantic: Detected as '{top_label}' ({top_score:.2f})", top_score
        
        except Exception as e:
            # If semantic validation fails, don't block the blurb
            print(f"       ⚠️  Semantic validation error: {e}")
            return True, "Semantic validation error (skipped)", 0.0
    
    def _simple_extract(self, text: str) -> str:
        """Simple fallback extraction if AI fails - with quality checks"""
        sentences = re.split(r'[.!?]+', text)
        
        # Filter out very short or nonsensical sentences
        filtered_sentences = []
        for sentence in sentences:
            sentence = sentence.strip()
            if len(sentence) < 15:  # Too short
                continue
            
            # Check for obvious junk
            if re.search(r'(cnn\.com|ireporter|http|www\.|\.com has been)', sentence.lower()):
                continue
            
            filtered_sentences.append(sentence)
        
        # Take first 2-3 sentences that make sense
        selected = []
        word_count = 0
        
        for sentence in filtered_sentences:
            if not sentence:
                continue
            
            words_in_sentence = len(sentence.split())
            
            # Ensure we get at least 30 words total
            if word_count + words_in_sentence <= 70:
                selected.append(sentence)
                word_count += words_in_sentence
                
                # Stop if we have enough
                if word_count >= 40 and len(selected) >= 2:
                    break
            else:
                break
        
        if not selected:
            return "Performance feedback available; requires manual review for summary"
        
        result = '. '.join(selected)
        if result and not result.endswith('.'):
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
                
                # Get employee ID and name (with bounds checking)
                emp_id = ''
                emp_name = ''
                
                try:
                    if 'employee_id' in col_indices and col_indices['employee_id'] < len(row):
                        emp_id = str(row[col_indices['employee_id']]).strip()
                    if 'employee_name' in col_indices and col_indices['employee_name'] < len(row):
                        emp_name = str(row[col_indices['employee_name']]).strip()
                except IndexError:
                    print(f"   ⚠️  Skipping row {i+1}: Index out of range")
                    continue
                
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

