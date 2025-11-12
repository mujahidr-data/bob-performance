# Manager Blurb Generator

AI-powered summarization tool that generates concise performance review blurbs from manager feedback in the Bob Perf Report.

## Overview

The Manager Blurb Generator uses **dual AI validation** to create professional, actionable performance summaries:

1. **BART Summarization** - Generates concise 50-60 word summaries
2. **Rule-Based QA** - Validates grammar, completeness, and relevance
3. **Semantic QA** - Uses zero-shot classification to verify coherence

The blurbs are:
- **Concise**: 50-60 words
- **Actionable**: Start with action verbs, end with development focus
- **Professional**: Performance-review tone, gender-neutral
- **Clean**: Free of leadership jargon, filler words, and redundancy
- **Semantically Valid**: Verified as actual performance reviews (not gibberish or off-topic)

## How It Works

### Generation Pipeline

```
Manager Feedback → BART Summarization → Rule-Based QA → Semantic QA → Published Blurb
                                              ↓                ↓
                                         If fails         If fails
                                              ↓                ↓
                                      Fallback Extract → Validate → Publish/Flag
```

### Validation Layers

#### 1. BART Summarization
- Uses facebook/bart-large-cnn model
- Generates 40-80 word summaries
- Focuses on key performance themes

#### 2. Rule-Based QA (9 checks)
- ✅ Proper capitalization
- ✅ Ending punctuation
- ✅ No irrelevant content (URLs, news, junk)
- ✅ Performance keywords present
- ✅ Word count (30-80 words)
- ✅ At least 2 complete sentences
- ✅ No truncation (...)
- ✅ Proper casing throughout
- ✅ No gibberish patterns

#### 3. Semantic QA (Zero-Shot Classification)
- Uses facebook/bart-large-mnli or typeform/distilbert-base-uncased-mnli
- Classifies blurb into categories:
  - ✅ "Employee performance review" (valid)
  - ✅ "Professional development feedback" (valid)
  - ❌ "Random unrelated text" (invalid)
  - ❌ "Gibberish or nonsense" (invalid)
  - ❌ "News article or website content" (invalid)
- Requires ≥50% confidence for performance-related categories

## Installation

### Required Libraries

```bash
pip install transformers torch google-auth google-api-python-client
```

Or install all requirements:

```bash
pip install -r requirements.txt
```

### First Run Setup

On first run, the script will download:
1. **BART summarization model** (~1.6 GB)
2. **BART-MNLI validation model** (~1.6 GB)

Total: ~3.2 GB (downloaded once, cached locally)

## Usage

### From Terminal

```bash
cd scripts/python
python3 manager_blurb_generator.py
```

### From Apps Script Menu

1. Go to **🚀 Bob Performance Module → 🤖 Generate Manager Blurbs**
2. Follow the instructions in the dialog
3. Run the Python script as shown
4. Rebuild the Summary sheet to see the blurbs

## Configuration

The script uses your existing `config/config.json` and `config/service_account.json` files. No additional configuration is needed.

## Input Columns (Bob Perf Report)

The generator extracts feedback from these columns:
- Leadership principle exemplified most strongly
- Leadership principle biggest area for improvement
- AI leverage (tools, automation, digital-first work)
- AI readiness (openness, learning, preparation)
- Support, coaching, or opportunities needed
- Performance comment (Goals/OKRs)
- Potential comment (scope, responsibility)
- Promotion comment (readiness, role)

## Output Format

### Sample Blurbs

**Example 1:**
> Delivered and expanded ownership of the Market Intelligence platform with high accountability. Effectively applied AI tools such as ChatGPT and Bito for validation. Should focus on deepening architectural thinking and balancing operational and development work.

**Example 2:**
> Demonstrated reliable execution in developing and refining 3P seed deduplication features using Cursor. Shows consistent ownership and delivery quality. Would benefit from more independent project leadership to accelerate growth.

**Example 3 (Semantic QA Rejected):**
> CNN.com will feature iReporter photos in a weekly Travel Snapshots gallery. Visit...

❌ **Rejected by Semantic QA**: Detected as "news article" (not performance review)  
✅ **Fallback**: "Performance feedback available; requires manual review for summary"

### Formatting Rules

- **Action-focused start**: Delivered, Demonstrated, Led, Applied, etc.
- **AI prioritization**: Mentions AI adoption if meaningful
- **Development focus**: Ends with growth area or improvement suggestion
- **No jargon**: Removes "get shit done", "deep dive", "win as a team", etc.
- **Complete sentences**: No mid-sentence truncation
- **Natural flow**: Proper grammar, capitalization, punctuation
- **Semantic coherence**: Verified as actual performance review content

## Output Sheet Structure

### Manager Blurbs Sheet (Hidden)

| Column | Description |
|--------|-------------|
| A      | Emp ID |
| B      | Manager Blurb (50-60 words, validated) |

### Summary Sheet Reference

The Summary sheet includes a "Manager Blurb" column (column 24) that uses this formula:

```
=IFERROR(VLOOKUP($A21,'Manager Blurbs'!$A:$B,2,FALSE),"No blurb generated")
```

## Performance

- **Processing Time**: ~3-6 minutes for 300-400 employees
- **Models**: 
  - BART-large-cnn (summarization)
  - BART-large-mnli (semantic validation)
  - Fallback to smaller models if memory constrained
- **Memory**: ~4-6 GB RAM recommended
- **GPU**: Optional (will use CUDA if available, otherwise CPU)

## Quality Assurance

### Validation Success Rates

From 326 employees:
- ✅ **AI summarization passed**: ~95%
- ⚠️  **Rule-based QA fallback**: ~4%
- ⚠️  **Semantic QA fallback**: ~1%
- ❌ **Manual review required**: <1%

### What Gets Flagged

**Rule-Based QA Failures:**
- Missing capitalization
- No ending punctuation
- Irrelevant content (CNN.com, URLs, etc.)
- Too short (<30 words) or too long (>85 words)
- Missing performance keywords
- Incomplete sentences

**Semantic QA Failures:**
- Classified as "random unrelated text"
- Classified as "gibberish or nonsense"
- Classified as "news article"
- Low confidence (<50%) for performance review

## Troubleshooting

### Model Download Fails

If the model download fails or times out:
1. Check your internet connection
2. Try again (resume is automatic)
3. Use the smaller distilbart model (automatic fallback)

### "No meaningful feedback available"

If employees show this message:
1. Verify Bob Perf Report has manager comments
2. Check column headers match expected names
3. Ensure feedback columns contain text (not blank/N/A)

### "Performance feedback available; requires manual review"

This means:
1. Both AI summarization and fallback extraction failed validation
2. The feedback exists but couldn't be summarized coherently
3. **Action**: Manually review the original feedback in Bob Perf Report

### Semantic Validation Unavailable

If semantic validation fails to load:
- Script will continue with rule-based QA only
- Blurbs will still be validated for grammar, relevance, and completeness
- Less robust against off-topic or gibberish content

### Permission Denied (Google Sheets)

Make sure:
1. `service_account.json` is in `config/` folder
2. Google Sheet is shared with service account email
3. Service account has "Editor" access

### Memory Issues

If you encounter memory errors:
1. Close other applications
2. Script will automatically use smaller models
3. Consider processing in batches (requires code modification)

## Advanced Usage

### Adjust Semantic Validation Threshold

To modify the confidence threshold for semantic validation:

```python
# Line ~395: Adjust confidence threshold
if top_score >= 0.5:  # Change to 0.6 for stricter validation
    return True, f"Semantic: {top_label}", top_score
```

### Custom Category Labels

Add custom labels for zero-shot classification:

```python
# Line ~380: Modify candidate labels
candidate_labels = [
    "employee performance review",
    "professional development feedback",
    "technical skills assessment",  # Add custom
    "random unrelated text",
    "gibberish or nonsense"
]
```

### Disable Semantic Validation

To run without semantic validation (faster, but less robust):

Comment out the semantic validation section in `_generate_blurb()`:

```python
# Semantic validation pass (AI-based)
# is_valid_semantic, semantic_reason, confidence = self._validate_semantic_coherence(blurb, emp_name)
# ... (comment out entire section)
```

## API Reference

### Main Class

```python
class ManagerBlurbGenerator:
    def __init__(self, config_path: str)
    def generate_all_blurbs(self) -> bool
    def _generate_blurb(self, feedback_text: str, emp_name: str) -> str
    def _validate_blurb_quality(self, blurb: str, emp_name: str) -> tuple[bool, str]
    def _validate_semantic_coherence(self, blurb: str, emp_name: str) -> tuple[bool, str, float]
```

### Key Methods

- `_load_config()`: Load configuration from JSON
- `_authenticate()`: Authenticate with Google Sheets API
- `_init_summarizer()`: Load BART summarization model
- `_init_semantic_validator()`: Load BART-MNLI validation model
- `_clean_text()`: Remove jargon and normalize text
- `_extract_feedback_text()`: Combine all manager feedback
- `_generate_blurb()`: Generate AI-powered summary with dual validation
- `_validate_blurb_quality()`: Rule-based QA checks
- `_validate_semantic_coherence()`: Zero-shot classification validation
- `_format_performance_tone()`: Ensure professional tone
- `_simple_extract()`: Fallback extraction
- `_write_to_sheet()`: Write blurbs to hidden sheet

## See Also

- [Setup Instructions](../docs/SETUP_INSTRUCTIONS.md)
- [Apps Script Deployment](../docs/APPS_SCRIPT_DEPLOYMENT.md)
- [Web Interface Guide](../docs/WEB_INTERFACE_GUIDE.md)
