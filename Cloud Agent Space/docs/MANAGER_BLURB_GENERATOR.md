# Manager Blurb Generator

AI-powered summarization tool that generates concise performance review blurbs from manager feedback in the Bob Perf Report.

## Overview

The Manager Blurb Generator uses BART (Bidirectional and Auto-Regressive Transformers) to create professional, actionable performance summaries that are:
- **Concise**: 50-60 words
- **Actionable**: Start with action verbs, end with development focus
- **Professional**: Performance-review tone, gender-neutral
- **Clean**: Free of leadership jargon, filler words, and redundancy

## How It Works

1. **Reads** manager feedback from Bob Perf Report (all comment columns)
2. **Cleans** text by removing leadership idioms, filler words, and incomplete sentences
3. **Summarizes** using BART AI model to generate concise, professional blurbs
4. **Stores** results in a hidden "Manager Blurbs" sheet
5. **References** blurbs in Summary sheet via VLOOKUP formulas

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

On first run, the script will download the BART model (~1-2 GB). This may take a few minutes depending on your internet connection.

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

### Formatting Rules

- **Action-focused start**: Delivered, Demonstrated, Led, Applied, etc.
- **AI prioritization**: Mentions AI adoption if meaningful
- **Development focus**: Ends with growth area or improvement suggestion
- **No jargon**: Removes "get shit done", "deep dive", "win as a team", etc.
- **Complete sentences**: No mid-sentence truncation
- **Natural flow**: Proper grammar, capitalization, punctuation

## Output Sheet Structure

### Manager Blurbs Sheet (Hidden)

| Column | Description |
|--------|-------------|
| A      | Emp ID |
| B      | Manager Blurb (50-60 words) |

### Summary Sheet Reference

The Summary sheet includes a "Manager Blurb" column (column 24) that uses this formula:

```
=IFERROR(VLOOKUP($A21,'Manager Blurbs'!$A:$B,2,FALSE),"No blurb generated")
```

## Performance

- **Processing Time**: ~2-5 minutes for 300-400 employees
- **Model**: facebook/bart-large-cnn (or distilbart-cnn-12-6 as fallback)
- **Memory**: ~2-4 GB RAM recommended
- **GPU**: Optional (will use CUDA if available, otherwise CPU)

## Troubleshooting

### Model Download Fails

If the model download fails or times out:
1. Check your internet connection
2. Try again (resume is automatic)
3. Use the smaller distilbart model (automatic fallback)

### "No feedback available"

If employees show this message:
1. Verify Bob Perf Report has manager comments
2. Check column headers match expected names
3. Ensure feedback columns contain text (not blank/N/A)

### Permission Denied (Google Sheets)

Make sure:
1. `service_account.json` is in `config/` folder
2. Google Sheet is shared with service account email
3. Service account has "Editor" access

### Memory Issues

If you encounter memory errors:
1. Close other applications
2. Script will automatically use smaller distilbart model
3. Consider processing in batches (requires code modification)

## Advanced Usage

### Custom Prompts

To modify the summarization behavior, edit `manager_blurb_generator.py`:

```python
# Line ~180: Adjust max_length, min_length
summary = self.summarizer(
    feedback_text,
    max_length=80,  # Increase for longer blurbs
    min_length=40,   # Decrease for shorter blurbs
    do_sample=False,
    truncation=True
)
```

### Additional Filtering

Add custom idioms or filler words to remove:

```python
# Line ~30: Add to LEADERSHIP_IDIOMS list
LEADERSHIP_IDIOMS = [
    "get shit done",
    "your custom phrase here",
    # ...
]

# Line ~40: Add to FILLER_WORDS list
FILLER_WORDS = [
    "very",
    "your filler word here",
    # ...
]
```

## API Reference

### Main Class

```python
class ManagerBlurbGenerator:
    def __init__(self, config_path: str)
    def generate_all_blurbs(self) -> bool
    def _generate_blurb(self, feedback_text: str, emp_name: str) -> str
```

### Key Methods

- `_load_config()`: Load configuration from JSON
- `_authenticate()`: Authenticate with Google Sheets API
- `_init_summarizer()`: Load BART model (lazy loading)
- `_clean_text()`: Remove jargon and normalize text
- `_extract_feedback_text()`: Combine all manager feedback
- `_generate_blurb()`: Generate AI-powered summary
- `_format_performance_tone()`: Ensure professional tone
- `_write_to_sheet()`: Write blurbs to hidden sheet

## See Also

- [Setup Instructions](../docs/SETUP_INSTRUCTIONS.md)
- [Apps Script Deployment](../docs/APPS_SCRIPT_DEPLOYMENT.md)
- [Web Interface Guide](../docs/WEB_INTERFACE_GUIDE.md)

