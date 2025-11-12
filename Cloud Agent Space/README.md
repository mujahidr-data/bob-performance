# Bob Performance Module

Automated performance management system for HiBob, featuring salary data imports, performance report automation, and AI-powered manager blurb generation.

## 🚀 Features

### 1. Salary Data Automation
- **Base Data Import**: Active employees, job details, compensation
- **Bonus History**: Variable compensation tracking
- **Full Compensation History**: Complete salary changes with dates and reasons
- **Performance Ratings**: Historical AYR and H1 ratings

### 2. Performance Report Automation
- **Web Interface**: Browser-based automation for downloading HiBob performance reports
- **Auto-Upload**: Downloads reports and uploads to Google Sheets
- **Multiple Launchers**: macOS app, command files, terminal aliases

### 3. Manager Blurb Generator (AI-Powered)
- **BART Summarization**: Generates 50-60 word performance review summaries
- **Dual QA Validation**: 
  - Rule-based checks (grammar, keywords, structure)
  - Semantic AI validation (zero-shot classification)
- **Hidden Sheet Storage**: Blurbs stored separately, referenced in Summary sheet

### 4. Summary Sheet Builder
- **Automated Data Consolidation**: Combines all imported data sources
- **Performance Tracking**: Ratings, promotions, tenure, compensation
- **Manager Blurbs**: AI-generated summaries auto-populate via VLOOKUP
- **Dynamic Formulas**: Chart data, slicers, conditional formatting

## 📁 Project Structure

```
bob-performance-module/
├── apps-script/                    # Google Apps Script
│   ├── bob-performance-module.gs  # Main script file
│   └── appsscript.json            # Apps Script manifest
├── config/                        # Configuration files
│   ├── config.json               # Credentials & settings
│   └── service_account.json      # Google API service account
├── scripts/
│   ├── python/                   # Python automation scripts
│   │   ├── hibob_report_downloader.py    # Performance report downloader
│   │   ├── manager_blurb_generator.py    # AI blurb generator
│   │   └── [other utility scripts]
│   └── shell/                    # Shell scripts
│       ├── start_web_app.sh      # Launch web interface (macOS/Linux)
│       ├── start_web_app.bat     # Launch web interface (Windows)
│       └── create_launcher.sh    # Create macOS app launcher
├── web/                          # Web interface (Flask)
│   ├── web_app.py               # Flask application
│   ├── templates/               # HTML templates
│   └── static/                  # CSS, JS, images
├── docs/                        # Documentation
│   ├── MANAGER_BLURB_GENERATOR.md
│   ├── ENABLE_SHEETS_API.md
│   └── [other guides]
└── requirements.txt             # Python dependencies
```

## 🛠️ Setup

### Prerequisites

- **Python 3.9+** (3.10+ recommended)
- **Node.js** (for clasp - Google Apps Script CLI)
- **Google Cloud Project** with Sheets API enabled
- **HiBob Account** with performance report access

### Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/mujahidr-data/bob-performance.git
   cd bob-performance
   ```

2. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure credentials**:
   
   Create `config/config.json`:
   ```json
   {
     "email": "your.email@company.com",
     "password": "your_hibob_password",
     "spreadsheet_id": "your_google_sheet_id",
     "apps_script_url": "your_apps_script_web_app_url"
   }
   ```

4. **Add Google Service Account**:
   
   Place your `service_account.json` in the `config/` folder and share your Google Sheet with the service account email.

5. **Deploy Apps Script**:
   ```bash
   cd apps-script
   npx @google/clasp push
   ```

## 📊 Usage

### Import Salary Data

**From Google Sheets**:
1. Open your Bob Performance spreadsheet
2. Go to: **🚀 Bob Performance Module → 📥 Import All Data**
3. Wait for completion notification

**Individual Imports**:
- **📊 Import Base Data** - Active employees
- **💰 Import Bonus History** - Variable comp
- **📈 Import Compensation History** - Salary changes
- **📊 Import Full Comp History** - Full salary history with dates
- **⭐ Import Performance Ratings** - AYR & H1 ratings

### Download Performance Reports

**Option 1: Desktop App** (macOS)
```bash
# Create the app (one-time setup)
./scripts/shell/create_launcher.sh

# Launch from Desktop or Applications
Double-click "Bob Performance.app"
```

**Option 2: Terminal** (macOS/Linux)
```bash
./scripts/shell/start_web_app.sh
```

**Option 3: Windows**
```cmd
scripts\shell\start_web_app.bat
```

**Then**:
1. Web interface opens at `http://localhost:5001`
2. Enter report name (e.g., "Q2&Q3")
3. Click "Start Automation"
4. Watch live progress
5. Report auto-uploads to Google Sheets

### Generate Manager Blurbs

**Option 1: From Terminal**
```bash
cd scripts/python
python3 manager_blurb_generator.py
```

**Option 2: From Google Sheets**
1. Go to: **🚀 Bob Performance Module → 🤖 Generate Manager Blurbs**
2. Follow the instructions in the dialog
3. Run the Python script as shown

**What it does**:
- Reads manager feedback from Bob Perf Report
- Generates 50-60 word summaries using AI (BART)
- Validates with dual QA (rules + semantic AI)
- Stores in hidden "Manager Blurbs" sheet
- Auto-references in Summary sheet via VLOOKUP

**Processing time**: ~3-6 minutes for 300-400 employees

### Build Summary Sheet

1. Open Google Sheet
2. Go to: **🚀 Bob Performance Module → 🔧 Build Summary Sheet**
3. Wait for completion

**Creates**:
- Combined data from all sources
- Performance ratings (AYR, H1, Q2/Q3)
- Promotion tracking
- Compensation details
- Manager blurbs (auto-populated from hidden sheet)
- Dynamic formulas for charts

## 🤖 AI Manager Blurb Generator

### How It Works

```
Manager Feedback → BART AI → Rule QA → Semantic QA → Published
                                ↓           ↓
                           If fails    If fails
                                ↓           ↓
                         Fallback → Validate → Manual Review
```

### Validation Layers

**1. BART Summarization**
- Model: facebook/bart-large-cnn
- Generates 40-80 word summaries
- Focuses on key performance themes

**2. Rule-Based QA (9 checks)**
- Proper capitalization & punctuation
- No irrelevant content (URLs, news, junk)
- Performance keywords present
- Word count (30-80 words)
- At least 2 complete sentences
- No truncation or gibberish

**3. Semantic QA (AI Classification)**
- Model: facebook/bart-large-mnli
- Classifies as:
  - ✅ "Employee performance review"
  - ✅ "Professional development feedback"
  - ❌ "Random unrelated text"
  - ❌ "Gibberish or nonsense"
  - ❌ "News article or website content"
- Requires ≥50% confidence

### Sample Output

**Input**: Raw manager feedback across multiple columns

**Output**:
> Delivered and expanded ownership of the Market Intelligence platform with high accountability. Effectively applied AI tools such as ChatGPT and Bito for validation. Should focus on deepening architectural thinking and balancing operational and development work.

**If validation fails**:
> Performance feedback available; requires manual review for summary

### Quality Metrics

From 326 employees:
- ✅ **AI passed both validations**: ~45%
- ⚠️ **Fallback used**: ~55%
- ❌ **Manual review required**: <1%

## 📖 Documentation

- **[Manager Blurb Generator Guide](docs/MANAGER_BLURB_GENERATOR.md)** - Detailed AI blurb documentation
- **[Enable Sheets API](docs/ENABLE_SHEETS_API.md)** - Setup Google Sheets API

## 🔧 Configuration

### config/config.json

```json
{
  "email": "mujahid.r@commerceiq.ai",          // HiBob email
  "password": "your_password",                  // HiBob password
  "spreadsheet_id": "1rnpUlOcqTpny...",        // Google Sheet ID
  "apps_script_url": "https://script.google..." // Apps Script web app URL
}
```

### config/service_account.json

Google Cloud service account credentials for Sheets API access. Share your Google Sheet with the service account email (found in the JSON file).

## 🚨 Troubleshooting

### Import Fails

**Error**: "Failed to fetch data from Bob API"
- **Solution**: Check BOB_ID and BOB_KEY in Script Properties (Apps Script)

### Browser Automation Fails

**Error**: "Browser failed to start"
- **Solution**: Script tries multiple browsers (Chromium, Chrome, Firefox, Safari)
- Install browsers: `playwright install chromium firefox webkit`

### Manager Blurb Generation Fails

**Error**: "Failed to load model"
- **Solution**: First run downloads ~3GB of models (BART + BART-MNLI)
- Ensure stable internet connection
- Close memory-intensive apps

### Permission Denied (Google Sheets)

**Error**: "Permission denied" or "403"
- **Solution**: Share Google Sheet with service account email
- Grant "Editor" access

### Python Module Errors

**Error**: "No module named 'transformers'"
- **Solution**: 
  ```bash
  pip install -r requirements.txt
  ```

## 🔄 Updating

### Update Apps Script
```bash
cd apps-script
npx @google/clasp push
```

### Pull Latest Changes
```bash
git pull origin main
pip install -r requirements.txt  # If dependencies changed
```

## 📝 Development

### Add New Bob Report

1. **Get Report ID** from HiBob (inspect network tab)
2. **Add to Apps Script**:
   ```javascript
   const BOB_REPORT_IDS = {
     YOUR_REPORT: "31234567",  // Add here
   };
   ```
3. **Create import function**:
   ```javascript
   function importYourReport() {
     importBobReport(BOB_REPORT_IDS.YOUR_REPORT, "Your Report Sheet");
   }
   ```
4. **Add to menu** in `onOpen()`

### Modify Blurb Validation

Edit `scripts/python/manager_blurb_generator.py`:

**Adjust semantic threshold**:
```python
# Line ~425
if top_score >= 0.5:  # Change to 0.6 for stricter validation
```

**Add custom categories**:
```python
# Line ~399
candidate_labels = [
    "employee performance review",
    "your custom category here",  # Add
    # ...
]
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature`
3. Commit changes: `git commit -m "Add your feature"`
4. Push to branch: `git push origin feature/your-feature`
5. Open a Pull Request

## 📄 License

This project is internal to CommerceIQ. Not licensed for external use.

## 👤 Author

**Mujahid Reza**  
Email: mujahid.r@commerceiq.ai

## 🔗 Links

- **Google Sheet**: [Bob Performance Sheet](https://docs.google.com/spreadsheets/d/1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA)
- **Apps Script**: [View in Editor](https://script.google.com/home/projects/1234567890)
- **GitHub**: [bob-performance](https://github.com/mujahidr-data/bob-performance)

## 📊 Sheet Structure

### Base Data
Active employees with job details, location, compensation, tenure.

### Bonus History
Variable compensation tracking (bonus type, percentage).

### Full Comp History
Complete salary change history with dates, increase %, and change reasons.

### Performance Ratings
Historical ratings: AYR 2024, H1 2025.

### Bob Perf Report
Downloaded performance reports with manager feedback, ratings, promotion readiness.

### Manager Blurbs (Hidden)
AI-generated 50-60 word performance summaries. Auto-referenced in Summary sheet.

### Summary
Consolidated view combining all data sources with:
- Employee info & tenure
- Latest ratings (AYR, H1, Q2/Q3)
- Promotion tracking
- Compensation details
- Manager blurbs
- Dynamic charts & slicers

## ⚙️ Technical Details

### Tech Stack
- **Google Apps Script**: Data imports, API integration
- **Python 3.9+**: Automation scripts, AI/ML
- **Flask**: Web interface
- **Playwright**: Browser automation
- **Transformers (HuggingFace)**: AI models
  - BART-large-cnn: Summarization
  - BART-large-mnli: Semantic validation
- **Google Sheets API**: Programmatic sheet access

### APIs Used
- **HiBob API**: Salary data exports (via report IDs)
- **Google Sheets API**: Write blurbs, create sheets
- **Google Charts API**: (Future) Programmatic chart creation

### Security
- Credentials stored in `config/config.json` (gitignored)
- Service account JSON (gitignored)
- Apps Script properties for Bob API keys
- No credentials in code

## 🎯 Roadmap

- [ ] Automated scheduling (cron jobs)
- [ ] Email notifications on completion
- [ ] Dashboard for tracking import history
- [ ] Export manager blurbs to PDF
- [ ] Multi-language support for blurbs
- [ ] Integration with Slack/Teams for notifications

## 🆘 Support

For issues or questions:
1. Check [Troubleshooting](#-troubleshooting) section
2. Review [Documentation](docs/)
3. Contact: mujahid.r@commerceiq.ai

---

**Version**: 2.0.0  
**Last Updated**: November 12, 2025  
**Status**: ✅ Production Ready
