# Bob Performance Module

Comprehensive Google Apps Script + Python solution for automating performance review data management, AI-powered manager feedback summarization, and AI readiness assessment.

---

## 🚀 Features

### 1. **Bob Salary Data Imports**
- Automated imports from HiBob API for:
  - Base Data (employees)
  - Bonus History
  - Compensation History
  - Full Comp History (with increase dates)
  - Performance Ratings (AYR 2024, H1 2025)
- Active/Permanent employee filtering
- Automated column mapping and data validation

### 2. **Performance Report Automation**
- Web interface for downloading "Bob Perf Report" (no HiBob API endpoint available)
- Playwright-based browser automation
- Automatic upload to Google Sheets (preserves formatting)
- Multiple browser support (Chromium, Chrome, Firefox, Safari/Webkit)

### 3. **Summary Sheet Builder**
- Consolidates data from multiple sheets
- Calculates tenure, ratings, compensation, FX conversion
- Integrates AI-powered manager blurbs
- **NEW**: AI Readiness Category auto-population
- Dynamic chart data and slicers
- Formatted for performance calibration

### 4. **AI-Powered Manager Blurb Generation**
- Uses BART (transformer-based summarization model)
- Extracts and summarizes manager feedback from multiple columns
- Two-layer QA validation:
  - **Rule-based**: Grammar, keywords, completeness
  - **Semantic**: AI coherence verification
- Generates 50-60 word performance-focused blurbs
- 65%+ success rate with automatic quality filtering

### 5. **AI Readiness Assessment** (NEW)
- Comprehensive function-to-AI-category mapping
- 12 function categories with specific AI transformation criteria
- Auto-populates AI Readiness Category based on employee department
- Leadership assessment framework (AI-Ready, AI-Capable, Not AI-Ready)
- Detailed implementation guide and documentation

---

## 📂 Project Structure

```
Cloud Agent Space/
├── apps-script/
│   ├── bob-performance-module.gs      # Main Apps Script (consolidated)
│   ├── .clasp.json                    # Clasp configuration
│   └── appsscript.json                # Apps Script manifest
├── scripts/
│   ├── python/
│   │   ├── hibob_report_downloader.py # Performance report automation
│   │   ├── manager_blurb_generator.py # AI summarization script
│   │   └── analyze_blurb_failures.py  # QA analysis tool
│   └── launchers/
│       ├── start_web_app.sh           # macOS/Linux launcher
│       ├── start_web_app.bat          # Windows launcher
│       └── create_launcher.sh         # macOS app creator
├── web/
│   ├── web_app.py                     # Flask web interface
│   ├── templates/
│   │   ├── index.html                 # Web UI
│   │   └── status.html                # Status page
│   └── static/
│       ├── bob-icon.png               # App icon
│       └── styles.css                 # Web styles
├── config/
│   ├── config.json                    # Configuration (spreadsheet ID, etc.)
│   └── service_account.json           # Google Sheets API credentials
├── docs/
│   ├── AI_READINESS_CATEGORIES.md     # AI transformation criteria by function
│   ├── AI_READINESS_IMPLEMENTATION.md # Implementation guide
│   ├── MANAGER_BLURB_GENERATOR.md     # Blurb generator documentation
│   └── [other docs]
├── requirements.txt                    # Python dependencies
├── package.json                        # NPM dependencies (clasp)
└── README.md                           # This file
```

---

## 🛠️ Setup

### Prerequisites

1. **Python 3.9+**
2. **Node.js** (for clasp)
3. **Google Account** with access to:
   - Google Sheets
   - Google Apps Script
4. **HiBob Account** with:
   - API credentials (BOB_ID, BOB_KEY)
   - Access to reports

### Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/mujahidr-data/bob-performance.git
   cd "Cloud Agent Space"
   ```

2. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Install Playwright browsers**:
   ```bash
   playwright install
   ```

4. **Configure Google Sheets API**:
   - Create a service account in Google Cloud Console
   - Download `service_account.json`
   - Place it in `config/` directory
   - Share your Google Sheet with the service account email

5. **Configure Apps Script**:
   - Create a new Google Apps Script project
   - Add BOB_ID and BOB_KEY to Script Properties
   - Deploy as a container-bound script to your Google Sheet

6. **Update config.json**:
   ```json
   {
     "spreadsheet_id": "YOUR_SPREADSHEET_ID",
     "SPREADSHEET_ID": "YOUR_SPREADSHEET_ID"
   }
   ```

7. **Sync with clasp** (optional):
   ```bash
   cd apps-script
   npx @google/clasp login
   npx @google/clasp push
   ```

---

## 📖 Usage

### Web Interface (Recommended)

1. **Launch the web app**:
   ```bash
   cd "Cloud Agent Space"
   ./scripts/launchers/start_web_app.sh
   ```
   Or double-click the `Bob Performance.app` (macOS only, created via `create_launcher.sh`)

2. **Open browser**: `http://localhost:5001`

3. **Download report**:
   - Enter report name (e.g., "Q2&Q3")
   - Click "Start Automation"
   - Browser opens, logs in, downloads, uploads to Sheets

### Google Sheets Menu

Open your Google Sheet and use the **🚀 Bob Performance Module** menu:

1. **Import Base Data** - Import employee data
2. **Import Bonus History** - Import bonus data
3. **Import Compensation History** - Import comp changes
4. **Import Full Comp History** - Import full comp with increase dates
5. **Import Performance Ratings** - Import AYR/H1 ratings
6. **Import All Data** - Run all imports sequentially
7. **Build Summary Sheet** - Generate consolidated summary
8. **Generate Manager Blurbs** - Instructions for AI summarization
9. **Create AI Readiness Mapping** - Create AI category mapping sheet
10. **Launch Web Interface** - Instructions for web app
11. **Instructions & Help** - Usage guide

### AI Manager Blurb Generation

1. Ensure `Bob Perf Report` sheet has manager feedback
2. Run from terminal:
   ```bash
   cd "Cloud Agent Space/scripts/python"
   python3 manager_blurb_generator.py
   ```
3. Blurbs are written to hidden "Manager Blurbs" sheet
4. Summary sheet references them via VLOOKUP
5. **First run**: ~5-10 min (downloads AI models)
6. **Subsequent runs**: ~2-3 min for 300+ employees

**QA Results**:
- ✅ **65.3% Valid Blurbs** - Ready to use
- ⚠️ **15.3% Manual Review** - Requires attention
- ℹ️ **19.3% No Feedback** - Missing source data

### AI Readiness Assessment

1. **Create AI Readiness Mapping**:
   - Run **🎯 Create AI Readiness Mapping** from menu
   - Reviews/customizes department-to-category mappings

2. **Build Summary Sheet**:
   - Run **🔧 Build Summary Sheet**
   - "AI Readiness Category" column auto-populates

3. **Add Assessment Column** (Manual):
   - Add "AI Readiness Assessment" column
   - Data validation: `AI-Ready, AI-Capable, Not AI-Ready`
   - Conditional formatting (Green/Amber/Red)

4. **Leadership Assessment**:
   - Review AI Readiness Category for each employee
   - Evaluate against 5 criteria (see `docs/AI_READINESS_CATEGORIES.md`)
   - Tag employees based on AI adoption readiness

**See full implementation guide**: `docs/AI_READINESS_IMPLEMENTATION.md`

---

## 📊 Google Sheets Structure

| Sheet Name | Purpose | Source |
|------------|---------|--------|
| Base Data | Active employee master data | HiBob API |
| Bonus History | Bonus/variable comp | HiBob API |
| Comp History | Salary change history | HiBob API |
| Full Comp History | Full comp with increase dates | HiBob API |
| Performance Ratings | AYR 2024, H1 2025 ratings | HiBob API |
| Bob Perf Report | Q2/Q3 performance report | Web automation (no API) |
| Summary | Consolidated performance view | Auto-generated |
| Manager Blurbs | AI-generated summaries (hidden) | Python script |
| AI Readiness Mapping | Function-to-AI-category map | Auto-generated |

---

## 🎯 AI Readiness Categories

### Function Categories

| Function | AI Readiness Category |
|----------|----------------------|
| Engineering | AI-Accelerated Developer |
| Product | AI-Forward Solutions Engineer |
| Customer Success | AI Outcome Manager |
| Marketing | AI Content Intelligence Strategist |
| Sales | AI-Augmented Revenue Accelerator |
| HR / People Ops | AI-Powered People Strategist |
| Finance | AI-Driven Financial Intelligence Analyst |
| G&A / Operations | AI Operations Optimizer |
| Support / Service Delivery | AI-Enhanced Support Specialist |
| Data / Analytics | AI-Native Insights Engineer |
| Design / UX | AI-Augmented Experience Designer |
| Information Security | AI-Powered Security Guardian |
| Legal / Compliance | AI-Assisted Legal Intelligence Analyst |

### Assessment Levels

- **AI-Ready**: 3+ criteria met, measurable impact, proactive adoption
- **AI-Capable**: 2 criteria met, uses AI but limited impact, needs coaching
- **Not AI-Ready**: 0-1 criteria, resistant to AI, manual-first approach

**Full criteria**: `docs/AI_READINESS_CATEGORIES.md`

---

## 🔧 Configuration

### config.json

```json
{
  "spreadsheet_id": "YOUR_SPREADSHEET_ID",
  "SPREADSHEET_ID": "YOUR_SPREADSHEET_ID"
}
```

### Apps Script Properties

- `BOB_ID`: HiBob API user ID
- `BOB_KEY`: HiBob API key

### Environment Variables (Optional)

- `PORT`: Web app port (default: 5001)
- `HIBOB_EMAIL`: HiBob login email
- `HIBOB_PASSWORD`: HiBob password

---

## 🆘 Troubleshooting

### Browser fails to start (macOS)

**Solution**: Install multiple browsers for fallback
```bash
playwright install chromium firefox webkit
```

### Upload failed - service account error

**Solution**: Share your Google Sheet with service account email
```bash
# Find service account email in service_account.json
cat config/service_account.json | grep client_email
# Share sheet with this email in Google Sheets
```

### "Not Mapped" in AI Readiness Category

**Solution**: Add department to AI Readiness Mapping sheet
1. Check employee's department in Base Data
2. Add row to AI Readiness Mapping: `[Department, AI Category]`
3. Re-run "Build Summary Sheet"

### Manager Blurbs stuck in "Manual Review"

**Causes**:
- Source feedback is sparse or generic
- Feedback contains off-topic content (news articles, etc.)
- Very short feedback (<25 words)

**Solution**: Review source data in Bob Perf Report, manually edit blurbs if needed

---

## 🔄 Updating

### Pull latest code:
```bash
cd "Cloud Agent Space"
git pull origin main
```

### Update Apps Script:
```bash
cd apps-script
npx @google/clasp push
```

### Update Python dependencies:
```bash
pip install -r requirements.txt --upgrade
```

---

## 👨‍💻 Development

### Running locally:
```bash
# Start web app
python3 web/web_app.py

# Test Python scripts
python3 scripts/python/hibob_report_downloader.py
python3 scripts/python/manager_blurb_generator.py
```

### Debugging Apps Script:
- Use Logger.log() for debugging
- View logs: **View** → **Logs** in Apps Script editor

### Code structure:
- `apps-script/bob-performance-module.gs`: Single consolidated Apps Script file
- `web/web_app.py`: Flask server for web interface
- `scripts/python/`: Python automation scripts

---

## 📚 Documentation

- **AI Readiness Categories**: `docs/AI_READINESS_CATEGORIES.md`
- **AI Readiness Implementation**: `docs/AI_READINESS_IMPLEMENTATION.md`
- **Manager Blurb Generator**: `docs/MANAGER_BLURB_GENERATOR.md`

---

## 🤝 Contributing

This is a private project for internal use. For questions or issues, contact the project maintainer.

---

## 📝 License

Internal use only. Not licensed for external distribution.

---

## 🎉 Recent Updates

### November 12, 2025
- ✨ **NEW**: AI Readiness Assessment feature
  - Comprehensive function-to-AI-category mapping (12 functions)
  - Auto-population in Summary sheet
  - Leadership assessment framework
  - Detailed implementation guide
- 🚀 **IMPROVED**: Manager Blurb Generator
  - Lowered semantic threshold (50% → 40%)
  - Reduced word count minimum (30 → 25 words)
  - Added hybrid validation (30%+ semantic + rules = accept)
  - **65.3% valid blurbs** (up from 24.5%)
- 📖 Added comprehensive documentation for AI readiness

---

**Version**: 2.1.0  
**Last Updated**: November 12, 2025  
**Author**: MR
