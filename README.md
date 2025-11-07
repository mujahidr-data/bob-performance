# Bob HR Analytics Project

## Overview
This project provides a comprehensive HR analytics solution that integrates with Bob (hibob.com) to import employee data and generate detailed HR metrics including headcount, turnover, retention, and attrition rates.

## Project Structure
```
Cursor/
├── bob-report-import.gs          # Main Google Apps Script file with all functions
├── specs/
│   └── bob-report-import.md     # Detailed specification and documentation
├── README.md                     # This file - project overview
├── .clasp.json                   # Clasp configuration (links to Google Apps Script)
├── appsscript.json              # Apps Script manifest
└── package.json                 # Node.js dependencies (for clasp)
```

## Features

### 1. Bob Data Import
- Fetches employee data from Bob API
- Filters by employment type (Regular Full-Time, Permanent)
- Automatically populates "RawData" sheet
- On-demand execution via custom menu

### 2. HR Metrics Calculation
- **Opening Headcount**: Employees at start of period
- **Closing Headcount**: Employees at end of period
- **Hires**: New employees in period
- **Terms**: Total terminations (Voluntary & Involuntary)
- **Attrition %**: Termination rate
- **Retention %**: Employee retention rate
- **Turnover %**: Turnover rate

### 3. Advanced Filtering
- Filter by **Site** (e.g., India, UK, USA)
- Filter by **ELT** (Executive Leadership Team)
- Filter by **Department** (e.g., Engineering, Product, Data Science)
- Combine multiple filters for detailed analysis

### 4. Automated Reporting
- Generates monthly metrics from Jan 2024 to current month
- Creates formatted "HR Metrics" sheet
- Auto-updates with latest data

## Setup Instructions

### Prerequisites
- Google Account with access to Google Sheets
- Node.js and npm (for clasp deployment)
- Access to Bob API with credentials

### Initial Setup

1. **Install Dependencies** (if not already done):
   ```bash
   npm install @google/clasp
   ```

2. **Login to Clasp**:
   ```bash
   npx clasp login
   ```

3. **Link to Your Google Sheet's Apps Script**:
   - Update `.clasp.json` with your script ID
   - Or use: `npx clasp clone <script-id>`

4. **Push Code to Apps Script**:
   ```bash
   npx clasp push
   ```

### Google Sheet Setup

1. Open your Google Sheet
2. Ensure you have a sheet named "RawData" (will be created automatically)
3. Refresh the sheet - you should see "Bob HR Analytics" menu
4. Run "Fetch Bob Report" to import data

## Usage Guide

### Menu Options

The "Bob HR Analytics" menu provides:

1. **Fetch Bob Report**
   - Imports latest data from Bob API
   - Filters by employment type
   - Populates "RawData" sheet

2. **Generate HR Metrics**
   - Calculates metrics for all months (Jan 2024 - current)
   - Creates/updates "HR Metrics" sheet
   - Applies filters from "FilterConfig" sheet if present

3. **Create Filter Config Sheet**
   - Creates "FilterConfig" sheet with dropdown menus
   - Select Site, ELT, and/or Department filters
   - Leave blank for "all" (no filter)

4. **Update Filter Options**
   - Creates "FilterOptions" sheet
   - Lists all unique values for review/vetting
   - Helps identify available filter options

### Workflow

#### Basic Usage (No Filters)
1. Run "Fetch Bob Report" to import data
2. Run "Generate HR Metrics" to calculate metrics
3. View results in "HR Metrics" sheet

#### Filtered Analysis
1. Run "Fetch Bob Report" to import data
2. Run "Create Filter Config Sheet" to set up filters
3. Select desired filters in "FilterConfig" sheet
4. Run "Generate HR Metrics" to calculate filtered metrics
5. View filtered results in "HR Metrics" sheet

## Key Functions

### `fetchBobReportCSV()`
Imports employee data from Bob API and populates RawData sheet.

### `calculateHRMetrics(periodStart, periodEnd, filters)`
Core function that calculates all HR metrics for a given period.
- **Parameters**: 
  - `periodStart`: Date object for period start
  - `periodEnd`: Date object for period end
  - `filters`: Optional object `{site, elt, department}`
- **Returns**: Object with all calculated metrics

### `generateHRMetrics()`
Main function to generate comprehensive metrics table for all months.

### `getUniqueFilterValues()`
Extracts unique values for Site, ELT, and Department from RawData.

### `createFilterConfigSheet()`
Creates filter selection interface with dropdown menus.

### `updateFilterOptions()`
Creates sheet listing all unique filter values for review.

## Column Mapping

The script uses the following column indices (0-based):
- **Start Date**: Column C (index 2)
- **Termination Date**: Column O (index 14)
- **Leave/Termination Type**: Column Q (index 16)
- **Department**: Column G (index 6)
- **ELT**: Column H (index 7)
- **Site**: Column I (index 8)

*Note: Adjust `COLUMN_INDICES` constant in code if your data structure differs.*

## Configuration

### Bob API Configuration
- **Endpoint**: `https://api.hibob.com/v1/company/reports/31048356/download?format=csv&includeInfo=true`
- **Authentication**: Basic Auth (configured in code)
- **Report ID**: 31048356

### Employment Type Filter
Currently filters for:
- "Regular Full-Time"
- "Permanent"

*Modify `allowedTypes` array in `fetchBobReportCSV()` to change.*

## Deployment

### Using Clasp (Recommended)
```bash
# Make changes to bob-report-import.gs
npx clasp push
```

### Manual Deployment
1. Open Google Apps Script editor (Extensions → Apps Script)
2. Copy code from `bob-report-import.gs`
3. Paste into editor
4. Save (Cmd+S / Ctrl+S)

## Sheets Created

The script automatically creates/updates these sheets:

1. **RawData**: Imported employee data from Bob
2. **HR Metrics**: Calculated metrics table
3. **FilterConfig**: Filter selection interface (optional)
4. **FilterOptions**: Unique filter values listing (optional)

## Troubleshooting

### "RawData sheet not found"
- Run "Fetch Bob Report" first to create and populate the sheet

### Metrics showing zeros
- Verify data exists in RawData sheet
- Check date columns are properly formatted
- Ensure employment type filtering is correct

### Filter dropdowns empty
- Run "Fetch Bob Report" to import data first
- Then run "Create Filter Config Sheet"

### Column index errors
- Verify your data structure matches the column mapping
- Adjust `COLUMN_INDICES` constant if needed

## Future Enhancements

Potential improvements:
- Custom date range selection for metrics
- Export to PDF/Excel functionality
- Additional metrics (diversity, tenure, etc.)
- Automated email reports
- Dashboard visualization
- Multi-sheet analysis

## Version History

- **v1.0** (Current): Initial implementation with basic HR metrics and filtering
  - Bob data import
  - HR metrics calculation
  - Filtering by Site, ELT, Department
  - Automated monthly reporting

## Support

For issues or questions:
1. Check the specification document: `specs/bob-report-import.md`
2. Review function comments in `bob-report-import.gs`
3. Check Google Apps Script execution logs

## License

Internal use only - CommerceIQ HR Analytics

---

**Last Updated**: November 2024  
**Maintained By**: HR Analytics Team

