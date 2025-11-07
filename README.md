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
- Supports aggregation by Halves (H1/H2) and Quarters (Q1-Q4)
- Creates formatted output sheets with preserved formatting
- Auto-updates with latest data
- Automatically hides processing sheets (RawData, FilterOptions, Termination Reason Mapping) after processing

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

2. **Generate All CIQ**
   - Calculates overall company metrics for all months (Jan 2024 - current)
   - Creates/updates "All CIQ" sheet
   - Supports aggregation by Halves (H1/H2) or Quarters (Q1-Q4) when Year filter is selected
   - Applies time period filters from "FilterConfig" sheet if present

3. **Generate Headcount by Site**
   - Generates site-wise headcount metrics month-on-month
   - Creates/updates "Headcount by Site" sheet
   - Supports aggregation by Halves or Quarters

4. **Generate Headcount by ELT**
   - Generates ELT-wise headcount metrics month-on-month
   - Creates/updates "Headcount by ELT" sheet
   - Supports aggregation by Halves or Quarters

5. **Generate Headcount by Job Level**
   - Generates job level headcount as of today with site-level breakdowns
   - Creates/updates "Headcount by Job Level" sheet
   - Includes interactive charts and ELT filter dropdown

6. **Generate Terminations Reasons Drilldown**
   - Generates termination reasons breakdown (Overall, by Site, by ELT)
   - Creates/updates "Terminations Reasons Drilldown" sheet
   - Supports filtering by Year + Period Type (Halves or Quarters)
   - Fully clears and reformats sheet when filters are applied

7. **Create Filter Config Sheet**
   - Creates "FilterConfig" sheet with dropdown menus
   - Select Site, ELT, Department, and Time Period filters
   - Select Year + Halves/Quarters for aggregated reporting
   - Leave blank for "all" (no filter)

8. **Update Filter Options**
   - Creates "FilterOptions" sheet
   - Lists all unique values for review/vetting
   - Helps identify available filter options

### Workflow

#### Basic Usage (No Filters)
1. Run "Fetch Bob Report" to import data
2. Run "Generate All CIQ" to calculate overall company metrics
3. View results in "All CIQ" sheet

#### Filtered Analysis
1. Run "Fetch Bob Report" to import data
2. Run "Create Filter Config Sheet" to set up filters
3. Select desired filters in "FilterConfig" sheet:
   - Time Period: Year, Year + Halves, or Year + Quarters
   - Site, ELT, Department (optional)
4. Run appropriate generate function:
   - "Generate All CIQ" for overall company metrics
   - "Generate Headcount by Site" for site breakdowns
   - "Generate Headcount by ELT" for ELT breakdowns
   - "Generate Terminations Reasons Drilldown" for termination analysis
5. View filtered results in respective sheets

#### Aggregated Reporting (Halves/Quarters)
1. Run "Fetch Bob Report" to import data
2. Run "Create Filter Config Sheet"
3. In "FilterConfig" sheet:
   - Select a Year
   - Select "Halves" to aggregate by H1/H2, or "Quarters" to aggregate by Q1-Q4
4. Run "Generate All CIQ", "Generate Headcount by Site", or "Generate Headcount by ELT"
5. View aggregated results (existing data is cleared but formatting is preserved)

## Key Functions

### `fetchBobReportCSV()`
Imports employee data from Bob API and populates RawData sheet.

### `calculateHRMetrics(periodStart, periodEnd, filters)`
Core function that calculates all HR metrics for a given period.
- **Parameters**: 
  - `periodStart`: Date object for period start
  - `periodEnd`: Date object for period end
  - `filters`: Optional object `{site, elt, department, terminationReason}`
- **Returns**: Object with all calculated metrics

### `generateAllCIQ()`
Generates overall company HR metrics for multiple periods. Creates/updates "All CIQ" sheet.
- Supports aggregation by Halves (H1/H2) or Quarters (Q1-Q4)
- Clears existing data but preserves formatting when aggregation is used

### `generateHeadcountBySite()`
Generates site-wise headcount metrics month-on-month. Creates/updates "Headcount by Site" sheet.
- Supports aggregation by Halves or Quarters
- Clears existing data but preserves formatting when aggregation is used

### `generateHeadcountByELT()`
Generates ELT-wise headcount metrics month-on-month. Creates/updates "Headcount by ELT" sheet.
- Supports aggregation by Halves or Quarters
- Clears existing data but preserves formatting when aggregation is used

### `generateHeadcountByJobLevel()`
Generates job level headcount as of today with site-level breakdowns. Creates/updates "Headcount by Job Level" sheet.
- Includes interactive charts and ELT filter dropdown

### `generateTerminationsReasonsDrilldown()`
Generates termination reasons breakdown (Overall, by Site, by ELT). Creates/updates "Terminations Reasons Drilldown" sheet.
- Supports filtering by Year + Period Type (Halves or Quarters)
- Fully clears sheet (data + formatting) when filters are applied to prevent formatting issues

### `getUniqueFilterValues()`
Extracts unique values for Site, ELT, Department, and Termination Reasons from RawData.

### `createFilterConfigSheet()`
Creates filter selection interface with dropdown menus for all filter types.

### `updateFilterOptions()`
Creates sheet listing all unique filter values for review.

### `hideProcessingSheets(ss)`
Automatically hides processing sheets (RawData, FilterOptions, Termination Reason Mapping) after processing is completed.

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

1. **RawData**: Imported employee data from Bob (hidden after processing)
2. **All CIQ**: Overall company HR metrics table
3. **Headcount by Site**: Site-wise headcount metrics
4. **Headcount by ELT**: ELT-wise headcount metrics
5. **Headcount by Job Level**: Job level headcount with charts
6. **Terminations Reasons Drilldown**: Termination reasons breakdown (Overall, by Site, by ELT)
7. **FilterConfig**: Filter selection interface (optional)
8. **FilterOptions**: Unique filter values listing (optional, hidden after processing)
9. **Termination Reason Mapping**: Mapping sheet for cleaning termination reasons (optional, hidden after processing)

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

- **v2.0** (Current): Enhanced reporting with aggregation and improved formatting
  - Renamed sheets and functions for clarity:
    - Overall Data → All CIQ
    - Headcount Metrics → Headcount by Site
    - ELT Metrics → Headcount by ELT
    - Job Level Headcount → Headcount by Job Level
    - Termination Reasons → Terminations Reasons Drilldown
  - Added Quarter/Halves aggregation support
  - Fixed formatting issues in Termination Reasons when filters are applied
  - Automatic hiding of processing sheets after completion
  - Improved data clearing logic (preserves formatting for aggregated reports)

- **v1.0**: Initial implementation with basic HR metrics and filtering
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

**Last Updated**: December 2024  
**Maintained By**: HR Analytics Team

