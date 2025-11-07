# Bob Report Import - Specification

## Overview
This is an import function from Bob (hibob.com) to get employee base data into Google Sheets.

## Function: `fetchBobReportCSV()`

### Purpose
Fetches employee data from Bob API, filters by employment type, and populates a Google Sheet.

### Details

**Target Sheet:** "RawData"

**API Endpoint:** `https://api.hibob.com/v1/company/reports/31048356/download?format=csv&includeInfo=true`

**Authentication:** Basic Auth
- Authorization header: `Basic U0VSVklDRS0yMzkyODpVZDN1RjRESFpBY2l1NFI2WWk5ek9oWHF2d2VUMlE3NE1vRGd0eUxT`

**Filtering Logic:**
- Filters rows by employment type (Column K, index 9)
- Allowed employment types: `["Regular Full-Time", "Permanent"]`
- Keeps header row, filters data rows

**Operations:**
1. Fetches CSV from Bob API
2. Parses CSV data
3. Filters by employment type
4. Clears existing sheet content
5. Writes filtered data to sheet
6. Auto-resizes columns

### Current Trigger
- Time-based trigger (scheduled)

### Requirements
- Add `onOpen` function to execute workflow on demand
- Allow manual execution from menu

## Code

```javascript
function fetchBobReportCSV() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RawData");
  const url = "https://api.hibob.com/v1/company/reports/31048356/download?format=csv&includeInfo=true";
  const headers = {
    "accept": "application/json",
    "authorization": "Basic U0VSVklDRS0yMzkyODpVZDN1RjRESFpBY2l1NFI2WWk5ek9oWHF2d2VUMlE3NE1vRGd0eUxT"
  };
  const options = {
    method: "get",
    headers: headers,
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  const csv = response.getContentText();
  const data = Utilities.parseCsv(csv);
  if (!data || data.length === 0) {
    Logger.log("No data received or CSV is empty.");
    return;
  }
  const header = data[0];
  const allowedTypes = ["Regular Full-Time", "Permanent",];
  const employmentTypeCol = 9; // Column K = index 10
  // Log first 10 employment types to help debugging
  for (let i = 1; i < Math.min(10, data.length); i++) {
    Logger.log(`Row ${i + 1} Employment Type: ${data[i][employmentTypeCol]}`);
  }
  // Filter to only allowed employment types
  const filteredData = data.filter((row, index) => {
    if (index === 0) return true; // keep header
    const empType = row[employmentTypeCol]?.trim();
    return allowedTypes.includes(empType);
  });
  // Clear previous content
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearContent();
  // Write filtered data
  const targetRange = sheet.getRange(1, 1, filteredData.length, filteredData[0].length);
  targetRange.setValues(filteredData);
  // Auto-resize columns
  sheet.autoResizeColumns(1, filteredData[0].length);
  Logger.log(`Bob report updated. Rows after filtering: ${filteredData.length - 1}`);
}
```

## Notes
- Employment type column is at index 9 (Column K)
- Debug logging included for first 10 rows
- Currently triggered by time-based schedule
- Needs on-demand execution capability

---

## HR Analytics Functions

### Overview
The script now includes comprehensive HR analytics capabilities that calculate headcount metrics, turnover, retention, and attrition rates. All formulas have been converted to Google Apps Script functions.

### Column Mapping
Based on the original formulas, the following column indices are used:
- **Start Date**: Column C (index 2)
- **Termination Date**: Column O (index 14)
- **Leave/Termination Type**: Column Q (index 16)
- **Department**: Column G (index 6)
- **ELT**: Column H (index 7)
- **Site**: Column I (index 8)

### Functions

#### `calculateHRMetrics(periodStart, periodEnd, filters)`
Core function that calculates all HR metrics for a given time period.

**Parameters:**
- `periodStart` (Date): Start date of the period
- `periodEnd` (Date): End date of the period
- `filters` (Object, optional): Filter object with `{site: string, elt: string, department: string}`

**Returns:**
- `openingHC`: Opening headcount
- `closingHC`: Closing headcount
- `hires`: Number of hires in period
- `terms`: Total terminations
- `voluntaryTerms`: Voluntary terminations
- `involuntaryTerms`: Involuntary terminations
- `attrition`: Attrition percentage
- `retention`: Retention percentage
- `turnover`: Turnover percentage

**Metrics Logic:**
- **Opening HC**: Employees who started ≤ periodStart AND (terminated ≥ periodStart OR not terminated)
- **Closing HC**: Employees who started ≤ periodEnd AND (terminated > periodEnd OR not terminated)
- **Hires**: Employees who started between periodStart and periodEnd
- **Terms**: Employees who terminated between periodStart and periodEnd
- **Voluntary Terms**: Includes "Voluntary", "Voluntary - Regrettable", "Voluntary - Non regrettable"
- **Involuntary Terms**: Includes "Involuntary", "Involuntary - Regrettable", "End of Contract"
- **Attrition**: `(terms / avgHC) * 100` where avgHC = (openingHC + closingHC) / 2
- **Retention**: `((openingHC - terms) / openingHC) * 100`
- **Turnover**: `(terms / avgHC) * 100`

#### `generateAllCIQ()`
Main function to generate overall company HR metrics table. Creates/updates "All CIQ" sheet with monthly metrics from Jan 2024 to current month.

**Features:**
- Automatically reads filter selections from "FilterConfig" sheet if it exists
- Generates monthly periods automatically
- Supports aggregation by Halves (H1/H2) or Quarters (Q1-Q4) when Year + Halves/Quarters is selected
- Formats percentage columns
- Displays applied filters at bottom of sheet
- Clears existing data but preserves formatting when aggregation is used

#### `generateHeadcountBySite()`
Generates site-wise headcount metrics month-on-month. Creates/updates "Headcount by Site" sheet.

**Features:**
- Supports aggregation by Halves or Quarters
- Clears existing data but preserves formatting when aggregation is used

#### `generateHeadcountByELT()`
Generates ELT-wise headcount metrics month-on-month. Creates/updates "Headcount by ELT" sheet.

**Features:**
- Supports aggregation by Halves or Quarters
- Clears existing data but preserves formatting when aggregation is used

#### `generateHeadcountByJobLevel()`
Generates job level headcount as of today with site-level breakdowns. Creates/updates "Headcount by Job Level" sheet.

**Features:**
- Includes interactive charts and ELT filter dropdown

#### `generateTerminationsReasonsDrilldown()`
Generates termination reasons breakdown (Overall, by Site, by ELT). Creates/updates "Terminations Reasons Drilldown" sheet.

**Features:**
- Supports filtering by Year + Period Type (Halves or Quarters)
- Fully clears sheet (data + formatting) when filters are applied to prevent formatting issues

#### `getUniqueFilterValues()`
Extracts unique values for Site, ELT, Department, and Termination Reasons from RawData sheet.

**Returns:**
- `sites`: Array of unique site values
- `elts`: Array of unique ELT values
- `departments`: Array of unique department values
- `terminationReasons`: Array of unique termination reason values

#### `updateFilterOptions()`
Creates/updates "FilterOptions" sheet with all unique filter values for review and vetting.

#### `createFilterConfigSheet()`
Creates a "FilterConfig" sheet with dropdown menus for selecting filters. Users can:
1. Select Site, ELT, Department, and/or Time Period from dropdowns
2. Select Year + Halves/Quarters for aggregated reporting
3. Leave blank for "all" (no filter)
4. Run appropriate generate function to apply selected filters

#### `hideProcessingSheets(ss)`
Automatically hides processing sheets (RawData, FilterOptions, Termination Reason Mapping) after processing is completed.

### Menu Structure
The `onOpen()` function creates a "Bob HR Analytics" menu with:
- **Fetch Bob Report**: Imports data from Bob API
- **Generate All CIQ**: Calculates and displays overall company metrics
- **Generate Headcount by Site**: Generates site-wise headcount metrics
- **Generate Headcount by ELT**: Generates ELT-wise headcount metrics
- **Generate Headcount by Job Level**: Generates job level headcount with charts
- **Generate Terminations Reasons Drilldown**: Generates termination reasons breakdown
- **Create Filter Config Sheet**: Sets up filter selection interface
- **Update Filter Options**: Refreshes unique filter values

### Workflow

1. **Initial Setup:**
   - Run "Fetch Bob Report" to import data
   - Run "Create Filter Config Sheet" to set up filters
   - Run "Update Filter Options" to see all available filter values

2. **Generate Metrics:**
   - Select filters in "FilterConfig" sheet (or leave blank for all)
   - Run appropriate generate function from menu:
     - "Generate All CIQ" for overall company metrics
     - "Generate Headcount by Site" for site breakdowns
     - "Generate Headcount by ELT" for ELT breakdowns
     - "Generate Terminations Reasons Drilldown" for termination analysis
   - View results in respective sheets

3. **Filtering:**
   - Filter values are extracted automatically from RawData
   - Users can review unique values in "FilterOptions" sheet
   - Filters can be applied via "FilterConfig" sheet dropdowns
   - Multiple filters can be combined (AND logic)

### Sheets Created
- **RawData**: Contains imported Bob data (hidden after processing)
- **All CIQ**: Overall company HR metrics table (auto-generated)
- **Headcount by Site**: Site-wise headcount metrics (auto-generated)
- **Headcount by ELT**: ELT-wise headcount metrics (auto-generated)
- **Headcount by Job Level**: Job level headcount with charts (auto-generated)
- **Terminations Reasons Drilldown**: Termination reasons breakdown (auto-generated)
- **FilterOptions**: Contains unique filter values for review (optional, hidden after processing)
- **FilterConfig**: Contains filter selection dropdowns (optional)
- **Termination Reason Mapping**: Mapping sheet for cleaning termination reasons (optional, hidden after processing)

### Notes
- All date comparisons use inclusive start dates and exclusive end dates for closing headcount
- Percentage values are rounded to 1 decimal place
- Empty termination dates are treated as active employees
- Filters use exact string matching (case-sensitive)

