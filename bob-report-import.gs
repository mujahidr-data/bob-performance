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
  
  // Reassign ELT values based on business rules
  reassignELTValues();
  
  // Fix department name typos
  fixDepartmentTypos();
}

/**
 * Creates a custom menu when the spreadsheet is opened
 * Allows on-demand execution of the Bob report import and analytics
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Bob HR Analytics')
    // Step 1: Import data
    .addItem('Step 1: Fetch Bob Report', 'fetchBobReportCSV')
    .addSeparator()
    // Step 2: Process/clean data
    .addItem('Step 2: Reassign ELT Values', 'reassignELTValues')
    .addItem('Step 2b: Fix Department Typos', 'fixDepartmentTypos')
    .addItem('Step 2c: Create Termination Reason Mapping Sheet', 'createTerminationReasonMappingSheet')
    .addItem('Step 2d: Map Termination Reasons', 'mapTerminationReasons')
    .addSeparator()
    // Step 3: Review available filters
    .addItem('Step 3: Update Filter Options', 'updateFilterOptions')
    // Step 4: Set up filter selection
    .addItem('Step 4: Create Filter Config Sheet', 'createFilterConfigSheet')
    .addSeparator()
    // Step 5: Generate metrics
    .addItem('Step 5: Generate Overall Data', 'generateOverallData')
    .addSeparator()
    // Additional reports
    .addItem('Generate Headcount Metrics', 'generateHeadcountMetrics')
    .addItem('Generate ELT Metrics', 'generateELTMetrics')
    .addItem('Generate Job Level Headcount', 'generateJobLevelHeadcount')
    .addItem('Generate Termination Reasons Table', 'generateTerminationReasonsTable')
    .addToUi();
}

/**
 * onEdit trigger to automatically update ELT chart when dropdown filter changes
 * This function is called automatically when any cell is edited in the spreadsheet
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // Check if the edit is in the "Job Level Headcount" sheet
  if (sheet.getName() !== "Job Level Headcount") {
    return;
  }
  
  // Check if the edit is in the ELT filter dropdown cell (column B, row with "Filter by ELT:")
  // We need to find which row has "Filter by ELT:" label
  try {
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // Find the row with "Filter by ELT:" label
    let eltFilterRow = null;
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] && String(values[i][0]).includes("Filter by ELT:")) {
        eltFilterRow = i + 1; // Convert to 1-indexed
        break;
      }
    }
    
    // If we found the filter row and the edit is in column B of that row
    if (eltFilterRow && range.getRow() === eltFilterRow && range.getColumn() === 2) {
      // Update the ELT chart
      updateELTChart(sheet, eltFilterRow);
    }
  } catch (error) {
    Logger.log(`Error in onEdit trigger: ${error.message}`);
    // Don't show error to user, just log it
  }
}

/**
 * Updates the ELT chart based on the dropdown filter selection
 */
function updateELTChart(jobLevelSheet, eltFilterRow) {
  try {
    const filterValue = jobLevelSheet.getRange(eltFilterRow, 2).getValue();
    
    // Find the ELT chart
    const existingCharts = jobLevelSheet.getCharts();
    let eltChart = null;
    
    existingCharts.forEach(chart => {
      try {
        const options = chart.getOptions();
        const title = options ? options.title : null;
        if (title && title.includes('Job Level Headcount by ELT')) {
          eltChart = chart;
        }
      } catch (e) {
        Logger.log(`Error reading chart: ${e.message}`);
      }
    });
    
    if (!eltChart) {
      Logger.log("ELT chart not found, skipping update");
      return;
    }
    
    // Find the filtered data range
    // The filtered data table should be below the ELT breakdown table
    const dataRange = jobLevelSheet.getDataRange();
    const values = dataRange.getValues();
    
    // Find the filtered headers row - look for "Job Level" in column A that's below the ELT breakdown table
    // First, find the ELT breakdown table to know where to look below it
    let eltBreakdownStartRow = null;
    let eltBreakdownEndRow = null;
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      // Look for a row that has "Job Level" as first cell and ELT names in subsequent cells (this is the ELT breakdown header)
      if (row[0] === "Job Level" && row.length > 2 && !eltBreakdownStartRow) {
        eltBreakdownStartRow = i + 1; // Convert to 1-indexed
        // Count how many data rows follow
        let dataRowCount = 0;
        for (let j = i + 1; j < values.length; j++) {
          if (values[j][0] && String(values[j][0]).trim() && !String(values[j][0]).includes(":")) {
            dataRowCount++;
          } else {
            break;
          }
        }
        eltBreakdownEndRow = i + 1 + dataRowCount; // End row of ELT breakdown table
        break;
      }
    }
    
    if (!eltBreakdownStartRow) {
      Logger.log("ELT breakdown table not found, skipping update");
      return;
    }
    
    // Now find the filtered data table header row (should be below the ELT breakdown table)
    // Look for "Job Level" in column A that appears after the ELT breakdown table
    let filteredHeadersRow = null;
    for (let i = eltBreakdownEndRow; i < values.length; i++) {
      if (values[i][0] === "Job Level") {
        filteredHeadersRow = i + 1; // Convert to 1-indexed
        break;
      }
    }
    
    if (!filteredHeadersRow) {
      Logger.log("Filtered data table header not found, skipping update");
      return;
    }
    
    // Count job level rows in the filtered data table (until we hit an empty row or end of data)
    let jobLevelCount = 0;
    for (let i = filteredHeadersRow; i < values.length; i++) {
      if (values[i][0] && String(values[i][0]).trim() && values[i][0] !== "Job Level") {
        jobLevelCount++;
      } else if (values[i][0] === "Job Level") {
        // Skip the header row
        continue;
      } else {
        break;
      }
    }
    
    // Determine how many columns to include based on filter
    // If "All", include all ELT columns. Otherwise, count non-empty columns in filtered data
    let numCols = 2; // At least Job Level + 1 ELT column
    if (filterValue === "All") {
      // Count ELT columns from the original breakdown table
      const headerRow = values[eltBreakdownStartRow - 1];
      numCols = headerRow.length - 1; // Exclude "Total" column
    } else {
      // When filtering to one ELT, we need Job Level + that ELT column
      // But the filtered table still has all columns (with formulas showing blanks)
      // So we'll use the same number of columns as the original table
      const headerRow = values[eltBreakdownStartRow - 1];
      numCols = headerRow.length - 1; // Exclude "Total" column
    }
    
    // Create the filtered chart range (header row + all data rows)
    // filteredHeadersRow is the header row, data rows start at filteredHeadersRow + 1
    const filteredChartRange = jobLevelSheet.getRange(filteredHeadersRow, 1, jobLevelCount + 1, numCols);
    
    // Update the chart
    try {
      const updatedChart = eltChart.modify()
        .clearRanges()
        .addRange(filteredChartRange)
        .build();
      jobLevelSheet.updateChart(updatedChart);
      Logger.log(`ELT chart updated for filter: ${filterValue}`);
    } catch (e) {
      Logger.log(`Error updating ELT chart: ${e.message}`);
    }
  } catch (error) {
    Logger.log(`Error in updateELTChart: ${error.message}`);
  }
}

/**
 * Column indices mapping based on Bob import structure
 * Based on formulas: C:C (Start Date), O:O (Termination Date), Q:Q (Leave/Termination Type)
 * Adjust these if your column structure differs
 * Note: Google Sheets uses 1-based columns, but arrays are 0-indexed
 */
const COLUMN_INDICES = {
  EMP_NAME: 0,             // Column A (0-indexed: 0) - Employee Name (used for unique counting)
  START_DATE: 2,           // Column C (0-indexed: 2) - Start Date in YYYY-MM-DD format
  TERMINATION_DATE: 17,    // Column R (0-indexed: 17) - Termination date (blank for active employees)
  LEAVE_TERMINATION_TYPE: 19, // Column T (0-indexed: 19) - Leave and termination type
  TERMINATION_CATEGORY: 20, // Column U (0-indexed: 20) - Termination Category (Regretted/Unregretted)
  TERMINATION_REASON: 18,   // Column S (0-indexed: 18) - Reason for termination
  JOB_LEVEL: 4,            // Column E (0-indexed: 4) - Job Level
  DEPARTMENT: 6,           // Column G (0-indexed: 6)
  ELT: 7,                  // Column H (0-indexed: 7)
  SITE: 8,                 // Column I (0-indexed: 8)
  STATUS: 16               // Column Q (0-indexed: 16) - Status (Active/Inactive, Employed/Terminated)
};

// Job level sorting order
const JOB_LEVEL_ORDER = [
  "L2 IC", "L3 IC", "L4 IC", "L5 IC", "L5.5 IC", "L6 IC", "L6.5 IC", "L7 IC",
  "L4 Mgr", "L5 Mgr", "L5.5 Mgr", "L6 Mgr", "L6.5 Mgr", "L7 Mgr", "L8 Mgr", "L9 Mgr", "L10 Mgr"
];

/**
 * Reassigns ELT values based on business rules:
 * - All "Heather" → "Jyoti"
 * - All "Moni" → "Gaurav" if Department is "Engineering" or "Product Success"
 * - All "Moni" → "Himanshu" if Department is "Product Management" or "Product Design"
 */
function reassignELTValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawDataSheet = ss.getSheetByName("RawData");
  
  if (!rawDataSheet) {
    SpreadsheetApp.getUi().alert("RawData sheet not found. Please run 'Fetch Bob Report' first.");
    return;
  }
  
  const data = rawDataSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert("No data found in RawData sheet.");
    return;
  }
  
  const header = data[0];
  const rows = data.slice(1);
  let reassignmentCount = 0;
  const reassignments = [];
  
  // First, find the exact ELT names for Jyoti, Gaurav, and Himanshu from existing data
  const allELTs = new Set();
  rows.forEach(row => {
    const elt = String(row[COLUMN_INDICES.ELT] || "").trim();
    if (elt) allELTs.add(elt);
  });
  
  // Find exact names (case-insensitive search)
  let jyotiName = null;
  let gauravName = null;
  let himanshuName = null;
  
  Array.from(allELTs).forEach(elt => {
    const eltLower = elt.toLowerCase();
    if (eltLower.startsWith("jyoti") && !jyotiName) {
      jyotiName = elt;
    } else if (eltLower.startsWith("gaurav") && !gauravName) {
      gauravName = elt;
    } else if (eltLower.startsWith("himanshu") && !himanshuName) {
      himanshuName = elt;
    }
  });
  
  // Default to simple names if not found in data
  if (!jyotiName) jyotiName = "Jyoti Gouri";
  if (!gauravName) gauravName = "Gaurav Khanna";
  if (!himanshuName) himanshuName = "Himanshu Jain";
  
  Logger.log(`Found ELT names: Jyoti="${jyotiName}", Gaurav="${gauravName}", Himanshu="${himanshuName}"`);
  
  // Process each row (skip header)
  rows.forEach((row, index) => {
    // Get ELT and Department values, handling various formats
    let currentELT = row[COLUMN_INDICES.ELT];
    if (currentELT instanceof Date) {
      currentELT = Utilities.formatDate(currentELT, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    currentELT = String(currentELT || "").trim();
    
    let department = row[COLUMN_INDICES.DEPARTMENT];
    if (department instanceof Date) {
      department = Utilities.formatDate(department, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    department = String(department || "").trim();
    
    let newELT = null;
    const eltLower = currentELT.toLowerCase();
    const deptLower = department.toLowerCase();
    
    // Rule 1: All "Heather" (or "Heather DeJong") → Jyoti (exact name from data)
    if (eltLower.startsWith("heather")) {
      newELT = jyotiName;
    }
    // Rule 2: All "Moni" (or "Moni Manor") → Assign based on department
    else if (eltLower.startsWith("moni")) {
      // Moni → Gaurav for: Engineering, Product Success, Data Science
      if (deptLower === "engineering" || deptLower === "product success" || deptLower === "data science") {
        newELT = gauravName;
      }
      // Moni → Himanshu for: Product Management, Product Design
      else if (deptLower === "product management" || deptLower === "product design") {
        newELT = himanshuName;
      }
      // For any other department with Moni, default to Gaurav
      else if (deptLower) {
        newELT = gauravName;
      }
    }
    
    // Update the row if reassignment is needed
    if (newELT) {
      row[COLUMN_INDICES.ELT] = newELT;
      reassignmentCount++;
      reassignments.push({
        row: index + 2, // +2 because: +1 for header, +1 for 0-index
        oldELT: currentELT,
        newELT: newELT,
        department: department
      });
    }
  });
  
  // Write updated data back to sheet
  if (reassignmentCount > 0) {
    const updatedData = [header, ...rows];
    const targetRange = rawDataSheet.getRange(1, 1, updatedData.length, updatedData[0].length);
    targetRange.setValues(updatedData);
    
    // Log reassignments
    Logger.log(`ELT reassignments completed: ${reassignmentCount} rows updated`);
    reassignments.forEach(r => {
      Logger.log(`Row ${r.row}: ${r.oldELT} → ${r.newELT} (Dept: ${r.department})`);
    });
    
    SpreadsheetApp.getUi().alert(`ELT reassignments completed.\n${reassignmentCount} rows updated.`);
  } else {
    // Debug: Check if Heather or Moni exist at all
    const allELTs = new Set();
    rows.forEach(row => {
      const elt = String(row[COLUMN_INDICES.ELT] || "").trim();
      if (elt) allELTs.add(elt);
    });
    
    const hasHeather = Array.from(allELTs).some(elt => elt.toLowerCase().includes("heather"));
    const hasMoni = Array.from(allELTs).some(elt => elt.toLowerCase().includes("moni"));
    
    let debugMsg = "No ELT reassignments needed.";
    if (hasHeather || hasMoni) {
      debugMsg += `\n\nDebug: Found ${hasHeather ? 'Heather' : ''} ${hasMoni ? 'Moni' : ''} in data but no matches.`;
      debugMsg += `\nCheck logs for column index and sample values.`;
      Logger.log(`All unique ELT values: ${Array.from(allELTs).join(", ")}`);
    }
    
    SpreadsheetApp.getUi().alert(debugMsg);
  }
}

/**
 * Fixes department name typos in the RawData sheet
 * Currently fixes: "Customer Sucess" → "Customer Success"
 */
function fixDepartmentTypos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawDataSheet = ss.getSheetByName("RawData");
  
  if (!rawDataSheet) {
    SpreadsheetApp.getUi().alert("RawData sheet not found. Please run 'Fetch Bob Report' first.");
    return;
  }
  
  const data = rawDataSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert("No data found in RawData sheet.");
    return;
  }
  
  const header = data[0];
  const rows = data.slice(1);
  let fixCount = 0;
  const fixes = [];
  
  // Department typo corrections (case-insensitive)
  const departmentFixes = {
    "customer sucess": "Customer Success"  // Fix typo: missing 'c'
  };
  
  // Process each row (skip header)
  rows.forEach((row, index) => {
    let department = row[COLUMN_INDICES.DEPARTMENT];
    if (department instanceof Date) {
      department = Utilities.formatDate(department, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    const originalDept = String(department || "").trim();
    const deptLower = originalDept.toLowerCase();
    
    // Check if this department needs fixing
    if (departmentFixes.hasOwnProperty(deptLower)) {
      const correctedDept = departmentFixes[deptLower];
      row[COLUMN_INDICES.DEPARTMENT] = correctedDept;
      fixCount++;
      fixes.push({
        row: index + 2, // +2 because: +1 for header, +1 for 0-index
        oldDept: originalDept,
        newDept: correctedDept
      });
    }
  });
  
  // Write updated data back to sheet
  if (fixCount > 0) {
    const updatedData = [header, ...rows];
    const targetRange = rawDataSheet.getRange(1, 1, updatedData.length, updatedData[0].length);
    targetRange.setValues(updatedData);
    
    // Log fixes
    Logger.log(`Department typo fixes completed: ${fixCount} rows updated`);
    fixes.forEach(f => {
      Logger.log(`Row ${f.row}: "${f.oldDept}" → "${f.newDept}"`);
    });
    
    SpreadsheetApp.getUi().alert(`Department typos fixed.\n${fixCount} rows updated.`);
  } else {
    SpreadsheetApp.getUi().alert("No department typos found. All department names are correct.");
  }
}

/**
 * Generates HR metrics for a given period with optional filters
 * @param {Date} periodStart - Start date of the period
 * @param {Date} periodEnd - End date of the period
 * @param {Object} filters - Optional filters {site: string, elt: string, department: string}
 * @returns {Object} Metrics object with all calculated values
 */
function calculateHRMetrics(periodStart, periodEnd, filters = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawDataSheet = ss.getSheetByName("RawData");
  
  if (!rawDataSheet) {
    throw new Error("RawData sheet not found. Please run fetchBobReportCSV first.");
  }
  
  const data = rawDataSheet.getDataRange().getValues();
  if (data.length <= 1) {
    throw new Error("No data found in RawData sheet.");
  }
  
  const header = data[0];
  const rows = data.slice(1);
  
  // Apply filters (with trimming and normalization for robust matching)
  let filteredRows = rows;
  const initialRowCount = filteredRows.length;
  
  if (filters.site) {
    const filterSite = String(filters.site).trim();
    filteredRows = filteredRows.filter(row => {
      const rowSite = String(row[COLUMN_INDICES.SITE] || "").trim();
      return rowSite === filterSite;
    });
    Logger.log(`After Site filter ("${filterSite}"): ${filteredRows.length} rows (from ${initialRowCount})`);
  }
  if (filters.elt) {
    const filterELT = String(filters.elt).trim();
    const beforeELT = filteredRows.length;
    filteredRows = filteredRows.filter(row => {
      const rowELT = String(row[COLUMN_INDICES.ELT] || "").trim();
      return rowELT === filterELT;
    });
    Logger.log(`After ELT filter ("${filterELT}"): ${filteredRows.length} rows (from ${beforeELT})`);
  }
  if (filters.department) {
    const filterDept = String(filters.department).trim();
    const beforeDept = filteredRows.length;
    filteredRows = filteredRows.filter(row => {
      const rowDept = String(row[COLUMN_INDICES.DEPARTMENT] || "").trim();
      return rowDept === filterDept;
    });
    Logger.log(`After Department filter ("${filterDept}"): ${filteredRows.length} rows (from ${beforeDept})`);
  }
  
  // Note: terminationReason filter is applied only to termination calculations, not headcount
  // This allows filtering terminations by reason while keeping headcount accurate
  
  if (filteredRows.length === 0) {
    Logger.log(`WARNING: No rows match the filters! Filters: ${JSON.stringify(filters)}`);
  }
  
  // Helper function to parse date and normalize to date only (no time)
  // Handles various formats: Date objects, strings like "1/1/2024", "2024-01-01", etc.
  const parseDate = (dateValue) => {
    if (!dateValue) return null;
    
    let date;
    
    // If it's already a Date object
    if (dateValue instanceof Date) {
      date = dateValue;
    }
    // If it's a number (serial date from Excel/Sheets)
    else if (typeof dateValue === 'number') {
      // Google Sheets serial dates start from 1899-12-30
      const sheetsEpoch = new Date(1899, 11, 30);
      date = new Date(sheetsEpoch.getTime() + dateValue * 86400000);
    }
    // If it's a string, try to parse it
    else if (typeof dateValue === 'string') {
      const trimmed = dateValue.trim();
      // Handle formats like "1/1/2024", "01/01/2024", "2024-01-01"
      if (trimmed) {
        date = new Date(trimmed);
        // If parsing failed, try alternative formats
        if (isNaN(date.getTime())) {
          // Try MM/DD/YYYY format explicitly
          const parts = trimmed.split('/');
          if (parts.length === 3) {
            const month = parseInt(parts[0], 10) - 1; // Month is 0-indexed
            const day = parseInt(parts[1], 10);
            const year = parseInt(parts[2], 10);
            date = new Date(year, month, day);
          }
        }
      }
    }
    
    // Normalize to date only (set time to midnight) and validate
    if (date && !isNaN(date.getTime()) && date instanceof Date) {
      return new Date(date.getFullYear(), date.getMonth(), date.getDate());
    }
    
    return null;
  };
  
  // Normalize period dates to date only
  const periodStartDate = new Date(periodStart.getFullYear(), periodStart.getMonth(), periodStart.getDate());
  const periodEndDate = new Date(periodEnd.getFullYear(), periodEnd.getMonth(), periodEnd.getDate());
  
  // Helper function to get employee name (for unique counting)
  const getEmpName = (row) => {
    return String(row[COLUMN_INDICES.EMP_NAME] || "").trim();
  };
  
  // Opening Headcount: Employees who were active at the START of the period
  // Should NOT include employees who start on the first day of the period
  // Original formula = COUNTIFS(C:C, "<=" & periodStart, O:O, ">=" & periodStart) 
  //                    + COUNTIFS(C:C, "<=" & periodStart, O:O, "")
  // But we need: Started < periodStart (not <=) to exclude same-day starts
  // Count unique employee names (since Emp ID may be missing)
  const openingHCNames = new Set();
  filteredRows.forEach(row => {
    const empName = getEmpName(row);
    if (!empName) return; // Skip rows without employee name
    
    const startDate = parseDate(row[COLUMN_INDICES.START_DATE]);
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    
    // Must have started BEFORE period start (not on the first day)
    // This ensures Opening HC matches previous month's Closing HC
    if (!startDate || startDate >= periodStartDate) return;
    
    // Either never terminated (active) OR terminated on/after period start
    const isActive = !termDate || termDate >= periodStartDate;
    if (isActive) {
      openingHCNames.add(empName);
    }
  });
  const openingHC = openingHCNames.size;
  
  // Closing Headcount: Employees who are active at the END of the period
  // Should NOT include employees who terminate on the last day of the period
  // Original formula = COUNTIFS(C:C, "<=" & periodEnd, O:O, ">" & periodEnd)
  //                     + COUNTIFS(C:C, "<=" & periodEnd, O:O, "")
  // Employees who: Started <= periodEnd AND (Terminated > periodEnd OR never terminated)
  // Count unique employee names
  const closingHCNames = new Set();
  filteredRows.forEach(row => {
    const empName = getEmpName(row);
    if (!empName) return; // Skip rows without employee name
    
    const startDate = parseDate(row[COLUMN_INDICES.START_DATE]);
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    
    // Must have started on or before period end
    if (!startDate || startDate > periodEndDate) return;
    
    // Either never terminated (active) OR terminated AFTER period end (not on the last day)
    // If terminated on the last day, they're not active at the end of the period
    const isActive = !termDate || termDate > periodEndDate;
    if (isActive) {
      closingHCNames.add(empName);
    }
  });
  const closingHC = closingHCNames.size;
  
  // Hires: Original formula = COUNTIFS(C:C, ">=" & periodStart, C:C, "<=" & periodEnd)
  // Employees who started between periodStart and periodEnd (inclusive)
  // Count unique employee names
  const hireNames = new Set();
  filteredRows.forEach(row => {
    const empName = getEmpName(row);
    if (!empName) return; // Skip rows without employee name
    
    const startDate = parseDate(row[COLUMN_INDICES.START_DATE]);
    if (!startDate) return;
    if (startDate >= periodStartDate && startDate <= periodEndDate) {
      hireNames.add(empName);
    }
  });
  const hires = hireNames.size;
  
  // Terms: Original formula = COUNTIFS(O:O, ">=" & periodStart, O:O, "<=" & periodEnd)
  // Employees who terminated between periodStart and periodEnd (inclusive)
  // Count unique employee names
  // If terminationReason filter is specified, only count terminations with that reason
  const termNames = new Set();
  filteredRows.forEach(row => {
    const empName = getEmpName(row);
    if (!empName) return; // Skip rows without employee name
    
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    if (!termDate) return;
    if (termDate >= periodStartDate && termDate <= periodEndDate) {
      // Apply termination reason filter if specified
      if (filters.terminationReason) {
        const termReason = String(row[COLUMN_INDICES.TERMINATION_REASON] || "").trim();
        if (termReason !== filters.terminationReason) return;
      }
      termNames.add(empName);
    }
  });
  const terms = termNames.size;
  
  // Voluntary Terms: Original formula = COUNTIFS with multiple conditions
  // Employees who terminated in period AND have voluntary termination type
  // Count unique employee names
  // If terminationReason filter is specified, only count terminations with that reason
  const voluntaryTermNames = new Set();
  filteredRows.forEach(row => {
    const empName = getEmpName(row);
    if (!empName) return; // Skip rows without employee name
    
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    if (!termDate) return;
    if (termDate < periodStartDate || termDate > periodEndDate) return;
    
    // Apply termination reason filter if specified
    if (filters.terminationReason) {
      const termReason = String(row[COLUMN_INDICES.TERMINATION_REASON] || "").trim();
      if (termReason !== filters.terminationReason) return;
    }
    
    const termType = String(row[COLUMN_INDICES.LEAVE_TERMINATION_TYPE] || "").trim();
    const isVoluntary = termType === "Voluntary" || 
           termType === "Voluntary - Regrettable" || 
           termType === "Voluntary - Non regrettable";
    if (isVoluntary) {
      voluntaryTermNames.add(empName);
    }
  });
  const voluntaryTerms = voluntaryTermNames.size;
  
  // Involuntary Terms: Original formula = COUNTIFS with multiple conditions
  // Employees who terminated in period AND have involuntary termination type
  // Count unique employee names
  // If terminationReason filter is specified, only count terminations with that reason
  const involuntaryTermNames = new Set();
  filteredRows.forEach(row => {
    const empName = getEmpName(row);
    if (!empName) return; // Skip rows without employee name
    
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    if (!termDate) return;
    if (termDate < periodStartDate || termDate > periodEndDate) return;
    
    // Apply termination reason filter if specified
    if (filters.terminationReason) {
      const termReason = String(row[COLUMN_INDICES.TERMINATION_REASON] || "").trim();
      if (termReason !== filters.terminationReason) return;
    }
    
    const termType = String(row[COLUMN_INDICES.LEAVE_TERMINATION_TYPE] || "").trim();
    const isInvoluntary = termType === "Involuntary" || 
           termType === "Involuntary - Regrettable" || 
           termType === "End of Contract";
    if (isInvoluntary) {
      involuntaryTermNames.add(empName);
    }
  });
  const involuntaryTerms = involuntaryTermNames.size;
  
  // Regrettable Terms: Count terminations that are regrettable
  // Logic: Use Column U (Termination Category) if available, otherwise use Column T (Leave/Termination Type)
  // - Column U = "Regretted" → Regrettable
  // - Column U = "Unregretted" → Not Regrettable
  // - Column U is blank:
  //   - Column T = "Voluntary - Regrettable" → Regrettable
  //   - Column T = "Involuntary" → Not Regrettable (always)
  //   - Column T = "Voluntary - Non regrettable" → Not Regrettable
  // Count unique employee names
  // If terminationReason filter is specified, only count terminations with that reason
  const regrettableTermNames = new Set();
  filteredRows.forEach(row => {
    const empName = getEmpName(row);
    if (!empName) return; // Skip rows without employee name
    
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    if (!termDate) return;
    if (termDate < periodStartDate || termDate > periodEndDate) return;
    
    // Apply termination reason filter if specified
    if (filters.terminationReason) {
      const termReason = String(row[COLUMN_INDICES.TERMINATION_REASON] || "").trim();
      if (termReason !== filters.terminationReason) return;
    }
    
    const termCategory = String(row[COLUMN_INDICES.TERMINATION_CATEGORY] || "").trim();
    const termType = String(row[COLUMN_INDICES.LEAVE_TERMINATION_TYPE] || "").trim();
    
    let isRegrettable = false;
    
    if (termCategory) {
      // Use Column U value if available
      isRegrettable = termCategory === "Regretted" || termCategory.toLowerCase() === "regretted";
    } else {
      // Fallback to Column T logic if Column U is blank
      if (termType === "Voluntary - Regrettable" || termType === "Involuntary - Regrettable") {
        isRegrettable = true; // Both voluntary and involuntary regrettable are counted
      } else if (termType === "Involuntary" || 
                 termType === "Voluntary - Non regrettable" || termType === "End of Contract") {
        isRegrettable = false; // Always non-regrettable
      }
    }
    
    if (isRegrettable) {
      regrettableTermNames.add(empName);
    }
  });
  const regrettableTerms = regrettableTermNames.size;
  
  // Calculate rates (return as decimals, sheet will format as percentage)
  // Original formulas:
  // - attrition = F4 / ((B4 + C4) / 2) where F4 = Voluntary Terms
  // - retention = (B4 - E4) / B4 where E4 = Terms
  // - turnover = E4 / ((B4 + C4) / 2) where E4 = Terms
  // - regrettable turnover = Regrettable Terms / ((B4 + C4) / 2)
  const avgHC = (openingHC + closingHC) / 2;
  const attrition = avgHC > 0 ? (voluntaryTerms / avgHC) : 0;  // Attrition = Voluntary Terms / Avg HC
  const retention = openingHC > 0 ? ((openingHC - terms) / openingHC) : 0;  // Retention = (Opening HC - Terms) / Opening HC
  const turnover = avgHC > 0 ? (terms / avgHC) : 0;  // Turnover = Total Terms / Avg HC
  const regrettableTurnover = avgHC > 0 ? (regrettableTerms / avgHC) : 0;  // Regrettable Turnover = Regrettable Terms / Avg HC
  
  return {
    openingHC,
    closingHC,
    hires,
    terms,
    voluntaryTerms,
    involuntaryTerms,
    regrettableTerms,
    // Round to 4 decimal places for accurate percentage display (0.9680 = 96.80%)
    attrition: Math.round(attrition * 10000) / 10000,
    retention: Math.round(retention * 10000) / 10000,
    turnover: Math.round(turnover * 10000) / 10000,
    regrettableTurnover: Math.round(regrettableTurnover * 10000) / 10000
  };
}


/**
 * Updates filter options sheet with unique values for vetting
 * Creates/updates a "FilterOptions" sheet with unique Site, ELT, and Department values
 */
function updateFilterOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let filterSheet = ss.getSheetByName("FilterOptions");
  
  if (!filterSheet) {
    filterSheet = ss.insertSheet("FilterOptions");
  } else {
    filterSheet.clear();
  }
  
  const uniqueValues = getUniqueFilterValues();
  
  // Write section headers to clarify which filters apply to which page
  filterSheet.getRange(1, 1).setValue("Filter Options - Available Values");
  filterSheet.getRange(1, 1).setFontWeight("bold");
  filterSheet.getRange(1, 1).setFontSize(12);
  
  filterSheet.getRange(3, 1).setValue("Available Values for Reference:");
  filterSheet.getRange(3, 1).setFontWeight("bold");
  filterSheet.getRange(3, 1).setFontSize(11);
  
  // Write headers
  filterSheet.getRange(4, 1, 1, 4).setValues([["Site", "ELT", "Department", "Termination Reason"]]);
  filterSheet.getRange(4, 1, 1, 4).setFontWeight("bold");
  
  // Find max length
  const maxLength = Math.max(
    uniqueValues.sites.length,
    uniqueValues.elts.length,
    uniqueValues.departments.length,
    uniqueValues.terminationReasons.length
  );
  
  // Write data
  if (maxLength > 0) {
    const values = [];
    for (let i = 0; i < maxLength; i++) {
      values.push([
        uniqueValues.sites[i] || "",
        uniqueValues.elts[i] || "",
        uniqueValues.departments[i] || "",
        uniqueValues.terminationReasons[i] || ""
      ]);
    }
    filterSheet.getRange(5, 1, maxLength, 4).setValues(values);
  }
  
  // Auto-resize columns
  filterSheet.autoResizeColumns(1, 4);
  
  // Add instructions
  const instructionRow = maxLength + 6;
  filterSheet.getRange(instructionRow, 1).setValue("Instructions:");
  filterSheet.getRange(instructionRow, 1).setFontWeight("bold");
  filterSheet.getRange(instructionRow + 1, 1).setValue("1. Review the unique values above for reference");
  filterSheet.getRange(instructionRow + 2, 1).setValue("2. Time Period filters are configured in FilterConfig sheet");
  filterSheet.getRange(instructionRow + 3, 1).setValue("3. Time Period filter applies to: Overall Data, Headcount Metrics & ELT Metrics");
  
  SpreadsheetApp.getUi().alert(`Filter options updated. Found ${uniqueValues.sites.length} sites, ${uniqueValues.elts.length} ELTs, ${uniqueValues.departments.length} departments, ${uniqueValues.terminationReasons.length} termination reasons.`);
}

/**
 * Generates overall HR metrics for multiple periods and writes to "Overall Data" sheet
 * Creates a comprehensive metrics table with all calculated values
 */
function generateOverallData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let metricsSheet = ss.getSheetByName("Overall Data");
  const isNewSheet = !metricsSheet;
  
  if (!metricsSheet) {
    metricsSheet = ss.insertSheet("Overall Data");
  }
  // Don't clear content - preserve user formatting
  // We'll just update the values directly
  
  // Overall Data shows entire company data - only time period filters apply
  let filters = {}; // No site/ELT/department filters for Overall Data
  let timePeriodFilter = null;
  let aggregationType = null; // 'halves' or 'quarters' when Year + Halves/Quarters selected
  const filterConfigSheet = ss.getSheetByName("FilterConfig");
  if (filterConfigSheet) {
    // Get time period filters only (for Overall Data)
    const yearFilter = filterConfigSheet.getRange(5, 2).getValue(); // Year filter
    const halvesFilter = filterConfigSheet.getRange(6, 2).getValue(); // Halves filter
    const quarterFilter = filterConfigSheet.getRange(7, 2).getValue(); // Quarters filter
    
    if (yearFilter) {
      const year = parseInt(yearFilter, 10);
      if (halvesFilter === "Halves") {
        // Year + Halves: aggregate by halves (show H1 and H2)
        timePeriodFilter = { type: 'year', value: year };
        aggregationType = { type: 'halves' }; // No specific value = show both
      } else if (quarterFilter === "Quarters") {
        // Year + Quarters: aggregate by quarters (show Q1-Q4)
        timePeriodFilter = { type: 'year', value: year };
        aggregationType = { type: 'quarters' }; // No specific value = show all
      } else {
        // Year only: show monthly
        timePeriodFilter = { type: 'year', value: year };
      }
    }
  }
  
  // Generate periods (monthly from Jan 2024 to current month)
  let allPeriods = [];
  const startDate = new Date(2024, 0, 1); // Jan 2024
  const endDate = new Date();
  
  let currentDate = new Date(startDate);
  while (currentDate <= endDate) {
    const periodStart = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
    // Using day 0 gives the last day of the previous month, which automatically handles leap years
    // e.g., new Date(2024, 2, 0) = Feb 29, 2024 (leap year)
    //      new Date(2025, 2, 0) = Feb 28, 2025 (non-leap year)
    const periodEnd = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 0);
    
    // Don't include future months
    if (periodStart <= endDate) {
      allPeriods.push({
        start: periodStart,
        end: periodEnd,
        label: Utilities.formatDate(periodStart, Session.getScriptTimeZone(), "MMM yyyy"),
        year: periodStart.getFullYear(),
        quarter: Math.floor(periodStart.getMonth() / 3) + 1,
        half: periodStart.getMonth() < 6 ? 1 : 2 // H1 = Jan-Jun (months 0-5), H2 = Jul-Dec (months 6-11)
      });
    }
    
    currentDate = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 1);
  }
  
  // Apply time period filter if specified
  let periods = allPeriods;
  Logger.log(`Generated ${allPeriods.length} total periods (Jan 2024 to current)`);
  
  if (timePeriodFilter) {
    const beforeFilter = periods.length;
    if (timePeriodFilter.type === 'year') {
      periods = allPeriods.filter(p => p.year === timePeriodFilter.value);
      Logger.log(`After Year filter (${timePeriodFilter.value}): ${periods.length} periods (from ${beforeFilter})`);
    } else if (timePeriodFilter.type === 'halves') {
      // For halves, filter by half number (works across years)
      periods = allPeriods.filter(p => p.half === timePeriodFilter.value);
      Logger.log(`After Halves filter (H${timePeriodFilter.value}): ${periods.length} periods (from ${beforeFilter})`);
    } else if (timePeriodFilter.type === 'quarter') {
      // For quarter, filter by quarter number (works across years)
      periods = allPeriods.filter(p => p.quarter === timePeriodFilter.value);
      Logger.log(`After Quarter filter (Q${timePeriodFilter.value}): ${periods.length} periods (from ${beforeFilter})`);
    }
  } else {
    Logger.log(`No time period filter applied, using all ${periods.length} periods`);
  }
  
  if (periods.length === 0) {
    Logger.log(`WARNING: No periods match the time period filter!`);
  }
  
  // If aggregation is requested (Year + Halves/Quarters), aggregate periods
  let aggregatedPeriods = [];
  if (aggregationType && timePeriodFilter && timePeriodFilter.type === 'year') {
    if (aggregationType.type === 'halves') {
      // Aggregate by halves: show both H1 and H2
      const h1Periods = periods.filter(p => p.half === 1);
      const h2Periods = periods.filter(p => p.half === 2);
      
      if (h1Periods.length > 0) {
        const h1Start = h1Periods[0].start;
        const h1End = h1Periods[h1Periods.length - 1].end;
        aggregatedPeriods.push({
          start: h1Start,
          end: h1End,
          label: `H1 ${timePeriodFilter.value}`,
          periods: h1Periods
        });
      }
      
      if (h2Periods.length > 0) {
        const h2Start = h2Periods[0].start;
        const h2End = h2Periods[h2Periods.length - 1].end;
        aggregatedPeriods.push({
          start: h2Start,
          end: h2End,
          label: `H2 ${timePeriodFilter.value}`,
          periods: h2Periods
        });
      }
    } else if (aggregationType.type === 'quarters') {
      // Aggregate by quarters: show all Q1-Q4
      for (let q = 1; q <= 4; q++) {
        const qPeriods = periods.filter(p => p.quarter === q);
        if (qPeriods.length > 0) {
          const qStart = qPeriods[0].start;
          const qEnd = qPeriods[qPeriods.length - 1].end;
          aggregatedPeriods.push({
            start: qStart,
            end: qEnd,
            label: `Q${q} ${timePeriodFilter.value}`,
            periods: qPeriods
          });
        }
      }
    }
    
    if (aggregatedPeriods.length > 0) {
      periods = aggregatedPeriods;
      Logger.log(`Aggregated into ${periods.length} periods`);
    }
  }
  
  // Write headers (only set formatting if it's a new sheet)
  const headers = [
    "Period",
    "Opening Headcount",
    "Closing Headcount",
    "Hires",
    "Terms",
    "Voluntary",
    "Involuntary",
    "Regrettable",
    "Attrition %",
    "Retention %",
    "Turnover %",
    "Regrettable Turnover %"
  ];
  
  metricsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  // Only set header formatting if it's a new sheet (preserve user formatting on existing sheets)
  if (isNewSheet) {
  metricsSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }
  
  // Calculate and write metrics for each period
  const values = [];
  for (let i = 0; i < periods.length; i++) {
    try {
      let metrics;
      if (periods[i].periods) {
        // Aggregated period: calculate metrics across all months in the period
        // Opening HC = opening HC of first month
        // Closing HC = closing HC of last month
        // Hires/Terms = sum across all months
        const firstPeriod = periods[i].periods[0];
        const lastPeriod = periods[i].periods[periods[i].periods.length - 1];
        
        const firstMetrics = calculateHRMetrics(firstPeriod.start, firstPeriod.end, filters);
        const lastMetrics = calculateHRMetrics(lastPeriod.start, lastPeriod.end, filters);
        
        // Sum up hires and terms across all months
        let totalHires = 0;
        let totalTerms = 0;
        let totalVoluntary = 0;
        let totalInvoluntary = 0;
        let totalRegrettable = 0;
        
        periods[i].periods.forEach(period => {
          const periodMetrics = calculateHRMetrics(period.start, period.end, filters);
          totalHires += periodMetrics.hires;
          totalTerms += periodMetrics.terms;
          totalVoluntary += periodMetrics.voluntaryTerms;
          totalInvoluntary += periodMetrics.involuntaryTerms;
          totalRegrettable += periodMetrics.regrettableTerms;
        });
        
        // Opening HC = first month's opening, Closing HC = last month's closing
        const openingHC = firstMetrics.openingHC;
        const closingHC = lastMetrics.closingHC;
        const avgHC = (openingHC + closingHC) / 2;
        
        // Calculate rates
        const attrition = avgHC > 0 ? (totalVoluntary / avgHC) : 0;
        const retention = openingHC > 0 ? ((openingHC - totalTerms) / openingHC) : 0;
        const turnover = avgHC > 0 ? (totalTerms / avgHC) : 0;
        const regrettableTurnover = avgHC > 0 ? (totalRegrettable / avgHC) : 0;
        
        metrics = {
          openingHC: openingHC,
          closingHC: closingHC,
          hires: totalHires,
          terms: totalTerms,
          voluntaryTerms: totalVoluntary,
          involuntaryTerms: totalInvoluntary,
          regrettableTerms: totalRegrettable,
          attrition: Math.round(attrition * 10000) / 10000,
          retention: Math.round(retention * 10000) / 10000,
          turnover: Math.round(turnover * 10000) / 10000,
          regrettableTurnover: Math.round(regrettableTurnover * 10000) / 10000
        };
      } else {
        // Regular period: calculate normally
        metrics = calculateHRMetrics(periods[i].start, periods[i].end, filters);
      }
      
      values.push([
        periods[i].label,
        metrics.openingHC,
        metrics.closingHC,
        metrics.hires,
        metrics.terms,
        metrics.voluntaryTerms,
        metrics.involuntaryTerms,
        metrics.regrettableTerms,
        metrics.attrition,
        metrics.retention,
        metrics.turnover,
        metrics.regrettableTurnover
      ]);
    } catch (error) {
      Logger.log(`Error calculating metrics for ${periods[i].label}: ${error.message}`);
      values.push([
        periods[i].label,
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        ""
      ]);
    }
  }
  
  if (values.length > 0) {
    // Update data values (preserves existing formatting for existing rows)
    metricsSheet.getRange(2, 1, values.length, headers.length).setValues(values);
    
    // Clear any extra rows if data shrunk
    const existingLastRow = metricsSheet.getLastRow();
    if (existingLastRow > values.length + 1) {
      const rowsToClear = existingLastRow - (values.length + 1);
      metricsSheet.getRange(values.length + 2, 1, rowsToClear, headers.length).clearContent();
    }
  }
  
  // Only set percentage format if it's a new sheet (preserve user formatting on existing sheets)
  // Percentage columns are at columns 9, 10, 11, 12 (Attrition %, Retention %, Turnover %, Regrettable Turnover %)
  const lastRow = values.length + 1;
  if (isNewSheet) {
    metricsSheet.getRange(2, 9, lastRow - 1, 4).setNumberFormat("0.0%");
  }
  
  // Only auto-resize columns if it's a new sheet (preserve user column widths)
  if (isNewSheet) {
  metricsSheet.autoResizeColumns(1, headers.length);
  }
  
  // Add filter info in columns O/P, rows 1 & 2 if filters are applied
  // Column O (15) = Filter labels, Column P (16) = Filter values
  const filterLabels = [];
  const filterValues = [];
  
  // Only show time period filter (Overall Data shows entire company)
  if (timePeriodFilter) {
    if (timePeriodFilter.type === 'year') {
      filterLabels.push("Year:");
      filterValues.push(timePeriodFilter.value);
    } else if (timePeriodFilter.type === 'halves') {
      filterLabels.push("Halves:");
      filterValues.push(`H${timePeriodFilter.value}`);
    } else if (timePeriodFilter.type === 'quarter') {
      filterLabels.push("Quarter:");
      filterValues.push(`Q${timePeriodFilter.value}`);
    }
  }
  
  // Clear old filter info (clear columns O/P, rows 1-10 to remove any old data)
  metricsSheet.getRange(1, 15, 10, 2).clearContent();
  
  // Write filter info if any filters are applied
  if (filterLabels.length > 0) {
    // Row 1: Headers
    metricsSheet.getRange(1, 15).setValue("Filters Applied:");
    metricsSheet.getRange(1, 16).setValue("Value");
    
    // Row 2 onwards: Filter labels and values
    const filterRows = [];
    for (let i = 0; i < filterLabels.length; i++) {
      filterRows.push([filterLabels[i], filterValues[i]]);
    }
    metricsSheet.getRange(2, 15, filterRows.length, 2).setValues(filterRows);
  }
  
  // Clear any old filter info that might be below the data (in column A)
  const existingLastRow = metricsSheet.getLastRow();
  if (existingLastRow > values.length + 1) {
    // Check if there's old filter info below data
    const checkRow = values.length + 2;
    const checkValue = metricsSheet.getRange(checkRow, 1).getValue();
    if (checkValue && String(checkValue).includes("Filters Applied")) {
      // Clear old filter info rows
      metricsSheet.getRange(checkRow, 1, 5, 1).clearContent();
    }
  }
  
  // Add formula descriptions in Column T, Row 1 only if they don't already exist
  // Check if "Formula Descriptions:" already exists in Column T, Row 1
  const descStartRow = 1;
  const descCellValue = metricsSheet.getRange(descStartRow, 20).getValue();
  const descriptionsExist = descCellValue && String(descCellValue).trim() === "Formula Descriptions:";
  
  if (!descriptionsExist) {
    const descriptions = [
      ["Formula Descriptions:"],
      [""],
      ["Attrition %:"],
      ["Voluntary Terms / Average Headcount"],
      ["Where Average Headcount = (Opening HC + Closing HC) / 2"],
      [""],
      ["Retention %:"],
      ["(Opening Headcount - Total Terms) / Opening Headcount"],
      [""],
      ["Turnover %:"],
      ["Total Terms / Average Headcount"],
      ["Where Average Headcount = (Opening HC + Closing HC) / 2"],
      [""],
      ["Regrettable Turnover %:"],
      ["Regrettable Terms / Average Headcount"],
      ["Where Average Headcount = (Opening HC + Closing HC) / 2"]
    ];
    
    // Column T is index 20 (1-based), start from Row 1
    metricsSheet.getRange(descStartRow, 20, descriptions.length, 1).setValues(descriptions);
    
    // Format the description section
    if (isNewSheet) {
      const descRange = metricsSheet.getRange(descStartRow, 20, 1, 1);
      descRange.setFontWeight("bold");
      // Make column T wider for readability
      metricsSheet.setColumnWidth(20, 400);
    }
  }
  
  SpreadsheetApp.getUi().alert(`Overall Data generated for ${periods.length} periods.`);
}

/**
 * Creates a FilterConfig sheet for easy filter selection
 * Users can enter filter values in this sheet, then run generateOverallData()
 */
function createFilterConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName("FilterConfig");
  
  if (!configSheet) {
    configSheet = ss.insertSheet("FilterConfig");
  } else {
    configSheet.clear();
  }
  
  // Get unique values for data validation
  const uniqueValues = getUniqueFilterValues();
  
  // Write section headers
  configSheet.getRange(1, 1).setValue("Filter Configuration");
  configSheet.getRange(1, 1).setFontWeight("bold");
  configSheet.getRange(1, 1).setFontSize(12);
  
  // Time Period Filters for Overall Data, Headcount Metrics, and ELT Metrics
  configSheet.getRange(3, 1).setValue("Time Period Filter (for 'Overall Data', 'Headcount Metrics', and 'ELT Metrics'):");
  configSheet.getRange(3, 1).setFontWeight("bold");
  configSheet.getRange(3, 1).setFontSize(11);
  
  // Write headers and labels
  configSheet.getRange(4, 1, 1, 2).setValues([["Filter", "Value"]]);
  configSheet.getRange(4, 1, 1, 2).setFontWeight("bold");
  
  configSheet.getRange(5, 1).setValue("Year:");
  configSheet.getRange(6, 1).setValue("Halves:");
  configSheet.getRange(7, 1).setValue("Quarters:");
  
  // Year dropdown (2024, 2025, 2026, etc.)
  const currentYear = new Date().getFullYear();
  const years = [];
  // Include current year and next year (since we're in November, include 2026)
  for (let y = 2024; y <= Math.max(currentYear, 2026); y++) {
    years.push(String(y));
  }
  const yearRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([""].concat(years), true)
    .build();
  configSheet.getRange(5, 2).setDataValidation(yearRule);
  configSheet.getRange(5, 3).setValue("(e.g., 2024 or 2025)");
  configSheet.getRange(5, 3).setFontStyle("italic");
  configSheet.getRange(5, 3).setFontColor("#666666");
  
  // Halves dropdown - works with Year filter to aggregate by halves (H1 and H2)
  const halvesRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["", "Halves"], true)
    .build();
  configSheet.getRange(6, 2).setDataValidation(halvesRule);
  configSheet.getRange(6, 3).setValue("(Select Year + Halves to show H1 & H2 aggregated)");
  configSheet.getRange(6, 3).setFontStyle("italic");
  configSheet.getRange(6, 3).setFontColor("#666666");
  
  // Quarters dropdown - works with Year filter to aggregate by quarters (Q1-Q4)
  const quarterRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["", "Quarters"], true)
    .build();
  configSheet.getRange(7, 2).setDataValidation(quarterRule);
  configSheet.getRange(7, 3).setValue("(Select Year + Quarters to show Q1-Q4 aggregated)");
  configSheet.getRange(7, 3).setFontStyle("italic");
  configSheet.getRange(7, 3).setFontColor("#666666");
  
  // Add instructions
  configSheet.getRange(9, 1).setValue("Instructions:");
  configSheet.getRange(9, 1).setFontWeight("bold");
  configSheet.getRange(10, 1).setValue("1. Select filter values from dropdowns above (leave blank for all)");
  configSheet.getRange(11, 1).setValue("2. Time Period filter applies to: Overall Data, Headcount Metrics & ELT Metrics");
  configSheet.getRange(12, 1).setValue("3. For aggregated periods: Select Year + Halves (H1/H2) OR Year + Quarters (Q1-Q4)");
  configSheet.getRange(13, 1).setValue("4. Halves/Quarters aggregate all months in that period into one column");
  configSheet.getRange(14, 1).setValue("5. Run 'Generate Overall Data', 'Generate Headcount Metrics', or 'Generate ELT Metrics' from menu");
  
  // Add separate filter section for Termination Reasons
  configSheet.getRange(16, 1).setValue("Filter for 'Termination Reasons' (separate from above):");
  configSheet.getRange(16, 1).setFontWeight("bold");
  configSheet.getRange(16, 1).setFontSize(11);
  
  configSheet.getRange(17, 1).setValue("Year:");
  configSheet.getRange(18, 1).setValue("Period Type:");
  
  // Year dropdown for Termination Reasons
  const termYearRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([""].concat(years), true)
    .build();
  configSheet.getRange(17, 2).setDataValidation(termYearRule);
  configSheet.getRange(17, 3).setValue("(Select year for termination reasons)");
  configSheet.getRange(17, 3).setFontStyle("italic");
  configSheet.getRange(17, 3).setFontColor("#666666");
  
  // Period Type dropdown: Halves or Quarters
  const periodTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["", "Halves", "Quarters"], true)
    .build();
  configSheet.getRange(18, 2).setDataValidation(periodTypeRule);
  configSheet.getRange(18, 3).setValue("(Select Halves or Quarters to toggle)");
  configSheet.getRange(18, 3).setFontStyle("italic");
  configSheet.getRange(18, 3).setFontColor("#666666");
  
  configSheet.getRange(20, 1).setValue("6. For Termination Reasons: Select Year + Period Type (Halves or Quarters)");
  configSheet.getRange(21, 1).setValue("7. Period Type toggles between H1/H2 (Halves) or Q1-Q4 (Quarters)");
  
  configSheet.autoResizeColumns(1, 3);
  
  SpreadsheetApp.getUi().alert("FilterConfig sheet created. Select your filters and run the appropriate generate function.");
}

/**
 * Gets unique values for filtering (Site, ELT, Department, Termination Reasons)
 * @returns {Object} Object with arrays of unique values
 */
function getUniqueFilterValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawDataSheet = ss.getSheetByName("RawData");
  
  if (!rawDataSheet) {
    return { sites: [], elts: [], departments: [], terminationReasons: [] };
  }
  
  const data = rawDataSheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { sites: [], elts: [], departments: [], terminationReasons: [] };
  }
  
  const rows = data.slice(1);
  const sites = new Set();
  const elts = new Set();
  const departments = new Set();
  const terminationReasons = new Set();
  
  rows.forEach(row => {
    const site = String(row[COLUMN_INDICES.SITE] || "").trim();
    const elt = String(row[COLUMN_INDICES.ELT] || "").trim();
    const dept = String(row[COLUMN_INDICES.DEPARTMENT] || "").trim();
    const termReason = String(row[COLUMN_INDICES.TERMINATION_REASON] || "").trim();
    
    if (site) sites.add(site);
    if (elt) elts.add(elt);
    if (dept) departments.add(dept);
    if (termReason) terminationReasons.add(termReason);
  });
  
  return {
    sites: Array.from(sites).sort(),
    elts: Array.from(elts).sort(),
    departments: Array.from(departments).sort(),
    terminationReasons: Array.from(terminationReasons).sort()
  };
}

/**
 * Generates site-wise headcount metrics month-on-month
 * Creates a table with Sites and Metrics as rows and Months/Years as columns
 * Metrics include: Average HC, Hires, Terminations, Regrettable Terms, Non-Regrettable Terms, Attrition %, Regrettable %
 */
function generateHeadcountMetrics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let headcountSheet = ss.getSheetByName("Headcount Metrics");
  const isNewSheet = !headcountSheet;
  
  if (!headcountSheet) {
    headcountSheet = ss.insertSheet("Headcount Metrics");
  }
  
  const rawDataSheet = ss.getSheetByName("RawData");
  if (!rawDataSheet) {
    SpreadsheetApp.getUi().alert("RawData sheet not found. Please run 'Fetch Bob Report' first.");
    return;
  }
  
  // Get unique sites
  const uniqueValues = getUniqueFilterValues();
  const sites = uniqueValues.sites;
  
  if (sites.length === 0) {
    SpreadsheetApp.getUi().alert("No sites found in data.");
    return;
  }
  
  // Generate periods (monthly from Jan 2024 to current month)
  // Check for time period filters in FilterConfig
  let allPeriods = [];
  const startDate = new Date(2024, 0, 1); // Jan 2024
  const endDate = new Date();
  
  let currentDate = new Date(startDate);
  while (currentDate <= endDate) {
    const periodStart = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
    // Using day 0 gives the last day of the previous month, which automatically handles leap years
    // e.g., new Date(2024, 2, 0) = Feb 29, 2024 (leap year)
    //      new Date(2025, 2, 0) = Feb 28, 2025 (non-leap year)
    const periodEnd = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 0);
    
    if (periodStart <= endDate) {
      allPeriods.push({
        start: periodStart,
        end: periodEnd,
        label: Utilities.formatDate(periodStart, Session.getScriptTimeZone(), "MMM yyyy"),
        year: periodStart.getFullYear(),
        quarter: Math.floor(periodStart.getMonth() / 3) + 1,
        half: periodStart.getMonth() < 6 ? 1 : 2 // H1 = Jan-Jun (months 0-5), H2 = Jul-Dec (months 6-11)
      });
    }
    
    currentDate = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 1);
  }
  
  // Apply time period filters from FilterConfig
  let periods = allPeriods;
  let aggregationType = null;
  const filterConfigSheet = ss.getSheetByName("FilterConfig");
  if (filterConfigSheet) {
    const yearFilter = filterConfigSheet.getRange(5, 2).getValue(); // Year filter
    const halvesFilter = filterConfigSheet.getRange(6, 2).getValue(); // Halves filter
    const quarterFilter = filterConfigSheet.getRange(7, 2).getValue(); // Quarters filter
    
    if (yearFilter) {
      const year = parseInt(yearFilter, 10);
      if (halvesFilter === "Halves") {
        // Year + Halves: aggregate by halves
        periods = periods.filter(p => p.year === year);
        aggregationType = { type: 'halves' };
      } else if (quarterFilter === "Quarters") {
        // Year + Quarters: aggregate by quarters
        periods = periods.filter(p => p.year === year);
        aggregationType = { type: 'quarters' };
      } else {
        // Year only: show monthly
        periods = periods.filter(p => p.year === year);
      }
    }
  }
  
  // Aggregate periods if needed
  let yearForLabel = null;
  if (filterConfigSheet) {
    const yearFilter = filterConfigSheet.getRange(5, 2).getValue();
    if (yearFilter) yearForLabel = parseInt(yearFilter, 10);
  }
  
  if (aggregationType) {
    let aggregatedPeriods = [];
    if (aggregationType.type === 'halves') {
      const h1Periods = periods.filter(p => p.half === 1);
      const h2Periods = periods.filter(p => p.half === 2);
      
      if (h1Periods.length > 0) {
        aggregatedPeriods.push({
          start: h1Periods[0].start,
          end: h1Periods[h1Periods.length - 1].end,
          label: `H1 ${yearForLabel || ''}`,
          periods: h1Periods
        });
      }
      
      if (h2Periods.length > 0) {
        aggregatedPeriods.push({
          start: h2Periods[0].start,
          end: h2Periods[h2Periods.length - 1].end,
          label: `H2 ${yearForLabel || ''}`,
          periods: h2Periods
        });
      }
    } else if (aggregationType.type === 'quarters') {
      for (let q = 1; q <= 4; q++) {
        const qPeriods = periods.filter(p => p.quarter === q);
        if (qPeriods.length > 0) {
          aggregatedPeriods.push({
            start: qPeriods[0].start,
            end: qPeriods[qPeriods.length - 1].end,
            label: `Q${q} ${yearForLabel || ''}`,
            periods: qPeriods
          });
        }
      }
    }
    
    if (aggregatedPeriods.length > 0) {
      periods = aggregatedPeriods;
    }
  }
  
  // Headcount Metrics is structured by site, so no additional filters needed
  let additionalFilters = {};
  
  // Build table: Site/Metric as rows, Periods as columns
  const headerRow = ["Site / Metric"].concat(periods.map(p => p.label));
  
  // Define metrics to calculate for each site
  const metricLabels = [
    "Opening HC",
    "Closing HC",
    "Average Headcount",
    "Hires",
    "Terminations",
    "Regrettable Terms",
    "Non-Regrettable Terms",
    "Retention %",
    "Turnover %",
    "Regrettable %"
  ];
  
  // OPTIMIZATION: Cache metrics calculations to avoid redundant processing
  // Key format: "site|periodLabel" -> metrics object
  const metricsCache = {};
  
  // Pre-calculate all metrics for each site/period combination once
  sites.forEach(site => {
    periods.forEach(period => {
      const cacheKey = `${site}|${period.label}`;
      if (!metricsCache[cacheKey]) {
        try {
          // Headcount Metrics is structured by site, so only filter by site
          const periodFilters = { site: site };
          
          // If aggregated period, calculate aggregated metrics
          if (period.periods) {
            const firstPeriod = period.periods[0];
            const lastPeriod = period.periods[period.periods.length - 1];
            
            const firstMetrics = calculateHRMetrics(firstPeriod.start, firstPeriod.end, periodFilters);
            const lastMetrics = calculateHRMetrics(lastPeriod.start, lastPeriod.end, periodFilters);
            
            let totalHires = 0;
            let totalTerms = 0;
            let totalVoluntary = 0;
            let totalInvoluntary = 0;
            let totalRegrettable = 0;
            
            period.periods.forEach(p => {
              const pMetrics = calculateHRMetrics(p.start, p.end, periodFilters);
              totalHires += pMetrics.hires;
              totalTerms += pMetrics.terms;
              totalVoluntary += pMetrics.voluntaryTerms;
              totalInvoluntary += pMetrics.involuntaryTerms;
              totalRegrettable += pMetrics.regrettableTerms;
            });
            
            const openingHC = firstMetrics.openingHC;
            const closingHC = lastMetrics.closingHC;
            const avgHC = (openingHC + closingHC) / 2;
            
            const attrition = avgHC > 0 ? (totalVoluntary / avgHC) : 0;
            const retention = openingHC > 0 ? ((openingHC - totalTerms) / openingHC) : 0;
            const turnover = avgHC > 0 ? (totalTerms / avgHC) : 0;
            const regrettableTurnover = avgHC > 0 ? (totalRegrettable / avgHC) : 0;
            
            metricsCache[cacheKey] = {
              openingHC: openingHC,
              closingHC: closingHC,
              hires: totalHires,
              terms: totalTerms,
              voluntaryTerms: totalVoluntary,
              involuntaryTerms: totalInvoluntary,
              regrettableTerms: totalRegrettable,
              attrition: Math.round(attrition * 10000) / 10000,
              retention: Math.round(retention * 10000) / 10000,
              turnover: Math.round(turnover * 10000) / 10000,
              regrettableTurnover: Math.round(regrettableTurnover * 10000) / 10000
            };
          } else {
            metricsCache[cacheKey] = calculateHRMetrics(period.start, period.end, periodFilters);
          }
        } catch (error) {
          Logger.log(`Error calculating metrics for ${site} in ${period.label}: ${error.message}`);
          metricsCache[cacheKey] = null;
        }
      }
    });
  });
  
  // Build data rows using cached metrics
  const dataRows = [];
  sites.forEach((site, siteIndex) => {
    // Add header row at the start of each site section
    dataRows.push(headerRow);
    
    // For each site, create rows for each metric
    metricLabels.forEach((metricLabel, metricIndex) => {
      const row = [`${site} - ${metricLabel}`];
      
      periods.forEach(period => {
        const cacheKey = `${site}|${period.label}`;
        const metrics = metricsCache[cacheKey];
        
        if (!metrics) {
          // Use blank for counts, 0 for percentages
          row.push(metricIndex < 7 ? "" : 0);
          return;
        }
        
        let value = 0;
        switch(metricIndex) {
          case 0: // Opening HC
            value = metrics.openingHC;
            break;
          case 1: // Closing HC
            value = metrics.closingHC;
            break;
          case 2: // Average Headcount - round to nearest integer
            value = Math.round((metrics.openingHC + metrics.closingHC) / 2);
            break;
          case 3: // Hires
            value = metrics.hires;
            break;
          case 4: // Terminations
            value = metrics.terms;
            break;
          case 5: // Regrettable Terms
            value = metrics.regrettableTerms;
            break;
          case 6: // Non-Regrettable Terms
            value = metrics.terms - metrics.regrettableTerms;
            break;
          case 7: // Retention %
            // Use the pre-calculated retention from calculateHRMetrics
            value = metrics.retention;
            break;
          case 8: // Turnover %
            // Use the pre-calculated turnover from calculateHRMetrics (Total Terms / Avg HC)
            value = metrics.turnover;
            break;
          case 9: // Regrettable %
            // Use the pre-calculated regrettableTurnover from calculateHRMetrics (Regrettable Terms / Avg HC)
            value = metrics.regrettableTurnover;
            break;
        }
        // For count metrics (0-6), use blank instead of 0. For percentages (7-9), keep 0
        if (metricIndex < 7 && value === 0) {
          row.push(""); // Blank for zero counts
        } else {
          row.push(value);
        }
      });
      dataRows.push(row);
    });
    
    // Add 1 blank row after each site (except the last one)
    if (siteIndex < sites.length - 1) {
      const blankRow = [""].concat(new Array(periods.length).fill(""));
      dataRows.push(blankRow);
    }
  });
  
  // Write data in one batch operation - only values, preserve all user formatting
  if (dataRows.length > 0) {
    try {
      headcountSheet.getRange(1, 1, dataRows.length, headerRow.length).setValues(dataRows);
      
      // Clear any extra rows if data shrunk
      const existingLastRow = headcountSheet.getLastRow();
      if (existingLastRow > dataRows.length) {
        const rowsToClear = existingLastRow - dataRows.length;
        headcountSheet.getRange(dataRows.length + 1, 1, rowsToClear, headerRow.length).clearContent();
      }
    } catch (error) {
      Logger.log(`Error writing headcount metrics: ${error.message}`);
      SpreadsheetApp.getUi().alert(`Error generating headcount metrics: ${error.message}`);
      return;
    }
  }
  
  SpreadsheetApp.getUi().alert(`Headcount Metrics generated for ${sites.length} sites with ${metricLabels.length} metrics each, across ${periods.length} periods.`);
}

/**
 * Generates ELT-wise headcount metrics month-on-month
 * Creates a table with ELTs and Metrics as rows and Months/Years as columns
 * Metrics include: Opening HC, Closing HC, Average HC, Hires, Terminations, Regrettable Terms, Non-Regrettable Terms, Retention %, Turnover %, Regrettable %
 * Can filter by Termination Reason
 */
function generateELTMetrics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let eltSheet = ss.getSheetByName("ELT Metrics");
  const isNewSheet = !eltSheet;
  
  if (!eltSheet) {
    eltSheet = ss.insertSheet("ELT Metrics");
  }
  
  const rawDataSheet = ss.getSheetByName("RawData");
  if (!rawDataSheet) {
    SpreadsheetApp.getUi().alert("RawData sheet not found. Please run 'Fetch Bob Report' first.");
    return;
  }
  
  // Get unique ELTs
  const uniqueValues = getUniqueFilterValues();
  const elts = uniqueValues.elts;
  
  if (elts.length === 0) {
    SpreadsheetApp.getUi().alert("No ELTs found in data.");
    return;
  }
  
  // ELT Metrics shows all ELTs - no filters needed (structured by ELT)
  // Time period filters are applied below
  
  // Generate periods (monthly from Jan 2024 to current month)
  // Check for time period filters in FilterConfig
  let allPeriods = [];
  const startDate = new Date(2024, 0, 1); // Jan 2024
  const endDate = new Date();
  
  let currentDate = new Date(startDate);
  while (currentDate <= endDate) {
    const periodStart = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
    const periodEnd = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 0);
    
    if (periodStart <= endDate) {
      allPeriods.push({
        start: periodStart,
        end: periodEnd,
        label: Utilities.formatDate(periodStart, Session.getScriptTimeZone(), "MMM yyyy"),
        year: periodStart.getFullYear(),
        quarter: Math.floor(periodStart.getMonth() / 3) + 1,
        half: periodStart.getMonth() < 6 ? 1 : 2
      });
    }
    
    currentDate = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 1);
  }
  
  // Apply time period filters from FilterConfig
  let periods = allPeriods;
  let aggregationType = null;
  const filterConfigSheet = ss.getSheetByName("FilterConfig");
  if (filterConfigSheet) {
    const yearFilter = filterConfigSheet.getRange(5, 2).getValue(); // Year filter
    const halvesFilter = filterConfigSheet.getRange(6, 2).getValue(); // Halves filter
    const quarterFilter = filterConfigSheet.getRange(7, 2).getValue(); // Quarters filter
    
    if (yearFilter) {
      const year = parseInt(yearFilter, 10);
      if (halvesFilter === "Halves") {
        // Year + Halves: aggregate by halves
        periods = periods.filter(p => p.year === year);
        aggregationType = { type: 'halves' };
      } else if (quarterFilter === "Quarters") {
        // Year + Quarters: aggregate by quarters
        periods = periods.filter(p => p.year === year);
        aggregationType = { type: 'quarters' };
      } else {
        // Year only: show monthly
        periods = periods.filter(p => p.year === year);
      }
    }
  }
  
  // Aggregate periods if needed
  let yearForLabel = null;
  if (filterConfigSheet) {
    const yearFilter = filterConfigSheet.getRange(5, 2).getValue();
    if (yearFilter) yearForLabel = parseInt(yearFilter, 10);
  }
  
  if (aggregationType) {
    let aggregatedPeriods = [];
    if (aggregationType.type === 'halves') {
      const h1Periods = periods.filter(p => p.half === 1);
      const h2Periods = periods.filter(p => p.half === 2);
      
      if (h1Periods.length > 0) {
        aggregatedPeriods.push({
          start: h1Periods[0].start,
          end: h1Periods[h1Periods.length - 1].end,
          label: `H1 ${yearForLabel || ''}`,
          periods: h1Periods
        });
      }
      
      if (h2Periods.length > 0) {
        aggregatedPeriods.push({
          start: h2Periods[0].start,
          end: h2Periods[h2Periods.length - 1].end,
          label: `H2 ${yearForLabel || ''}`,
          periods: h2Periods
        });
      }
    } else if (aggregationType.type === 'quarters') {
      for (let q = 1; q <= 4; q++) {
        const qPeriods = periods.filter(p => p.quarter === q);
        if (qPeriods.length > 0) {
          aggregatedPeriods.push({
            start: qPeriods[0].start,
            end: qPeriods[qPeriods.length - 1].end,
            label: `Q${q} ${yearForLabel || ''}`,
            periods: qPeriods
          });
        }
      }
    }
    
    if (aggregatedPeriods.length > 0) {
      periods = aggregatedPeriods;
    }
  }
  
  // Build table: ELT/Metric as rows, Periods as columns
  const headerRow = ["ELT / Metric"].concat(periods.map(p => p.label));
  
  // Define metrics to calculate for each ELT
  const metricLabels = [
    "Opening HC",
    "Closing HC",
    "Average Headcount",
    "Hires",
    "Terminations",
    "Regrettable Terms",
    "Non-Regrettable Terms",
    "Retention %",
    "Turnover %",
    "Regrettable %"
  ];
  
  // OPTIMIZATION: Cache metrics calculations to avoid redundant processing
  const metricsCache = {};
  
  // Pre-calculate all metrics for each ELT/period combination once
  elts.forEach(elt => {
    periods.forEach(period => {
      const cacheKey = `${elt}|${period.label}`;
      if (!metricsCache[cacheKey]) {
        try {
          // ELT Metrics is structured by ELT
          const periodFilters = { elt: elt };
          
          // If aggregated period, calculate aggregated metrics
          if (period.periods) {
            const firstPeriod = period.periods[0];
            const lastPeriod = period.periods[period.periods.length - 1];
            
            const firstMetrics = calculateHRMetrics(firstPeriod.start, firstPeriod.end, periodFilters);
            const lastMetrics = calculateHRMetrics(lastPeriod.start, lastPeriod.end, periodFilters);
            
            let totalHires = 0;
            let totalTerms = 0;
            let totalVoluntary = 0;
            let totalInvoluntary = 0;
            let totalRegrettable = 0;
            
            period.periods.forEach(p => {
              const pMetrics = calculateHRMetrics(p.start, p.end, periodFilters);
              totalHires += pMetrics.hires;
              totalTerms += pMetrics.terms;
              totalVoluntary += pMetrics.voluntaryTerms;
              totalInvoluntary += pMetrics.involuntaryTerms;
              totalRegrettable += pMetrics.regrettableTerms;
            });
            
            const openingHC = firstMetrics.openingHC;
            const closingHC = lastMetrics.closingHC;
            const avgHC = (openingHC + closingHC) / 2;
            
            const attrition = avgHC > 0 ? (totalVoluntary / avgHC) : 0;
            const retention = openingHC > 0 ? ((openingHC - totalTerms) / openingHC) : 0;
            const turnover = avgHC > 0 ? (totalTerms / avgHC) : 0;
            const regrettableTurnover = avgHC > 0 ? (totalRegrettable / avgHC) : 0;
            
            metricsCache[cacheKey] = {
              openingHC: openingHC,
              closingHC: closingHC,
              hires: totalHires,
              terms: totalTerms,
              voluntaryTerms: totalVoluntary,
              involuntaryTerms: totalInvoluntary,
              regrettableTerms: totalRegrettable,
              attrition: Math.round(attrition * 10000) / 10000,
              retention: Math.round(retention * 10000) / 10000,
              turnover: Math.round(turnover * 10000) / 10000,
              regrettableTurnover: Math.round(regrettableTurnover * 10000) / 10000
            };
          } else {
            metricsCache[cacheKey] = calculateHRMetrics(period.start, period.end, periodFilters);
          }
        } catch (error) {
          Logger.log(`Error calculating metrics for ${elt} in ${period.label}: ${error.message}`);
          metricsCache[cacheKey] = null;
        }
      }
    });
  });
  
  // Build data rows using cached metrics
  const dataRows = [];
  elts.forEach((elt, eltIndex) => {
    // Add header row at the start of each ELT section
    dataRows.push(headerRow);
    
    // For each ELT, create rows for each metric
    metricLabels.forEach((metricLabel, metricIndex) => {
      const row = [`${elt} - ${metricLabel}`];
      
      periods.forEach(period => {
        const cacheKey = `${elt}|${period.label}`;
        const metrics = metricsCache[cacheKey];
        
        if (!metrics) {
          row.push(metricIndex < 7 ? "" : 0);
          return;
        }
        
        let value = 0;
        switch(metricIndex) {
          case 0: // Opening HC
            value = metrics.openingHC;
            break;
          case 1: // Closing HC
            value = metrics.closingHC;
            break;
          case 2: // Average Headcount - round to nearest integer
            value = Math.round((metrics.openingHC + metrics.closingHC) / 2);
            break;
          case 3: // Hires
            value = metrics.hires;
            break;
          case 4: // Terminations
            value = metrics.terms;
            break;
          case 5: // Regrettable Terms
            value = metrics.regrettableTerms;
            break;
          case 6: // Non-Regrettable Terms
            value = metrics.terms - metrics.regrettableTerms;
            break;
          case 7: // Retention %
            value = metrics.retention;
            break;
          case 8: // Turnover %
            value = metrics.turnover;
            break;
          case 9: // Regrettable %
            value = metrics.regrettableTurnover;
            break;
        }
        // For count metrics (0-6), use blank instead of 0. For percentages (7-9), keep 0
        if (metricIndex < 7 && value === 0) {
          row.push(""); // Blank for zero counts
        } else {
          row.push(value);
        }
      });
      dataRows.push(row);
    });
    
    // Add 1 blank row after each ELT (except the last one)
    if (eltIndex < elts.length - 1) {
      const blankRow = [""].concat(new Array(periods.length).fill(""));
      dataRows.push(blankRow);
    }
  });
  
  // Write data in one batch operation - only values, preserve all user formatting
  if (dataRows.length > 0) {
    try {
      eltSheet.getRange(1, 1, dataRows.length, headerRow.length).setValues(dataRows);
    } catch (error) {
      Logger.log(`Error writing ELT metrics: ${error.message}`);
      SpreadsheetApp.getUi().alert(`Error generating ELT metrics: ${error.message}`);
      return;
    }
  }
  
  SpreadsheetApp.getUi().alert(`ELT Metrics generated for ${elts.length} ELTs with ${metricLabels.length} metrics each, across ${periods.length} periods.`);
}

/**
 * Generates job level headcount as of today with site-level breakdowns
 * Creates tables with job levels sorted in specific order, plus charts
 */
function generateJobLevelHeadcount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let jobLevelSheet = ss.getSheetByName("Job Level Headcount");
  const isNewSheet = !jobLevelSheet;
  
  if (!jobLevelSheet) {
    jobLevelSheet = ss.insertSheet("Job Level Headcount");
  }
  
  const rawDataSheet = ss.getSheetByName("RawData");
  if (!rawDataSheet) {
    SpreadsheetApp.getUi().alert("RawData sheet not found. Please run 'Fetch Bob Report' first.");
    return;
  }
  
  const data = rawDataSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert("No data found in RawData sheet.");
    return;
  }
  
  const rows = data.slice(1);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  // Helper function to parse date
  const parseDate = (dateStr) => {
    if (!dateStr) return null;
    let date;
    if (dateStr instanceof Date) {
      date = dateStr;
    } else if (typeof dateStr === 'string') {
      date = new Date(dateStr);
    }
    if (date && !isNaN(date.getTime())) {
      return new Date(date.getFullYear(), date.getMonth(), date.getDate());
    }
    return null;
  };
  
  // Count unique employees by job level, site, and ELT (as of today)
  // Structure: jobLevel -> site -> Set of employee names
  // Structure: jobLevel -> elt -> Set of employee names
  const jobLevelSiteCounts = new Map();
  const jobLevelELTCounts = new Map();
  const allSites = new Set();
  const allELTs = new Set();
  
  rows.forEach(row => {
    const empName = String(row[COLUMN_INDICES.EMP_NAME] || "").trim();
    if (!empName) return;
    
    const startDate = parseDate(row[COLUMN_INDICES.START_DATE]);
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    const jobLevel = String(row[COLUMN_INDICES.JOB_LEVEL] || "").trim();
    const site = String(row[COLUMN_INDICES.SITE] || "").trim();
    const elt = String(row[COLUMN_INDICES.ELT] || "").trim();
    
    // Check if employee is active as of today
    if (!startDate || startDate > today) return; // Not started yet
    if (termDate && termDate <= today) return; // Already terminated
    
    if (!jobLevel || !site || !elt) return;
    
    // Track all sites and ELTs
    allSites.add(site);
    allELTs.add(elt);
    
    // Count unique employees by job level and site
    if (!jobLevelSiteCounts.has(jobLevel)) {
      jobLevelSiteCounts.set(jobLevel, new Map());
    }
    const siteMap = jobLevelSiteCounts.get(jobLevel);
    if (!siteMap.has(site)) {
      siteMap.set(site, new Set());
    }
    siteMap.get(site).add(empName);
    
    // Count unique employees by job level and ELT
    if (!jobLevelELTCounts.has(jobLevel)) {
      jobLevelELTCounts.set(jobLevel, new Map());
    }
    const eltMap = jobLevelELTCounts.get(jobLevel);
    if (!eltMap.has(elt)) {
      eltMap.set(elt, new Set());
    }
    eltMap.get(elt).add(empName);
  });
  
  // Get sorted sites and ELTs
  const sortedSites = Array.from(allSites).sort();
  const sortedELTs = Array.from(allELTs).sort();
  
  // Get all job levels that exist in data, sorted according to specified order
  const allJobLevels = Array.from(jobLevelSiteCounts.keys());
  const sortedJobLevels = allJobLevels.sort((a, b) => {
    const indexA = JOB_LEVEL_ORDER.indexOf(a);
    const indexB = JOB_LEVEL_ORDER.indexOf(b);
    
    // If both in order, sort by order
    if (indexA !== -1 && indexB !== -1) return indexA - indexB;
    // If only A in order, A comes first
    if (indexA !== -1) return -1;
    // If only B in order, B comes first
    if (indexB !== -1) return 1;
    // If neither in order, alphabetical
    return a.localeCompare(b);
  });
  
  // Build Overall Summary table (Job Level, Total Headcount)
  const overallHeaders = ["Job Level", "Headcount"];
  jobLevelSheet.getRange(1, 1, 1, 2).setValues([overallHeaders]);
  if (isNewSheet) {
    jobLevelSheet.getRange(1, 1, 1, 2).setFontWeight("bold");
  }
  
  const overallDataRows = sortedJobLevels.map(level => {
    const siteMap = jobLevelSiteCounts.get(level);
    let total = 0;
    siteMap.forEach(siteSet => {
      total += siteSet.size;
    });
    return [level, total];
  });
  
  if (overallDataRows.length > 0) {
    jobLevelSheet.getRange(2, 1, overallDataRows.length, 2).setValues(overallDataRows);
  }
  
  // Build Site Breakdown table (Job Level, Site1, Site2, ..., Total)
  const siteBreakdownStartRow = overallDataRows.length + 4; // Leave 2 blank rows
  const siteHeaders = ["Job Level"].concat(sortedSites).concat(["Total"]);
  jobLevelSheet.getRange(siteBreakdownStartRow, 1, 1, siteHeaders.length).setValues([siteHeaders]);
  if (isNewSheet) {
    jobLevelSheet.getRange(siteBreakdownStartRow, 1, 1, siteHeaders.length).setFontWeight("bold");
  }
  
  const siteBreakdownRows = sortedJobLevels.map(level => {
    const siteMap = jobLevelSiteCounts.get(level);
    const row = [level];
    let total = 0;
    
    sortedSites.forEach(site => {
      const count = siteMap.has(site) ? siteMap.get(site).size : 0;
      row.push(count);
      total += count;
    });
    
    row.push(total);
    return row;
  });
  
  if (siteBreakdownRows.length > 0) {
    jobLevelSheet.getRange(siteBreakdownStartRow + 1, 1, siteBreakdownRows.length, siteHeaders.length).setValues(siteBreakdownRows);
  }
  
  // Build ELT Breakdown table (Job Level, ELT1, ELT2, ..., Total)
  const eltBreakdownStartRow = siteBreakdownStartRow + siteBreakdownRows.length + 4; // Leave 2 blank rows after site breakdown
  
  // Add dropdown filter for ELT selection (above the ELT breakdown table)
  const eltFilterRow = eltBreakdownStartRow - 1;
  jobLevelSheet.getRange(eltFilterRow, 1).setValue("Filter by ELT:");
  if (isNewSheet) {
    jobLevelSheet.getRange(eltFilterRow, 1).setFontWeight("bold");
  }
  
  // Create dropdown with "All" option plus all ELTs
  const eltFilterOptions = ["All"].concat(sortedELTs);
  const eltFilterRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(eltFilterOptions, true)
    .build();
  jobLevelSheet.getRange(eltFilterRow, 2).setDataValidation(eltFilterRule);
  jobLevelSheet.getRange(eltFilterRow, 2).setValue("All"); // Default to "All"
  
  const eltHeaders = ["Job Level"].concat(sortedELTs).concat(["Total"]);
  jobLevelSheet.getRange(eltBreakdownStartRow, 1, 1, eltHeaders.length).setValues([eltHeaders]);
  if (isNewSheet) {
    jobLevelSheet.getRange(eltBreakdownStartRow, 1, 1, eltHeaders.length).setFontWeight("bold");
  }
  
  const eltBreakdownRows = sortedJobLevels.map(level => {
    const eltMap = jobLevelELTCounts.get(level);
    const row = [level];
    let total = 0;
    
    sortedELTs.forEach(elt => {
      const count = eltMap.has(elt) ? eltMap.get(elt).size : 0;
      row.push(count);
      total += count;
    });
    
    row.push(total);
    return row;
  });
  
  if (eltBreakdownRows.length > 0) {
    jobLevelSheet.getRange(eltBreakdownStartRow + 1, 1, eltBreakdownRows.length, eltHeaders.length).setValues(eltBreakdownRows);
  }
  
  // Create filtered data range for chart (starts below the ELT breakdown table)
  // This will be a formula-based table that filters based on the dropdown selection
  const filteredEltStartRow = eltBreakdownStartRow + eltBreakdownRows.length + 3; // Leave 2 blank rows
  const filteredHeadersRow = filteredEltStartRow + 1;
  
  // Create filtered data table with formulas
  // Header row: Job Level + selected ELT column (or all columns if "All")
  const filterCellRef = `$B$${eltFilterRow}`;
  const dataStartCol = eltBreakdownStartRow;
  const dataEndCol = eltBreakdownStartRow + eltBreakdownRows.length;
  
  // Helper function to convert column number to letter (1=A, 2=B, etc.)
  const colNumToLetter = (num) => {
    let letter = '';
    while (num > 0) {
      const remainder = (num - 1) % 26;
      letter = String.fromCharCode(65 + remainder) + letter;
      num = Math.floor((num - 1) / 26);
    }
    return letter;
  };
  
  // For each job level row, create a formula that shows Job Level + selected ELT column
  // Data rows start one row BELOW the header row
  sortedJobLevels.forEach((level, levelIndex) => {
    const dataRow = eltBreakdownStartRow + 1 + levelIndex; // +1 for header row in original ELT breakdown table
    const filteredDataRow = filteredHeadersRow + 1 + levelIndex; // +1 to start BELOW the header row
    
    // Job Level column (always show) - Column A
    jobLevelSheet.getRange(filteredDataRow, 1).setFormula(`=A${dataRow}`);
    
    // ELT column(s) - if "All", show all ELT columns, otherwise show only selected ELT
    if (sortedELTs.length > 0) {
      sortedELTs.forEach((elt, eltIndex) => {
        const eltColIndex = eltIndex + 2; // +2 because column A is Job Level, column B is first ELT
        const eltColLetter = colNumToLetter(eltColIndex);
        
        // Formula: If dropdown is "All", show all columns. If dropdown matches this ELT, show it. Otherwise blank.
        // Escape quotes in ELT name for formula
        const escapedElt = elt.replace(/"/g, '""');
        const formula = `=IF(${filterCellRef}="All", ${eltColLetter}${dataRow}, IF(${filterCellRef}="${escapedElt}", ${eltColLetter}${dataRow}, ""))`;
        jobLevelSheet.getRange(filteredDataRow, eltIndex + 2).setFormula(formula);
      });
    }
  });
  
  // Set header row for filtered data
  // Column A should always be "Job Level"
  jobLevelSheet.getRange(filteredHeadersRow, 1).setValue("Job Level");
  
  if (sortedELTs.length > 0) {
    // Header formula: If "All", show all ELT headers, otherwise show only selected ELT header
    sortedELTs.forEach((elt, eltIndex) => {
      const eltColIndex = eltIndex + 2;
      const eltColLetter = colNumToLetter(eltColIndex);
      const escapedElt = elt.replace(/"/g, '""');
      const headerFormula = `=IF(${filterCellRef}="All", ${eltColLetter}${eltBreakdownStartRow}, IF(${filterCellRef}="${escapedElt}", ${eltColLetter}${eltBreakdownStartRow}, ""))`;
      jobLevelSheet.getRange(filteredHeadersRow, eltIndex + 2).setFormula(headerFormula);
    });
  }
  
  // Create the filtered chart range (excludes empty columns when filtering)
  // We'll use a dynamic range that adjusts based on the filter
  const filteredChartNumCols = sortedELTs.length + 1; // Job Level + all ELTs (formulas will handle filtering)
  const filteredChartRange = jobLevelSheet.getRange(filteredHeadersRow, 1, sortedJobLevels.length + 1, filteredChartNumCols);
  
  // Auto-resize columns if new sheet
  if (isNewSheet) {
    jobLevelSheet.autoResizeColumns(1, Math.max(2, Math.max(siteHeaders.length, eltHeaders.length)));
  }
  
  // Create or update charts
  try {
    const existingCharts = jobLevelSheet.getCharts();
    let overallChart = null;
    let siteChart = null;
    let eltChart = null;
    
    // Define expected ranges for each chart type
    const overallChartRange = jobLevelSheet.getRange(1, 1, overallDataRows.length + 1, 2);
    const siteChartNumCols = sortedSites.length + 1; // Job Level + sites (excluding Total)
    const siteChartRange = jobLevelSheet.getRange(siteBreakdownStartRow, 1, siteBreakdownRows.length + 1, siteChartNumCols);
    const eltChartNumCols = sortedELTs.length + 1; // Job Level + ELTs (excluding Total)
    const eltChartRange = jobLevelSheet.getRange(eltBreakdownStartRow, 1, eltBreakdownRows.length + 1, eltChartNumCols);
    
    // Find existing charts by title and/or range
    // We'll check all charts and try multiple identification methods
    Logger.log(`Checking ${existingCharts.length} existing charts for matches...`);
    
    // Log all chart titles and ranges for debugging
    existingCharts.forEach((chart, idx) => {
      try {
        const opts = chart.getOptions();
        const t = opts ? opts.title : null;
        const ranges = chart.getRanges();
        const r = ranges && ranges.length > 0 ? ranges[0] : null;
        Logger.log(`Chart ${idx}: title="${t}", range=${r ? r.getA1Notation() : 'N/A'}, rows=${r ? r.getNumRows() : 'N/A'}, cols=${r ? r.getNumColumns() : 'N/A'}`);
      } catch (e) {
        Logger.log(`Chart ${idx}: Error reading - ${e.message}`);
      }
    });
    
    existingCharts.forEach((chart, index) => {
      try {
        const options = chart.getOptions();
        const title = options ? options.title : null;
        
        // Get chart ranges for range-based matching
        let chartRanges = null;
        let firstRange = null;
        try {
          chartRanges = chart.getRanges();
          if (chartRanges && chartRanges.length > 0) {
            firstRange = chartRanges[0];
          }
        } catch (rangeError) {
          Logger.log(`Error reading chart ${index} ranges: ${rangeError.message}`);
        }
        
        // Try to identify Overall chart
        if (!overallChart) {
          const matchesTitle = title && (
            title.includes('Overall') || 
            title.includes('Job Level Headcount (Overall)') ||
            (title.toLowerCase().includes('overall') && title.toLowerCase().includes('headcount'))
          );
          // More flexible range matching: starts at A1 (row 1, col 1) with 2 columns
          // Don't check exact row count since data might change
          const matchesRange = firstRange && 
            firstRange.getRow() === 1 && 
            firstRange.getColumn() === 1 && 
            firstRange.getNumColumns() === 2;
          
          if (matchesTitle || matchesRange) {
            overallChart = chart;
            Logger.log(`Found overall chart (index ${index}) - title: "${title}", range: ${firstRange ? firstRange.getA1Notation() : 'N/A'}, numRows: ${firstRange ? firstRange.getNumRows() : 'N/A'}, numCols: ${firstRange ? firstRange.getNumColumns() : 'N/A'}`);
          }
        }
        
        // Try to identify Site chart
        if (!siteChart) {
          const matchesTitle = title && title.includes('by Site');
          const matchesRange = firstRange && firstRange.getRow() === siteBreakdownStartRow && firstRange.getColumn() === 1 && firstRange.getNumColumns() === siteChartNumCols;
          
          if (matchesTitle || matchesRange) {
            siteChart = chart;
            Logger.log(`Found site chart (index ${index}) - title: ${title}, range: ${firstRange ? firstRange.getA1Notation() : 'N/A'}`);
          }
        }
        
        // Try to identify ELT chart
        if (!eltChart) {
          const matchesTitle = title && title.includes('by ELT');
          const matchesRange = firstRange && (
            (firstRange.getRow() === eltBreakdownStartRow && firstRange.getColumn() === 1 && firstRange.getNumColumns() === eltChartNumCols) ||
            (firstRange.getRow() === filteredHeadersRow && firstRange.getColumn() === 1)
          );
          
          if (matchesTitle || matchesRange) {
            eltChart = chart;
            Logger.log(`Found ELT chart (index ${index}) - title: ${title}, range: ${firstRange ? firstRange.getA1Notation() : 'N/A'}`);
          }
        }
      } catch (e) {
        // Skip charts that can't be read
        Logger.log(`Error reading chart ${index}: ${e.message}`);
      }
    });
    
    Logger.log(`Chart detection results - Overall: ${overallChart ? 'FOUND' : 'NOT FOUND'}, Site: ${siteChart ? 'FOUND' : 'NOT FOUND'}, ELT: ${eltChart ? 'FOUND' : 'NOT FOUND'}`);
    
    // Chart 1: Overall Job Level Headcount (Column Chart)
    if (overallChart) {
      // Update existing chart - ONLY update data range, preserve all formatting
      try {
        Logger.log(`Updating existing overall chart - current range: ${overallChart.getRanges()[0] ? overallChart.getRanges()[0].getA1Notation() : 'N/A'}, new range: ${overallChartRange.getA1Notation()}`);
        // Only modify the range, don't touch any other options to preserve user formatting
        const updatedChart = overallChart.modify()
          .clearRanges()
          .addRange(overallChartRange)
          .build();
        jobLevelSheet.updateChart(updatedChart);
        Logger.log("Overall chart updated successfully (formatting preserved)");
      } catch (e) {
        Logger.log(`ERROR updating overall chart: ${e.message}`);
        Logger.log(`Stack trace: ${e.stack}`);
        // Don't create a new chart if update fails - this prevents duplicates
      }
    } else {
      // Only create new chart if it doesn't exist AND we're absolutely sure
      // Double-check one more time before creating
      const doubleCheckCharts = jobLevelSheet.getCharts();
      let foundInDoubleCheck = false;
      for (const ch of doubleCheckCharts) {
        try {
          const opts = ch.getOptions();
          const t = opts ? opts.title : null;
          const ranges = ch.getRanges();
          const r = ranges && ranges.length > 0 ? ranges[0] : null;
          if ((t && (t.includes('Overall') || t.includes('Job Level Headcount (Overall)'))) ||
              (r && r.getRow() === 1 && r.getColumn() === 1 && r.getNumColumns() === 2)) {
            foundInDoubleCheck = true;
            Logger.log(`Found overall chart in double-check - NOT creating new one`);
            break;
          }
        } catch (e) {
          // Continue checking
        }
      }
      
      if (!foundInDoubleCheck) {
        Logger.log("Overall chart not found after double-check, creating new chart");
        const newOverallChart = jobLevelSheet.newChart()
          .setChartType(Charts.ChartType.COLUMN)
          .addRange(overallChartRange)
          .setPosition(overallDataRows.length + 2, 4, 0, 0)
          .setOption('title', 'Job Level Headcount (Overall)')
          .setOption('legend.position', 'none')
          .setOption('hAxis.title', 'Job Level')
          .setOption('vAxis.title', 'Headcount')
          .setOption('vAxis.textStyle.color', '#FFFFFF') // White font to hide axis values
          .setOption('vAxis.gridlines.count', 0) // Turn off major gridlines
          .setOption('vAxis.minorGridlines.count', 0) // Turn off minor gridlines
          .setOption('dataLabelPosition', 'out') // Data labels outside end
          .setOption('useFirstColumnAsDomain', true) // Use Col A as labels
          .setOption('width', 600)
          .setOption('height', 400)
          .build();
        jobLevelSheet.insertChart(newOverallChart);
        Logger.log("New overall chart created with formatting");
      }
    }
    
    // Chart 2: Site Breakdown (Stacked Column Chart)
    // Exclude the "Total" column - only include Job Level + Site columns
    if (siteChart) {
      // Update existing chart - ONLY update data range, preserve all formatting
      try {
        const updatedChart = siteChart.modify()
          .clearRanges()
          .addRange(siteChartRange)
          .build();
        jobLevelSheet.updateChart(updatedChart);
        Logger.log("Site chart updated (formatting preserved)");
      } catch (e) {
        Logger.log(`Error updating site chart: ${e.message}`);
      }
    } else {
      // Create new chart with formatting
      const newSiteChart = jobLevelSheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(siteChartRange)
        .setPosition(siteBreakdownStartRow + siteBreakdownRows.length + 2, 4, 0, 0)
        .setOption('title', 'Job Level Headcount by Site')
        .setOption('isStacked', true)
        .setOption('legend.position', 'right')
        .setOption('hAxis.title', 'Job Level')
        .setOption('vAxis.title', 'Headcount')
        .setOption('vAxis.textStyle.color', '#FFFFFF') // White font to hide axis values
        .setOption('vAxis.gridlines.count', 0) // Turn off major gridlines
        .setOption('vAxis.minorGridlines.count', 0) // Turn off minor gridlines
        .setOption('dataLabelPosition', 'out') // Data labels outside end
        .setOption('useFirstColumnAsDomain', true) // Use Col A as labels
        .setOption('width', 800)
        .setOption('height', 400)
        .build();
      jobLevelSheet.insertChart(newSiteChart);
      Logger.log("New site chart created with formatting");
    }
    
    // Chart 3: ELT Breakdown (Stacked Column Chart) - Interactive with dropdown filter
    // The chart references the filtered data table which updates automatically based on dropdown
    if (eltChart) {
      // Update existing chart - ONLY update data range, preserve all formatting
      try {
        const updatedChart = eltChart.modify()
          .clearRanges()
          .addRange(filteredChartRange)
          .build();
        jobLevelSheet.updateChart(updatedChart);
        Logger.log("ELT chart updated (formatting preserved)");
      } catch (e) {
        Logger.log(`Error updating ELT chart: ${e.message}`);
      }
    } else {
      // Only create new chart if it doesn't exist
      Logger.log("ELT chart not found, creating new chart");
      const newELTChart = jobLevelSheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(filteredChartRange)
        .setPosition(eltBreakdownStartRow + eltBreakdownRows.length + 2, 4, 0, 0)
        .setOption('title', 'Job Level Headcount by ELT (Use dropdown above to filter)')
        .setOption('isStacked', true)
        .setOption('legend.position', 'right')
        .setOption('hAxis.title', 'Job Level')
        .setOption('vAxis.title', 'Headcount')
        .setOption('vAxis.textStyle.color', '#FFFFFF') // White font to hide axis values
        .setOption('vAxis.gridlines.count', 0) // Turn off major gridlines
        .setOption('vAxis.minorGridlines.count', 0) // Turn off minor gridlines
        .setOption('dataLabelPosition', 'out') // Data labels outside end
        .setOption('useFirstColumnAsDomain', true) // Use Col A as labels
        .setOption('width', 800)
        .setOption('height', 400)
        .build();
      jobLevelSheet.insertChart(newELTChart);
      Logger.log("New ELT chart created with formatting");
    }
    
    // Add instruction text
    jobLevelSheet.getRange(eltFilterRow, 3).setValue("(Change dropdown to filter chart by ELT - chart updates automatically)");
    if (isNewSheet) {
      jobLevelSheet.getRange(eltFilterRow, 3).setFontStyle("italic");
      jobLevelSheet.getRange(eltFilterRow, 3).setFontColor("#666666");
    }
  } catch (e) {
    Logger.log(`Error creating/updating charts: ${e.message}`);
    // Charts are optional, continue even if they fail
  }
  
  SpreadsheetApp.getUi().alert(`Job Level Headcount generated for ${sortedJobLevels.length} job levels across ${sortedSites.length} sites and ${sortedELTs.length} ELTs as of today.`);
}

/**
 * Creates/updates a Termination Reason Mapping sheet for editing remappings
 * Users can edit this sheet to map old reasons to new reasons
 */
function createTerminationReasonMappingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let mappingSheet = ss.getSheetByName("Termination Reason Mapping");
  
  if (!mappingSheet) {
    mappingSheet = ss.insertSheet("Termination Reason Mapping");
  } else {
    mappingSheet.clear();
  }
  
  // Get unique termination reasons from RawData
  const rawDataSheet = ss.getSheetByName("RawData");
  if (!rawDataSheet) {
    SpreadsheetApp.getUi().alert("RawData sheet not found. Please run 'Fetch Bob Report' first.");
    return;
  }
  
  const data = rawDataSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert("No data found in RawData sheet.");
    return;
  }
  
  const rows = data.slice(1);
  const uniqueReasons = new Set();
  
  rows.forEach(row => {
    const reason = String(row[COLUMN_INDICES.TERMINATION_REASON] || "").trim();
    if (reason) {
      uniqueReasons.add(reason);
    }
  });
  
  // Create mapping table
  const headers = ["Original Reason", "Mapped To", "Instructions"];
  mappingSheet.getRange(1, 1, 1, 3).setValues([headers]);
  mappingSheet.getRange(1, 1, 1, 3).setFontWeight("bold");
  
  // Add all unique reasons
  const sortedReasons = Array.from(uniqueReasons).sort();
  const mappingRows = sortedReasons.map(reason => [reason, reason, "Edit 'Mapped To' column to remap this reason"]);
  
  if (mappingRows.length > 0) {
    mappingSheet.getRange(2, 1, mappingRows.length, 3).setValues(mappingRows);
    
    // Add data validation for "Mapped To" column - allow any text or selection from existing reasons
    const mappedToRange = mappingSheet.getRange(2, 2, mappingRows.length, 1);
    // Allow free text input (no strict validation, but users can type or use dropdown)
    mappingSheet.getRange(2, 3, mappingRows.length, 1).setFontStyle("italic");
    mappingSheet.getRange(2, 3, mappingRows.length, 1).setFontColor("#666666");
  }
  
  // Add instructions
  const instructionRow = mappingRows.length + 3;
  mappingSheet.getRange(instructionRow, 1).setValue("Instructions:");
  mappingSheet.getRange(instructionRow, 1).setFontWeight("bold");
  mappingSheet.getRange(instructionRow + 1, 1).setValue("1. Edit the 'Mapped To' column to remap termination reasons");
  mappingSheet.getRange(instructionRow + 2, 1).setValue("2. Leave 'Mapped To' same as 'Original Reason' to keep it unchanged");
  mappingSheet.getRange(instructionRow + 3, 1).setValue("3. Use the same 'Mapped To' value for multiple reasons to consolidate them");
  mappingSheet.getRange(instructionRow + 4, 1).setValue("4. Run 'Map Termination Reasons' from menu to apply the mappings");
  
  mappingSheet.autoResizeColumns(1, 3);
  
  SpreadsheetApp.getUi().alert(`Termination Reason Mapping sheet created with ${sortedReasons.length} unique reasons. Edit the 'Mapped To' column and run 'Map Termination Reasons' to apply.`);
}

/**
 * Maps/cleans termination reasons based on mapping sheet
 * Reads from "Termination Reason Mapping" sheet and applies mappings to RawData
 */
function mapTerminationReasons() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawDataSheet = ss.getSheetByName("RawData");
  
  if (!rawDataSheet) {
    SpreadsheetApp.getUi().alert("RawData sheet not found. Please run 'Fetch Bob Report' first.");
    return;
  }
  
  const mappingSheet = ss.getSheetByName("Termination Reason Mapping");
  if (!mappingSheet) {
    SpreadsheetApp.getUi().alert("Termination Reason Mapping sheet not found. Please run 'Create Termination Reason Mapping Sheet' first.");
    return;
  }
  
  // Read mapping from sheet
  const mappingData = mappingSheet.getDataRange().getValues();
  if (mappingData.length <= 1) {
    SpreadsheetApp.getUi().alert("No mappings found in Termination Reason Mapping sheet.");
    return;
  }
  
  const mappingRows = mappingData.slice(1); // Skip header
  const reasonMapping = {};
  
  mappingRows.forEach(row => {
    const originalReason = String(row[0] || "").trim();
    const mappedTo = String(row[1] || "").trim();
    if (originalReason && mappedTo) {
      reasonMapping[originalReason] = mappedTo;
    }
  });
  
  if (Object.keys(reasonMapping).length === 0) {
    SpreadsheetApp.getUi().alert("No valid mappings found. Please edit the 'Mapped To' column in Termination Reason Mapping sheet.");
    return;
  }
  
  // Apply mappings to RawData
  const data = rawDataSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert("No data found in RawData sheet.");
    return;
  }
  
  const header = data[0];
  const rows = data.slice(1);
  
  let updateCount = 0;
  rows.forEach((row, index) => {
    const currentReason = String(row[COLUMN_INDICES.TERMINATION_REASON] || "").trim();
    if (currentReason && reasonMapping.hasOwnProperty(currentReason)) {
      const newReason = reasonMapping[currentReason];
      if (currentReason !== newReason) {
        row[COLUMN_INDICES.TERMINATION_REASON] = newReason;
        updateCount++;
      }
    }
  });
  
  if (updateCount > 0) {
    const updatedData = [header, ...rows];
    const targetRange = rawDataSheet.getRange(1, 1, updatedData.length, updatedData[0].length);
    targetRange.setValues(updatedData);
    SpreadsheetApp.getUi().alert(`Termination reasons mapped. ${updateCount} rows updated.`);
  } else {
    SpreadsheetApp.getUi().alert("No changes needed. All termination reasons are already mapped correctly.");
  }
}

/**
 * Generates termination reasons table for pie chart
 * Filterable by time period (month/year, quarter, year), site, department, ELT
 */
function generateTerminationReasonsTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let termReasonsSheet = ss.getSheetByName("Termination Reasons");
  const isNewSheet = !termReasonsSheet;
  
  if (!termReasonsSheet) {
    termReasonsSheet = ss.insertSheet("Termination Reasons");
  }
  
  const rawDataSheet = ss.getSheetByName("RawData");
  if (!rawDataSheet) {
    SpreadsheetApp.getUi().alert("RawData sheet not found. Please run 'Fetch Bob Report' first.");
    return;
  }
  
  const data = rawDataSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert("No data found in RawData sheet.");
    return;
  }
  
  const rows = data.slice(1);
  
  // No filters needed - we'll show breakdowns by Site and ELT, plus overall
  const filterConfigSheet = ss.getSheetByName("FilterConfig");
  
  // Get time period filters for Termination Reasons (separate section)
  let timePeriodFilter = null;
  if (filterConfigSheet) {
    // Use Termination Reasons filter section (rows 17-18)
    const termYearFilter = filterConfigSheet.getRange(17, 2).getValue(); // Year filter for Termination Reasons
    const periodTypeFilter = filterConfigSheet.getRange(18, 2).getValue(); // Period Type: Halves or Quarters
    
    if (termYearFilter) {
      const year = parseInt(termYearFilter, 10);
      if (periodTypeFilter === "Halves") {
        // Year + Halves: filter by both H1 and H2 (show both)
        timePeriodFilter = { type: 'year_halves', value: year };
      } else if (periodTypeFilter === "Quarters") {
        // Year + Quarters: filter by all quarters (show Q1-Q4)
        timePeriodFilter = { type: 'year_quarters', value: year };
      } else {
        // Year only: show all months in year
        timePeriodFilter = { type: 'year', value: year };
      }
    }
  }
  
  // Helper function to parse date
  const parseDate = (dateStr) => {
    if (!dateStr) return null;
    let date;
    if (dateStr instanceof Date) {
      date = dateStr;
    } else if (typeof dateStr === 'number') {
      // Google Sheets serial dates
      const sheetsEpoch = new Date(1899, 11, 30);
      date = new Date(sheetsEpoch.getTime() + dateStr * 86400000);
    } else if (typeof dateStr === 'string') {
      date = new Date(dateStr);
    }
    if (date && !isNaN(date.getTime())) {
      return new Date(date.getFullYear(), date.getMonth(), date.getDate());
    }
    return null;
  };
  
  // Helper function to check if date matches time period filter
  const matchesTimePeriod = (termDate) => {
    if (!timePeriodFilter || !termDate) return true; // No filter = include all
    
    const termYear = termDate.getFullYear();
    const termMonth = termDate.getMonth();
    const termQuarter = Math.floor(termMonth / 3) + 1;
    const termHalf = termMonth < 6 ? 1 : 2; // H1 = Jan-Jun (months 0-5), H2 = Jul-Dec (months 6-11)
    
    if (timePeriodFilter.type === 'year') {
      return termYear === timePeriodFilter.value;
    } else if (timePeriodFilter.type === 'year_halves') {
      // Year + Halves: include all months in the year (both H1 and H2)
      return termYear === timePeriodFilter.value;
    } else if (timePeriodFilter.type === 'year_quarters') {
      // Year + Quarters: include all months in the year (all quarters)
      return termYear === timePeriodFilter.value;
    } else if (timePeriodFilter.type === 'halves') {
      return termHalf === timePeriodFilter.value;
    } else if (timePeriodFilter.type === 'quarter') {
      return termQuarter === timePeriodFilter.value;
    }
    return true;
  };
  
  // Count terminations by reason with filters applied
  const overallReasonCounts = new Map(); // Overall counts: reason -> Set of employee names
  const siteReasonCounts = new Map(); // Site counts: site -> Map(reason -> Set of employee names)
  const eltReasonCounts = new Map(); // ELT counts: elt -> Map(reason -> Set of employee names)
  
  rows.forEach(row => {
    const empName = String(row[COLUMN_INDICES.EMP_NAME] || "").trim();
    if (!empName) return;
    
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    const termReason = String(row[COLUMN_INDICES.TERMINATION_REASON] || "").trim();
    const site = String(row[COLUMN_INDICES.SITE] || "").trim();
    const elt = String(row[COLUMN_INDICES.ELT] || "").trim();
    
    if (!termDate || !termReason) return; // Skip if no termination date or reason
    
    // Apply time period filter
    if (!matchesTimePeriod(termDate)) return;
    
    // Count unique employees by termination reason, site, and ELT
    // Overall counts
    if (!overallReasonCounts.has(termReason)) {
      overallReasonCounts.set(termReason, new Set());
    }
    overallReasonCounts.get(termReason).add(empName);
    
    // Site counts
    if (site) {
      if (!siteReasonCounts.has(site)) {
        siteReasonCounts.set(site, new Map());
      }
      const siteMap = siteReasonCounts.get(site);
      if (!siteMap.has(termReason)) {
        siteMap.set(termReason, new Set());
      }
      siteMap.get(termReason).add(empName);
    }
    
    // ELT counts
    if (elt) {
      if (!eltReasonCounts.has(elt)) {
        eltReasonCounts.set(elt, new Map());
      }
      const eltMap = eltReasonCounts.get(elt);
      if (!eltMap.has(termReason)) {
        eltMap.set(termReason, new Set());
      }
      eltMap.get(termReason).add(empName);
    }
  });
  
  // Get unique sites and ELTs for breakdowns
  const uniqueValues = getUniqueFilterValues();
  const sortedSites = uniqueValues.sites.sort();
  const sortedELTs = uniqueValues.elts.sort();
  
  // Column assignments:
  // Overall: Columns A, B, C (columns 1, 2, 3)
  // By Site: Columns E, F, G (columns 5, 6, 7)
  // By ELT: Columns I, J, K (columns 9, 10, 11)
  const OVERALL_COL_START = 1; // Column A
  const SITE_COL_START = 5;    // Column E
  const ELT_COL_START = 9;     // Column I
  
  // Clear rows 1-2 in column A first (to remove any old filter info)
  termReasonsSheet.getRange(1, OVERALL_COL_START, 2, 1).clearContent();
  
  // Add filter info in rows 1 and 2 of column A (before all tables)
  if (timePeriodFilter) {
    const filterInfo = ["Filters Applied:"];
    if (timePeriodFilter.type === 'year') {
      filterInfo.push(`Year: ${timePeriodFilter.value}`);
    } else if (timePeriodFilter.type === 'year_halves') {
      filterInfo.push(`Year: ${timePeriodFilter.value}, Period Type: Halves (H1/H2)`);
    } else if (timePeriodFilter.type === 'year_quarters') {
      filterInfo.push(`Year: ${timePeriodFilter.value}, Period Type: Quarters (Q1-Q4)`);
    } else if (timePeriodFilter.type === 'halves') {
      filterInfo.push(`Halves: H${timePeriodFilter.value}`);
    } else if (timePeriodFilter.type === 'quarter') {
      filterInfo.push(`Quarter: Q${timePeriodFilter.value}`);
    }
    termReasonsSheet.getRange(1, OVERALL_COL_START, filterInfo.length, 1).setValues(filterInfo.map(f => [f]));
  }
  
  // All tables start at row 3
  let overallRow = 3;
  let siteRow = 3;
  let eltRow = 3;
  
  // Build Overall table in columns A, B, C
  termReasonsSheet.getRange(overallRow, OVERALL_COL_START).setValue("Termination Reasons (Overall)");
  termReasonsSheet.getRange(overallRow, OVERALL_COL_START).setFontWeight("bold");
  termReasonsSheet.getRange(overallRow, OVERALL_COL_START).setFontSize(12);
  overallRow++;
  
  // Sort by count (descending)
  const sortedReasons = Array.from(overallReasonCounts.entries())
    .sort((a, b) => b[1].size - a[1].size);
  
  // Headers
  termReasonsSheet.getRange(overallRow, OVERALL_COL_START, 1, 3).setValues([["Termination Reason", "Count", "Percentage"]]);
  termReasonsSheet.getRange(overallRow, OVERALL_COL_START, 1, 3).setFontWeight("bold");
  overallRow++;
  
  const totalTerms = Array.from(overallReasonCounts.values())
    .reduce((sum, set) => sum + set.size, 0);
  
  const dataRows = sortedReasons.map(([reason, empSet]) => {
    const count = empSet.size;
    const percentage = totalTerms > 0 ? count / totalTerms : 0;
    return [reason, count, percentage];
  });
  
  if (dataRows.length > 0) {
    termReasonsSheet.getRange(overallRow, OVERALL_COL_START, dataRows.length, 3).setValues(dataRows);
    termReasonsSheet.getRange(overallRow, OVERALL_COL_START + 2, dataRows.length, 1).setNumberFormat("0.0%");
    
    // Add horizontal borders to each data row
    const borderColor = '#9900ff';
    for (let i = 0; i < dataRows.length; i++) {
      const rowRange = termReasonsSheet.getRange(overallRow + i, OVERALL_COL_START, 1, 3);
      rowRange.setBorder(true, false, false, false, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID);
    }
    
    overallRow += dataRows.length;
  }
  
  // Build Site breakdown tables in columns E, F, G
  if (sortedSites.length > 0) {
    termReasonsSheet.getRange(siteRow, SITE_COL_START).setValue("Termination Reasons by Site");
    termReasonsSheet.getRange(siteRow, SITE_COL_START).setFontWeight("bold");
    termReasonsSheet.getRange(siteRow, SITE_COL_START).setFontSize(12);
    siteRow++;
    
    sortedSites.forEach((site, siteIndex) => {
      const siteMap = siteReasonCounts.get(site);
      
      if (siteMap && siteMap.size > 0) {
        // Site header
        termReasonsSheet.getRange(siteRow, SITE_COL_START).setValue(`${site}:`);
        termReasonsSheet.getRange(siteRow, SITE_COL_START).setFontWeight("bold");
        siteRow++;
        
        // Headers
        termReasonsSheet.getRange(siteRow, SITE_COL_START, 1, 3).setValues([["Termination Reason", "Count", "Percentage"]]);
        termReasonsSheet.getRange(siteRow, SITE_COL_START, 1, 3).setFontWeight("bold");
        siteRow++;
        
        // Sort reasons by count
        const siteReasons = Array.from(siteMap.entries())
          .sort((a, b) => b[1].size - a[1].size);
        
        const siteTotal = Array.from(siteMap.values()).reduce((sum, set) => sum + set.size, 0);
        
        const siteDataRows = siteReasons.map(([reason, empSet]) => {
          const count = empSet.size;
          const percentage = siteTotal > 0 ? count / siteTotal : 0;
          return [reason, count, percentage];
        });
        
        if (siteDataRows.length > 0) {
          termReasonsSheet.getRange(siteRow, SITE_COL_START, siteDataRows.length, 3).setValues(siteDataRows);
          termReasonsSheet.getRange(siteRow, SITE_COL_START + 2, siteDataRows.length, 1).setNumberFormat("0.0%");
          
          // Add horizontal borders to each data row
          const borderColor = '#9900ff';
          for (let i = 0; i < siteDataRows.length; i++) {
            const rowRange = termReasonsSheet.getRange(siteRow + i, SITE_COL_START, 1, 3);
            rowRange.setBorder(true, false, false, false, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID);
          }
          
          siteRow += siteDataRows.length;
        }
        
        // Add blank row between sites
        if (siteIndex < sortedSites.length - 1) {
          siteRow++;
        }
      }
    });
  }
  
  // Build ELT breakdown tables in columns I, J, K
  if (sortedELTs.length > 0) {
    termReasonsSheet.getRange(eltRow, ELT_COL_START).setValue("Termination Reasons by ELT");
    termReasonsSheet.getRange(eltRow, ELT_COL_START).setFontWeight("bold");
    termReasonsSheet.getRange(eltRow, ELT_COL_START).setFontSize(12);
    eltRow++;
    
    sortedELTs.forEach((elt, eltIndex) => {
      const eltMap = eltReasonCounts.get(elt);
      
      if (eltMap && eltMap.size > 0) {
        // ELT header
        termReasonsSheet.getRange(eltRow, ELT_COL_START).setValue(`${elt}:`);
        termReasonsSheet.getRange(eltRow, ELT_COL_START).setFontWeight("bold");
        eltRow++;
        
        // Headers
        termReasonsSheet.getRange(eltRow, ELT_COL_START, 1, 3).setValues([["Termination Reason", "Count", "Percentage"]]);
        termReasonsSheet.getRange(eltRow, ELT_COL_START, 1, 3).setFontWeight("bold");
        eltRow++;
        
        // Sort reasons by count
        const eltReasons = Array.from(eltMap.entries())
          .sort((a, b) => b[1].size - a[1].size);
        
        const eltTotal = Array.from(eltMap.values()).reduce((sum, set) => sum + set.size, 0);
        
        const eltDataRows = eltReasons.map(([reason, empSet]) => {
          const count = empSet.size;
          const percentage = eltTotal > 0 ? count / eltTotal : 0;
          return [reason, count, percentage];
        });
        
        if (eltDataRows.length > 0) {
          termReasonsSheet.getRange(eltRow, ELT_COL_START, eltDataRows.length, 3).setValues(eltDataRows);
          termReasonsSheet.getRange(eltRow, ELT_COL_START + 2, eltDataRows.length, 1).setNumberFormat("0.0%");
          
          // Add horizontal borders to each data row
          const borderColor = '#9900ff';
          for (let i = 0; i < eltDataRows.length; i++) {
            const rowRange = termReasonsSheet.getRange(eltRow + i, ELT_COL_START, 1, 3);
            rowRange.setBorder(true, false, false, false, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID);
          }
          
          eltRow += eltDataRows.length;
        }
        
        // Add blank row between ELTs
        if (eltIndex < sortedELTs.length - 1) {
          eltRow++;
        }
      }
    });
  }
  
  // Filter info is already placed in rows 1-2 above, so no need to add it here
  
  if (isNewSheet) {
    // Auto-resize all columns used (A-C, E-G, I-K)
    termReasonsSheet.autoResizeColumns(OVERALL_COL_START, 3); // Columns A-C
    termReasonsSheet.autoResizeColumns(SITE_COL_START, 3);    // Columns E-G
    termReasonsSheet.autoResizeColumns(ELT_COL_START, 3);     // Columns I-K
  } else {
    // On existing sheets, preserve user formatting but ensure columns are visible
    // Only auto-resize if columns are too narrow (optional - can be removed if you want to preserve all formatting)
  }
  
  SpreadsheetApp.getUi().alert(`Termination Reasons table generated. ${sortedReasons.length} unique reasons found (${totalTerms} total terminations).`);
}

