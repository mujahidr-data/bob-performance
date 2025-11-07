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
    .addItem('Map Termination Reasons', 'mapTerminationReasons')
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
    .addItem('Generate Job Level Headcount', 'generateJobLevelHeadcount')
    .addItem('Generate Termination Reasons Table', 'generateTerminationReasonsTable')
    .addToUi();
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
  
  // Apply filters
  let filteredRows = rows;
  if (filters.site) {
    filteredRows = filteredRows.filter(row => row[COLUMN_INDICES.SITE] === filters.site);
  }
  if (filters.elt) {
    filteredRows = filteredRows.filter(row => row[COLUMN_INDICES.ELT] === filters.elt);
  }
  if (filters.department) {
    filteredRows = filteredRows.filter(row => row[COLUMN_INDICES.DEPARTMENT] === filters.department);
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
  const termNames = new Set();
  filteredRows.forEach(row => {
    const empName = getEmpName(row);
    if (!empName) return; // Skip rows without employee name
    
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    if (!termDate) return;
    if (termDate >= periodStartDate && termDate <= periodEndDate) {
      termNames.add(empName);
    }
  });
  const terms = termNames.size;
  
  // Voluntary Terms: Original formula = COUNTIFS with multiple conditions
  // Employees who terminated in period AND have voluntary termination type
  // Count unique employee names
  const voluntaryTermNames = new Set();
  filteredRows.forEach(row => {
    const empName = getEmpName(row);
    if (!empName) return; // Skip rows without employee name
    
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    if (!termDate) return;
    if (termDate < periodStartDate || termDate > periodEndDate) return;
    
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
  const involuntaryTermNames = new Set();
  filteredRows.forEach(row => {
    const empName = getEmpName(row);
    if (!empName) return; // Skip rows without employee name
    
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    if (!termDate) return;
    if (termDate < periodStartDate || termDate > periodEndDate) return;
    
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
  const regrettableTermNames = new Set();
  filteredRows.forEach(row => {
    const empName = getEmpName(row);
    if (!empName) return; // Skip rows without employee name
    
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    if (!termDate) return;
    if (termDate < periodStartDate || termDate > periodEndDate) return;
    
    const termCategory = String(row[COLUMN_INDICES.TERMINATION_CATEGORY] || "").trim();
    const termType = String(row[COLUMN_INDICES.LEAVE_TERMINATION_TYPE] || "").trim();
    
    let isRegrettable = false;
    
    if (termCategory) {
      // Use Column U value if available
      isRegrettable = termCategory === "Regretted" || termCategory.toLowerCase() === "regretted";
    } else {
      // Fallback to Column T logic if Column U is blank
      if (termType === "Voluntary - Regrettable") {
        isRegrettable = true;
      } else if (termType === "Involuntary" || termType === "Involuntary - Regrettable" || 
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
  
  // Write headers
  filterSheet.getRange(1, 1, 1, 4).setValues([["Site", "ELT", "Department", "Termination Reason"]]);
  filterSheet.getRange(1, 1, 1, 4).setFontWeight("bold");
  
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
    filterSheet.getRange(2, 1, maxLength, 4).setValues(values);
  }
  
  // Auto-resize columns
  filterSheet.autoResizeColumns(1, 4);
  
  // Add instructions
  filterSheet.getRange(maxLength + 3, 1).setValue("Instructions:");
  filterSheet.getRange(maxLength + 4, 1).setValue("1. Review the unique values above");
  filterSheet.getRange(maxLength + 5, 1).setValue("2. Use generateOverallData() function with filters parameter");
  filterSheet.getRange(maxLength + 6, 1).setValue("3. Example: generateOverallData(new Date('2024-01-01'), new Date('2024-01-31'), {site: 'India'})");
  
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
  
  // Get filter selections from a "FilterConfig" sheet if it exists
  let filters = {};
  const filterConfigSheet = ss.getSheetByName("FilterConfig");
  if (filterConfigSheet) {
    const siteCell = filterConfigSheet.getRange("B2"); // Row 2 for Site
    const eltCell = filterConfigSheet.getRange("B3");  // Row 3 for ELT
    const deptCell = filterConfigSheet.getRange("B4"); // Row 4 for Department
    
    const site = siteCell.getValue();
    const elt = eltCell.getValue();
    const dept = deptCell.getValue();
    
    if (site) filters.site = site;
    if (elt) filters.elt = elt;
    if (dept) filters.department = dept;
  }
  
  // Generate periods (monthly from Jan 2024 to current month)
  const periods = [];
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
      periods.push({
        start: periodStart,
        end: periodEnd,
        label: Utilities.formatDate(periodStart, Session.getScriptTimeZone(), "MMM yyyy")
      });
    }
    
    currentDate = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 1);
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
      const metrics = calculateHRMetrics(periods[i].start, periods[i].end, filters);
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
    // Update data values (preserves existing formatting)
    metricsSheet.getRange(2, 1, values.length, headers.length).setValues(values);
    
    // If there are more existing data rows than new data, clear the extra rows (but preserve formatting)
    const existingLastRow = metricsSheet.getLastRow();
    if (existingLastRow > values.length + 1) {
      // Clear content only (not formatting) for rows beyond the new data
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
  
  // Add filter info if filters are applied
  if (Object.keys(filters).length > 0) {
    const filterInfo = ["Filters Applied:"];
    if (filters.site) filterInfo.push(`Site: ${filters.site}`);
    if (filters.elt) filterInfo.push(`ELT: ${filters.elt}`);
    if (filters.department) filterInfo.push(`Department: ${filters.department}`);
    metricsSheet.getRange(lastRow + 2, 1, filterInfo.length, 1).setValues(filterInfo.map(f => [f]));
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
  
  // Write headers and labels
  configSheet.getRange(1, 1, 1, 2).setValues([["Filter", "Value"]]);
  configSheet.getRange(1, 1, 1, 2).setFontWeight("bold");
  
  configSheet.getRange(2, 1).setValue("Site:");
  configSheet.getRange(3, 1).setValue("ELT:");
  configSheet.getRange(4, 1).setValue("Department:");
  
  // Add data validation dropdowns
  if (uniqueValues.sites.length > 0) {
    const siteRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([""].concat(uniqueValues.sites), true)
      .build();
    configSheet.getRange(2, 2).setDataValidation(siteRule);
  }
  
  if (uniqueValues.elts.length > 0) {
    const eltRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([""].concat(uniqueValues.elts), true)
      .build();
    configSheet.getRange(3, 2).setDataValidation(eltRule);
  }
  
  if (uniqueValues.departments.length > 0) {
    const deptRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([""].concat(uniqueValues.departments), true)
      .build();
    configSheet.getRange(4, 2).setDataValidation(deptRule);
  }
  
  // Add instructions
  configSheet.getRange(6, 1).setValue("Instructions:");
  configSheet.getRange(7, 1).setValue("1. Select filter values from dropdowns above (leave blank for all)");
  configSheet.getRange(8, 1).setValue("2. Run 'Generate Overall Data' from the menu to apply filters");
  configSheet.getRange(9, 1).setValue("3. Metrics will be calculated based on selected filters");
  
  configSheet.autoResizeColumns(1, 2);
  
  SpreadsheetApp.getUi().alert("FilterConfig sheet created. Select your filters and run 'Generate Overall Data'.");
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
  const periods = [];
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
      periods.push({
        start: periodStart,
        end: periodEnd,
        label: Utilities.formatDate(periodStart, Session.getScriptTimeZone(), "MMM yyyy")
      });
    }
    
    currentDate = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 1);
  }
  
  // Build table: Site/Metric as rows, Periods as columns
  const headerRow = ["Site / Metric"].concat(periods.map(p => p.label));
  
  if (isNewSheet) {
    headcountSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    headcountSheet.getRange(1, 1, 1, headerRow.length).setFontWeight("bold");
  } else {
    // Update header row
    headcountSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  }
  
  // Define metrics to calculate for each site
  const metricLabels = [
    "Average Headcount",
    "Hires",
    "Terminations",
    "Regrettable Terms",
    "Non-Regrettable Terms",
    "Attrition %",
    "Regrettable %"
  ];
  
  // Calculate metrics for each site and period
  const dataRows = [];
  sites.forEach(site => {
    // For each site, create rows for each metric
    metricLabels.forEach((metricLabel, metricIndex) => {
      const row = [`${site} - ${metricLabel}`];
      
      periods.forEach(period => {
        try {
          const metrics = calculateHRMetrics(period.start, period.end, { site: site });
          
          let value = 0;
          switch(metricIndex) {
            case 0: // Average Headcount
              value = (metrics.openingHC + metrics.closingHC) / 2;
              break;
            case 1: // Hires
              value = metrics.hires;
              break;
            case 2: // Terminations
              value = metrics.terms;
              break;
            case 3: // Regrettable Terms
              value = metrics.regrettableTerms;
              break;
            case 4: // Non-Regrettable Terms
              value = metrics.terms - metrics.regrettableTerms;
              break;
            case 5: // Attrition %
              value = metrics.attrition;
              break;
            case 6: // Regrettable %
              const avgHC = (metrics.openingHC + metrics.closingHC) / 2;
              value = avgHC > 0 ? (metrics.regrettableTerms / avgHC) : 0;
              break;
          }
          row.push(value);
        } catch (error) {
          Logger.log(`Error calculating ${metricLabel} for ${site} in ${period.label}: ${error.message}`);
          row.push(0);
        }
      });
      dataRows.push(row);
    });
  });
  
  // Write data
  if (dataRows.length > 0) {
    headcountSheet.getRange(2, 1, dataRows.length, headerRow.length).setValues(dataRows);
    
    // Format percentage columns (Attrition % and Regrettable %)
    // These are at indices 5 and 6 in metricLabels array
    if (isNewSheet && dataRows.length > 0) {
      try {
        // Format percentage rows: Attrition % (index 5) and Regrettable % (index 6)
        // For each site, format rows at positions: baseRow + 5 and baseRow + 6
        sites.forEach((site, siteIndex) => {
          const baseRow = siteIndex * metricLabels.length + 2; // +2 because data starts at row 2
          const attritionRow = baseRow + 5; // Attrition % is at index 5
          const regrettableRow = baseRow + 6; // Regrettable % is at index 6
          
          if (attritionRow <= dataRows.length + 1) {
            headcountSheet.getRange(attritionRow, 2, 1, periods.length).setNumberFormat("0.0%");
          }
          if (regrettableRow <= dataRows.length + 1) {
            headcountSheet.getRange(regrettableRow, 2, 1, periods.length).setNumberFormat("0.0%");
          }
        });
      } catch (error) {
        Logger.log(`Error formatting percentage columns: ${error.message}`);
        // Continue execution even if formatting fails
      }
    }
  }
  
  // Only auto-resize if new sheet
  if (isNewSheet) {
    headcountSheet.autoResizeColumns(1, headerRow.length);
  }
  
  SpreadsheetApp.getUi().alert(`Headcount Metrics generated for ${sites.length} sites with ${metricLabels.length} metrics each, across ${periods.length} periods.`);
}

/**
 * Generates job level headcount as of today
 * Creates a table with job levels sorted in specific order
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
  
  // Count unique employees by job level (as of today)
  const jobLevelCounts = new Map();
  
  rows.forEach(row => {
    const empName = String(row[COLUMN_INDICES.EMP_NAME] || "").trim();
    if (!empName) return;
    
    const startDate = parseDate(row[COLUMN_INDICES.START_DATE]);
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    const jobLevel = String(row[COLUMN_INDICES.JOB_LEVEL] || "").trim();
    
    // Check if employee is active as of today
    if (!startDate || startDate > today) return; // Not started yet
    if (termDate && termDate <= today) return; // Already terminated
    
    // Count unique employees by job level
    if (!jobLevelCounts.has(jobLevel)) {
      jobLevelCounts.set(jobLevel, new Set());
    }
    jobLevelCounts.get(jobLevel).add(empName);
  });
  
  // Sort job levels according to specified order
  const sortedJobLevels = Array.from(jobLevelCounts.keys()).sort((a, b) => {
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
  
  // Build table
  const headers = ["Job Level", "Headcount"];
  if (isNewSheet) {
    jobLevelSheet.getRange(1, 1, 1, 2).setValues([headers]);
    jobLevelSheet.getRange(1, 1, 1, 2).setFontWeight("bold");
  } else {
    jobLevelSheet.getRange(1, 1, 1, 2).setValues([headers]);
  }
  
  const dataRows = sortedJobLevels.map(level => [
    level,
    jobLevelCounts.get(level).size
  ]);
  
  if (dataRows.length > 0) {
    jobLevelSheet.getRange(2, 1, dataRows.length, 2).setValues(dataRows);
  }
  
  if (isNewSheet) {
    jobLevelSheet.autoResizeColumns(1, 2);
  }
  
  SpreadsheetApp.getUi().alert(`Job Level Headcount generated for ${sortedJobLevels.length} job levels as of today.`);
}

/**
 * Maps/cleans termination reasons based on user-provided mapping
 * This function will be updated when user provides the mapping
 */
function mapTerminationReasons() {
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
  
  // TODO: Add termination reason mapping when user provides it
  // Example structure:
  // const reasonMapping = {
  //   "Old Reason 1": "New Reason 1",
  //   "Old Reason 2": "New Reason 2"
  // };
  
  const reasonMapping = {}; // Will be populated when user provides mapping
  
  let updateCount = 0;
  rows.forEach((row, index) => {
    const currentReason = String(row[COLUMN_INDICES.TERMINATION_REASON] || "").trim();
    if (currentReason && reasonMapping.hasOwnProperty(currentReason)) {
      row[COLUMN_INDICES.TERMINATION_REASON] = reasonMapping[currentReason];
      updateCount++;
    }
  });
  
  if (updateCount > 0) {
    const updatedData = [header, ...rows];
    const targetRange = rawDataSheet.getRange(1, 1, updatedData.length, updatedData[0].length);
    targetRange.setValues(updatedData);
    SpreadsheetApp.getUi().alert(`Termination reasons mapped. ${updateCount} rows updated.`);
  } else {
    SpreadsheetApp.getUi().alert("No termination reason mappings defined yet. Please add mappings to the function.");
  }
}

/**
 * Generates termination reasons table for pie chart
 * Filterable by time period (month/year, quarter, year)
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
  
  // Count terminations by reason (all time, can be filtered later)
  const reasonCounts = new Map();
  
  rows.forEach(row => {
    const empName = String(row[COLUMN_INDICES.EMP_NAME] || "").trim();
    if (!empName) return;
    
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    const termReason = String(row[COLUMN_INDICES.TERMINATION_REASON] || "").trim();
    
    if (!termDate || !termReason) return; // Skip if no termination date or reason
    
    // Count unique employees by termination reason
    if (!reasonCounts.has(termReason)) {
      reasonCounts.set(termReason, new Set());
    }
    reasonCounts.get(termReason).add(empName);
  });
  
  // Sort by count (descending)
  const sortedReasons = Array.from(reasonCounts.entries())
    .sort((a, b) => b[1].size - a[1].size);
  
  // Build table
  const headers = ["Termination Reason", "Count", "Percentage"];
  if (isNewSheet) {
    termReasonsSheet.getRange(1, 1, 1, 3).setValues([headers]);
    termReasonsSheet.getRange(1, 1, 1, 3).setFontWeight("bold");
  } else {
    termReasonsSheet.getRange(1, 1, 1, 3).setValues([headers]);
  }
  
  const totalTerms = Array.from(reasonCounts.values()).reduce((sum, set) => sum + set.size, 0);
  
  const dataRows = sortedReasons.map(([reason, empSet]) => {
    const count = empSet.size;
    const percentage = totalTerms > 0 ? count / totalTerms : 0;
    return [reason, count, percentage];
  });
  
  if (dataRows.length > 0) {
    termReasonsSheet.getRange(2, 1, dataRows.length, 3).setValues(dataRows);
    
    // Format percentage column
    if (isNewSheet) {
      termReasonsSheet.getRange(2, 3, dataRows.length, 1).setNumberFormat("0.0%");
    }
  }
  
  if (isNewSheet) {
    termReasonsSheet.autoResizeColumns(1, 3);
  }
  
  SpreadsheetApp.getUi().alert(`Termination Reasons table generated. ${sortedReasons.length} unique reasons found (${totalTerms} total terminations).`);
}

