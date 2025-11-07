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
    .addSeparator()
    // Step 3: Review available filters
    .addItem('Step 3: Update Filter Options', 'updateFilterOptions')
    // Step 4: Set up filter selection
    .addItem('Step 4: Create Filter Config Sheet', 'createFilterConfigSheet')
    .addSeparator()
    // Step 5: Generate metrics
    .addItem('Step 5: Generate HR Metrics', 'generateHRMetrics')
    .addToUi();
}

/**
 * Column indices mapping based on Bob import structure
 * Based on formulas: C:C (Start Date), O:O (Termination Date), Q:Q (Leave/Termination Type)
 * Adjust these if your column structure differs
 * Note: Google Sheets uses 1-based columns, but arrays are 0-indexed
 */
const COLUMN_INDICES = {
  START_DATE: 2,           // Column C (0-indexed: 2) - Start Date in YYYY-MM-DD format
  TERMINATION_DATE: 17,    // Column R (0-indexed: 17) - Termination date (blank for active employees)
  LEAVE_TERMINATION_TYPE: 19, // Column T (0-indexed: 19) - Leave and termination type
  DEPARTMENT: 6,           // Column G (0-indexed: 6)
  ELT: 7,                  // Column H (0-indexed: 7)
  SITE: 8,                 // Column I (0-indexed: 8)
  STATUS: 16               // Column Q (0-indexed: 16) - Status (Active/Inactive, Employed/Terminated)
};

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
  
  // Opening Headcount: Original formula = COUNTIFS(C:C, "<=" & periodStart, O:O, ">=" & periodStart) 
  //                    + COUNTIFS(C:C, "<=" & periodStart, O:O, "")
  // Employees who: Started <= periodStart AND (Terminated >= periodStart OR never terminated)
  const openingHC = filteredRows.filter(row => {
    const startDate = parseDate(row[COLUMN_INDICES.START_DATE]);
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    
    // Must have started on or before period start
    if (!startDate || startDate > periodStartDate) return false;
    
    // Either never terminated (active) OR terminated on/after period start
    if (!termDate) return true; // Active employee
    return termDate >= periodStartDate;
  }).length;
  
  // Closing Headcount: Original formula = COUNTIFS(C:C, "<=" & periodEnd, O:O, ">" & periodEnd)
  //                     + COUNTIFS(C:C, "<=" & periodEnd, O:O, "")
  // Employees who: Started <= periodEnd AND (Terminated > periodEnd OR never terminated)
  const closingHC = filteredRows.filter(row => {
    const startDate = parseDate(row[COLUMN_INDICES.START_DATE]);
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    
    // Must have started on or before period end
    if (!startDate || startDate > periodEndDate) return false;
    
    // Either never terminated (active) OR terminated after period end
    if (!termDate) return true; // Active employee
    return termDate > periodEndDate;
  }).length;
  
  // Hires: Original formula = COUNTIFS(C:C, ">=" & periodStart, C:C, "<=" & periodEnd)
  // Employees who started between periodStart and periodEnd (inclusive)
  const hires = filteredRows.filter(row => {
    const startDate = parseDate(row[COLUMN_INDICES.START_DATE]);
    if (!startDate) return false;
    return startDate >= periodStartDate && startDate <= periodEndDate;
  }).length;
  
  // Terms: Original formula = COUNTIFS(O:O, ">=" & periodStart, O:O, "<=" & periodEnd)
  // Employees who terminated between periodStart and periodEnd (inclusive)
  const terms = filteredRows.filter(row => {
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    if (!termDate) return false;
    return termDate >= periodStartDate && termDate <= periodEndDate;
  }).length;
  
  // Voluntary Terms: Original formula = COUNTIFS with multiple conditions
  // Employees who terminated in period AND have voluntary termination type
  const voluntaryTerms = filteredRows.filter(row => {
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    if (!termDate) return false;
    if (termDate < periodStartDate || termDate > periodEndDate) return false;
    
    const termType = String(row[COLUMN_INDICES.LEAVE_TERMINATION_TYPE] || "").trim();
    return termType === "Voluntary" || 
           termType === "Voluntary - Regrettable" || 
           termType === "Voluntary - Non regrettable";
  }).length;
  
  // Involuntary Terms: Original formula = COUNTIFS with multiple conditions
  // Employees who terminated in period AND have involuntary termination type
  const involuntaryTerms = filteredRows.filter(row => {
    const termDate = parseDate(row[COLUMN_INDICES.TERMINATION_DATE]);
    if (!termDate) return false;
    if (termDate < periodStartDate || termDate > periodEndDate) return false;
    
    const termType = String(row[COLUMN_INDICES.LEAVE_TERMINATION_TYPE] || "").trim();
    return termType === "Involuntary" || 
           termType === "Involuntary - Regrettable" || 
           termType === "End of Contract";
  }).length;
  
  // Calculate rates (return as decimals, sheet will format as percentage)
  // Original formulas: attrition = terms / avgHC, retention = (openingHC - terms) / openingHC, turnover = terms / avgHC
  const avgHC = (openingHC + closingHC) / 2;
  const attrition = avgHC > 0 ? (terms / avgHC) : 0;
  const retention = openingHC > 0 ? ((openingHC - terms) / openingHC) : 0;
  const turnover = avgHC > 0 ? (terms / avgHC) : 0;
  
  return {
    openingHC,
    closingHC,
    hires,
    terms,
    voluntaryTerms,
    involuntaryTerms,
    // Round to 4 decimal places for accurate percentage display (0.9680 = 96.80%)
    attrition: Math.round(attrition * 10000) / 10000,
    retention: Math.round(retention * 10000) / 10000,
    turnover: Math.round(turnover * 10000) / 10000
  };
}

/**
 * Gets unique values for filtering (Site, ELT, Department)
 * @returns {Object} Object with arrays of unique values
 */
function getUniqueFilterValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawDataSheet = ss.getSheetByName("RawData");
  
  if (!rawDataSheet) {
    return { sites: [], elts: [], departments: [] };
  }
  
  const data = rawDataSheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { sites: [], elts: [], departments: [] };
  }
  
  const rows = data.slice(1);
  const sites = new Set();
  const elts = new Set();
  const departments = new Set();
  
  rows.forEach(row => {
    const site = String(row[COLUMN_INDICES.SITE] || "").trim();
    const elt = String(row[COLUMN_INDICES.ELT] || "").trim();
    const dept = String(row[COLUMN_INDICES.DEPARTMENT] || "").trim();
    
    if (site) sites.add(site);
    if (elt) elts.add(elt);
    if (dept) departments.add(dept);
  });
  
  return {
    sites: Array.from(sites).sort(),
    elts: Array.from(elts).sort(),
    departments: Array.from(departments).sort()
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
  filterSheet.getRange(1, 1, 1, 3).setValues([["Site", "ELT", "Department"]]);
  filterSheet.getRange(1, 1, 1, 3).setFontWeight("bold");
  
  // Find max length
  const maxLength = Math.max(
    uniqueValues.sites.length,
    uniqueValues.elts.length,
    uniqueValues.departments.length
  );
  
  // Write data
  if (maxLength > 0) {
    const values = [];
    for (let i = 0; i < maxLength; i++) {
      values.push([
        uniqueValues.sites[i] || "",
        uniqueValues.elts[i] || "",
        uniqueValues.departments[i] || ""
      ]);
    }
    filterSheet.getRange(2, 1, maxLength, 3).setValues(values);
  }
  
  // Auto-resize columns
  filterSheet.autoResizeColumns(1, 3);
  
  // Add instructions
  filterSheet.getRange(maxLength + 3, 1).setValue("Instructions:");
  filterSheet.getRange(maxLength + 4, 1).setValue("1. Review the unique values above");
  filterSheet.getRange(maxLength + 5, 1).setValue("2. Use generateHRMetrics() function with filters parameter");
  filterSheet.getRange(maxLength + 6, 1).setValue("3. Example: generateHRMetrics(new Date('2024-01-01'), new Date('2024-01-31'), {site: 'India'})");
  
  SpreadsheetApp.getUi().alert(`Filter options updated. Found ${uniqueValues.sites.length} sites, ${uniqueValues.elts.length} ELTs, ${uniqueValues.departments.length} departments.`);
}

/**
 * Generates HR metrics for multiple periods and writes to "HR Metrics" sheet
 * Creates a comprehensive metrics table with all calculated values
 */
function generateHRMetrics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let metricsSheet = ss.getSheetByName("HR Metrics");
  
  if (!metricsSheet) {
    metricsSheet = ss.insertSheet("HR Metrics");
  } else {
    // Clear existing data but keep structure
    const lastRow = metricsSheet.getLastRow();
    if (lastRow > 1) {
      metricsSheet.getRange(2, 1, lastRow - 1, metricsSheet.getLastColumn()).clearContent();
    }
  }
  
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
  
  // Write headers
  const headers = [
    "Period",
    "Opening Headcount",
    "Closing Headcount",
    "Hires",
    "Terms",
    "Voluntary",
    "Involuntary",
    "Attrition %",
    "Retention %",
    "Turnover %"
  ];
  
  metricsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  metricsSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  
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
        metrics.attrition,
        metrics.retention,
        metrics.turnover
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
        ""
      ]);
    }
  }
  
  if (values.length > 0) {
    metricsSheet.getRange(2, 1, values.length, headers.length).setValues(values);
  }
  
  // Format percentage columns
  const lastRow = values.length + 1;
  metricsSheet.getRange(2, 8, lastRow - 1, 3).setNumberFormat("0.0%");
  
  // Auto-resize columns
  metricsSheet.autoResizeColumns(1, headers.length);
  
  // Add filter info if filters are applied
  if (Object.keys(filters).length > 0) {
    const filterInfo = ["Filters Applied:"];
    if (filters.site) filterInfo.push(`Site: ${filters.site}`);
    if (filters.elt) filterInfo.push(`ELT: ${filters.elt}`);
    if (filters.department) filterInfo.push(`Department: ${filters.department}`);
    metricsSheet.getRange(lastRow + 2, 1, filterInfo.length, 1).setValues(filterInfo.map(f => [f]));
  }
  
  SpreadsheetApp.getUi().alert(`HR Metrics generated for ${periods.length} periods.`);
}

/**
 * Creates a FilterConfig sheet for easy filter selection
 * Users can enter filter values in this sheet, then run generateHRMetrics()
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
  configSheet.getRange(8, 1).setValue("2. Run 'Generate HR Metrics' from the menu to apply filters");
  configSheet.getRange(9, 1).setValue("3. Metrics will be calculated based on selected filters");
  
  configSheet.autoResizeColumns(1, 2);
  
  SpreadsheetApp.getUi().alert("FilterConfig sheet created. Select your filters and run 'Generate HR Metrics'.");
}

