/**
 * Bob Salary Data - Consolidated Google Apps Script
 * 
 * This script imports employee data, bonus history, and compensation history
 * from Bob (hibob.com) API into Google Sheets.
 */

// ============================================================================
// CONSTANTS
// ============================================================================

const BOB_REPORT_IDS = {
  BASE_DATA: "31048356",
  BONUS_HISTORY: "31054302",
  COMP_HISTORY: "31054312",
  FULL_COMP_HISTORY: "31168524"
};

const SHEET_NAMES = {
  BASE_DATA: "Base Data",
  BONUS_HISTORY: "Bonus History",
  COMP_HISTORY: "Comp History",
  FULL_COMP_HISTORY: "Full Comp History"
};

const WRITE_COLS = 23; // Column W - limit for Base Data sheet

const ALLOWED_EMP_TYPES = new Set(["Permanent", "Regular Full-Time"]);

// ============================================================================
// UI FUNCTIONS
// ============================================================================

/**
 * Creates a custom menu when the spreadsheet is opened
 * Allows easy access to all import functions
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Bob Salary Data')
    .addItem('Import Base Data', 'importBobDataSimpleWithLookup')
    .addItem('Import Bonus History', 'importBobBonusHistoryLatest')
    .addItem('Import Compensation History', 'importBobCompHistoryLatest')
    .addItem('Import Full Comp History', 'importBobFullCompHistory')
    .addSeparator()
    .addItem('Import All Data', 'importAllBobData')
    .addSeparator()
    .addItem('Convert Tenure to Array Formula', 'convertTenureToArrayFormula')
    .addToUi();
}

/**
 * Runs all three import functions in sequence
 * Useful for initial setup or full refresh
 */
function importAllBobData() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Import All Data',
      'This will import Base Data, Bonus History, and Compensation History. This may take a few minutes. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    Logger.log('Starting full data import...');
    
    // Import in sequence
    Logger.log('1/3: Importing Base Data...');
    importBobDataSimpleWithLookup();
    
    Logger.log('2/3: Importing Bonus History...');
    importBobBonusHistoryLatest();
    
    Logger.log('3/4: Importing Compensation History...');
    importBobCompHistoryLatest();
    
    Logger.log('4/4: Importing Full Comp History...');
    importBobFullCompHistory();
    
    Logger.log('All imports completed successfully!');
    ui.alert('Success', 'All data has been imported successfully!', ui.ButtonSet.OK);
  } catch (error) {
    Logger.log(`Error in importAllBobData: ${error.message}`);
    SpreadsheetApp.getUi().alert('Error', `Failed to import all data: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    throw error;
  }
}

/**
 * Converts individual tenure formulas to a single array formula for better performance
 * Works on the active sheet or a specified sheet name
 * Looks for "Tenure" column (D) and "Start Date" column (C)
 */
function convertTenureToArrayFormula() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    // Find header row by looking for "Tenure" column
    let headerRow = 1;
    let tenureCol = -1;
    let startDateCol = -1;
    
    // Search for header row (check first 30 rows, first 20 columns)
    for (let r = 1; r <= 30; r++) {
      const row = sheet.getRange(r, 1, 1, 20).getValues()[0];
      const tenureIdx = row.findIndex(cell => 
        cell && String(cell).toLowerCase().includes("tenure")
      );
      const startDateIdx = row.findIndex(cell => 
        cell && (String(cell).toLowerCase().includes("start date") || 
                 String(cell).toLowerCase().includes("startdate"))
      );
      
      if (tenureIdx !== -1) {
        headerRow = r;
        tenureCol = tenureIdx + 1; // Convert to 1-based
        if (startDateIdx !== -1) {
          startDateCol = startDateIdx + 1; // Convert to 1-based
        }
        break;
      }
    }
    
    if (tenureCol === -1) {
      SpreadsheetApp.getUi().alert('Error', 'Could not find "Tenure" column in the active sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    if (startDateCol === -1) {
      // Try to find Start Date column by checking common positions (column C = 3)
      startDateCol = 3; // Default to column C
      Logger.log('Start Date column not found in header, using column C as default');
    }
    
    const startDateColLetter = columnToLetter(startDateCol);
    const tenureColLetter = columnToLetter(tenureCol);
    
    // Find the first data row and last data row
    const firstDataRow = headerRow + 1;
    const lastRow = sheet.getLastRow();
    
    if (lastRow < firstDataRow) {
      SpreadsheetApp.getUi().alert('Error', 'No data rows found below the header.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Clear existing formulas in the tenure column
    const tenureRange = sheet.getRange(firstDataRow, tenureCol, lastRow - firstDataRow + 1, 1);
    tenureRange.clearContent();
    
    // Create array formula with inline categorization logic
    // This matches the categorizeTenure function logic
    const arrayFormula = `=ARRAYFORMULA(IF(${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastRow}="", "", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastRow})>=1460, "4 Years+", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastRow})>=1095, "3 Years", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastRow})>=730, "2 Years", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastRow})>=545, "1.5 Years", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastRow})>=365, "1 Year", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastRow})>=180, "6 Months", "Less than 6 Months"))))))))`;
    
    // Set the array formula in the first data row
    sheet.getRange(firstDataRow, tenureCol).setFormula(arrayFormula);
    
    Logger.log(`Converted tenure formulas to array formula in column ${tenureColLetter} (rows ${firstDataRow}-${lastRow})`);
    SpreadsheetApp.getUi().alert('Success', `Converted tenure formulas to array formula in column ${tenureColLetter}!\n\nThis will be much faster than individual cell formulas.`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    Logger.log(`Error in convertTenureToArrayFormula: ${error.message}`);
    SpreadsheetApp.getUi().alert('Error', `Failed to convert tenure formulas: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    throw error;
  }
}

// ============================================================================
// MAIN FUNCTIONS
// ============================================================================

/**
 * Imports base employee data from Bob API with variable type and percentage lookups
 */
function importBobDataSimpleWithLookup() {
  try {
    const reportId = BOB_REPORT_IDS.BASE_DATA;
    const sheetName = SHEET_NAMES.BASE_DATA;
    const bonusSheetName = SHEET_NAMES.BONUS_HISTORY;
    
    Logger.log(`Starting import of ${sheetName}...`);
    
    // Fetch data from Bob API
    const rows = fetchBobReport(reportId);
    
    // Cache normalized header for performance
    const srcHeader = rows[0];
    const normalizedHeader = normalizeHeader(srcHeader);
    
    const idxEmpId       = findColCached(normalizedHeader, srcHeader, ["Employee ID", "Emp ID", "Employee Id"]);
    const idxJobLevel    = findColCached(normalizedHeader, srcHeader, ["Job Level", "Job level"]);
    const idxBasePay     = findColCached(normalizedHeader, srcHeader, ["Base Pay", "Base salary", "Base Salary"]);
    const idxEmpType     = findColCached(normalizedHeader, srcHeader, ["Employment Type", "Employment type"]);
    const idxStartDate   = findColCached(normalizedHeader, srcHeader, ["Start Date", "Start date", "Original start date", "Original Start Date"]);
    const idxTermination = findColCached(normalizedHeader, srcHeader, ["Termination Date", "Termination date"]);
    
    let header = srcHeader.slice();
    header = [...header, "Variable Type", "Variable %"];
    
    const out = [header];
  
    // Process rows
    for (let r = 1; r < rows.length; r++) {
      const src = rows[r];
      if (!src || !Array.isArray(src) || src.length === 0) continue;
      
      const row = src.slice();
      const empType = safeCell(row, idxEmpType);
      if (!ALLOWED_EMP_TYPES.has(empType)) continue;
      
      const empId  = safeCell(row, idxEmpId);
      const jobLvl = safeCell(row, idxJobLevel);
      if (!empId || !jobLvl) continue;
      
      // Ensure Employee ID is stored as text (not number) for consistent XLOOKUP matching
      // Convert to number first to remove any formatting, then back to string to ensure consistency
      const empIdNum = toNumberSafe(empId);
      if (isFinite(empIdNum)) {
        row[idxEmpId] = String(empIdNum); // Store as text string
      } else {
        row[idxEmpId] = empId.trim(); // Keep as text but trimmed
      }
      
      const basePayNum = toNumberSafe(safeCell(row, idxBasePay));
      if (!isFinite(basePayNum) || basePayNum === 0) continue;
      
      row[idxBasePay] = basePayNum;
      row.push("", ""); // Variable Type, Variable %
      out.push(row);
    }
    
    Logger.log(`Processed ${out.length - 1} rows for ${sheetName}`);
    
    // Write to sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    
    // Only touch A:W (1..23)
    const sliced = out.map(r => r.slice(0, WRITE_COLS));
    
    // Clear contents in A:W only (keep everything after W intact)
    const rowsToClear = Math.max(sheet.getMaxRows(), sliced.length);
    sheet.getRange(1, 1, rowsToClear, WRITE_COLS).clearContent();
    
    // Write A:W
    const dataRange = sheet.getRange(1, 1, sliced.length, sliced[0].length);
    dataRange.setValues(sliced);
    
    // Delete extra empty rows - check maxRows to ensure we delete all extra rows
    const maxRows = sheet.getMaxRows();
    const dataRows = sliced.length;
    if (maxRows > dataRows) {
      // Delete all rows after the data
      sheet.deleteRows(dataRows + 1, maxRows - dataRows);
    }
    
    // Apply default formatting: Roboto 10, auto-justified (auto-resize columns)
    dataRange.setFontFamily("Roboto");
    dataRange.setFontSize(10);
    sheet.autoResizeColumns(1, sliced[0].length);
    
    const numRows = sliced.length - 1;
    const firstDataRow = 2;
    
    if (numRows > 0) {
      // Format Employee ID column as plain text (number-plaintext) for XLOOKUP compatibility
      if (idxEmpId + 1 <= WRITE_COLS) {
        const empIdRange = sheet.getRange(firstDataRow, idxEmpId + 1, numRows, 1);
        empIdRange.setNumberFormat("@"); // Set as text format
        // Ensure values are stored as plain text by setting them again
        const empIdValues = empIdRange.getValues();
        empIdRange.setValues(empIdValues.map(row => [String(row[0])]));
      }
      
      // Base Pay formatting (if within A:W)
      if (idxBasePay + 1 <= WRITE_COLS) {
        sheet.getRange(firstDataRow, idxBasePay + 1, numRows, 1).setNumberFormat("#,##0.00");
      }
      
      // Validate bonus sheet exists before creating formulas
      const bonusSheet = ss.getSheetByName(bonusSheetName);
      if (!bonusSheet) {
        Logger.log(`Warning: ${bonusSheetName} sheet not found. Skipping XLOOKUP formulas.`);
      } else {
        // Fill XLOOKUPs only if the columns exist within A:W
        const vtCol = header.indexOf("Variable Type") + 1;
        const vpCol = header.indexOf("Variable %") + 1;
        
        if (vtCol > 0 && vtCol <= WRITE_COLS) {
          const empIdColLetter = columnToLetter(idxEmpId + 1);
          const vtFormulas = Array.from({ length: numRows }, (_, i) =>
            [`=XLOOKUP($${empIdColLetter}${firstDataRow + i}, '${bonusSheetName}'!A:A, '${bonusSheetName}'!D:D, "")`]
          );
          sheet.getRange(firstDataRow, vtCol, numRows, 1).setFormulas(vtFormulas);
        }
        
        if (vpCol > 0 && vpCol <= WRITE_COLS) {
          const empIdColLetter = columnToLetter(idxEmpId + 1);
          const vpFormulas = Array.from({ length: numRows }, (_, i) =>
            [`=XLOOKUP($${empIdColLetter}${firstDataRow + i}, '${bonusSheetName}'!A:A, '${bonusSheetName}'!E:E, "")`]
          );
          sheet.getRange(firstDataRow, vpCol, numRows, 1).setFormulas(vpFormulas);
        }
      }
      
      // Add array formula for Tenure Category if Start Date column exists
      if (idxStartDate >= 0 && idxStartDate < header.length) {
        const startDateColLetter = columnToLetter(idxStartDate + 1);
        const tenureCategoryHeaderIndex = header.indexOf("Tenure Category");
        
        // Determine which column to use for Tenure Category
        let tenureCol;
        if (tenureCategoryHeaderIndex !== -1) {
          // Column already exists in header
          tenureCol = tenureCategoryHeaderIndex + 1;
        } else {
          // Add "Tenure Category" column after the last column
          // Check if we can add it within A:W, otherwise add after W
          if (header.length < WRITE_COLS) {
            tenureCol = header.length + 1;
            sheet.getRange(1, tenureCol).setValue("Tenure Category");
          } else {
            // Add after column W
            tenureCol = WRITE_COLS + 1;
            sheet.getRange(1, tenureCol).setValue("Tenure Category");
          }
        }
        
        // Clear any existing formulas in the tenure category column (in case individual formulas exist)
        const lastRow = sheet.getLastRow();
        if (lastRow >= firstDataRow) {
          sheet.getRange(firstDataRow, tenureCol, lastRow - firstDataRow + 1, 1).clearContent();
        }
        
        // Create array formula for tenure categorization
        // Formula calculates days from Start Date to TODAY and categorizes
        // Uses exact range to match data rows for better performance
        const lastDataRow = firstDataRow + numRows - 1;
        const arrayFormula = `=ARRAYFORMULA(IF(${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastDataRow}="", "", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastDataRow})>=1460, "4 Years+", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastDataRow})>=1095, "3 Years", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastDataRow})>=730, "2 Years", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastDataRow})>=545, "1.5 Years", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastDataRow})>=365, "1 Year", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastDataRow})>=180, "6 Months", "Less than 6 Months"))))))))`;
        
        // Set the array formula in the first data row (it will automatically fill down to lastDataRow)
        sheet.getRange(firstDataRow, tenureCol).setFormula(arrayFormula);
        
        Logger.log(`Set array formula for Tenure Category in column ${columnToLetter(tenureCol)}`);
      }
    }
    
    Logger.log(`Successfully imported ${sheetName}`);
  } catch (error) {
    Logger.log(`Error in importBobDataSimpleWithLookup: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error importing Base Data: ${error.message}`);
    throw error;
  }
}

/**
 * Imports bonus history from Bob API, keeping only the latest entry per employee
 */
function importBobBonusHistoryLatest() {
  try {
    const reportId = BOB_REPORT_IDS.BONUS_HISTORY;
    const targetSheetName = SHEET_NAMES.BONUS_HISTORY;
    
    Logger.log(`Starting import of ${targetSheetName}...`);
    
    // Fetch data from Bob API
    const rows = fetchBobReport(reportId);
    
    // Cache normalized header for performance
    const header = rows[0];
    const normalizedHeader = normalizeHeader(header);
    
    const iEmpId   = findColCached(normalizedHeader, header, ["Employee ID", "Emp ID", "Employee Id"]);
    const iName    = findColCached(normalizedHeader, header, ["Display name", "Emp Name", "Display Name", "Name"]);
    const iEffDate = findColCached(normalizedHeader, header, ["Effective date", "Effective Date", "Effective"]);
    const iType    = findColCached(normalizedHeader, header, ["Variable type", "Variable Type", "Type"]);
    const iPct     = findColCached(normalizedHeader, header, ["Commission/Bonus %", "Variable %", "Commission %", "Bonus %"]);
    const iAmt     = findColCached(normalizedHeader, header, ["Amount", "Variable Amount", "Commission/Bonus Amount"]);
    const iCurr    = findColOptionalCached(normalizedHeader, header, [
    "Variable Amount currency", "Variable Amount Currency",
    "Amount currency", "Amount Currency",
    "Currency", "Currency code", "Currency Code"
  ]);
  
  // Keep latest row per Emp ID (accepts ISO with optional time)
  const latest = new Map(); // empId -> { row, effKey }
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.length === 0) continue;
    
    let empId = safeCell(row, iEmpId);
    // Ensure Employee ID is stored as text for consistent XLOOKUP matching
    const empIdNum = toNumberSafe(empId);
    if (isFinite(empIdNum)) {
      empId = String(empIdNum); // Store as text string
    } else {
      empId = empId.trim(); // Keep as text but trimmed
    }
    
    const effRaw = safeCell(row, iEffDate);
    const effKey = (effRaw.match(/^\d{4}-\d{2}-\d{2}/) || [])[0]; // first 10 chars if ISO
    
    if (!empId || !effKey) continue;
    
    const existing = latest.get(empId);
    if (!existing || effKey > existing.effKey) {
      latest.set(empId, { row, effKey });
    }
  }
  
  const outHeader = [
    "Employee ID", "Display name", "Effective date",
    "Variable type", "Commission/Bonus %", "Amount", "Amount currency"
  ];
  const out = [outHeader];
  
  latest.forEach(({ row, effKey }) => {
    let empId = safeCell(row, iEmpId);
    // Ensure Employee ID is stored as text for consistent XLOOKUP matching
    const empIdNum = toNumberSafe(empId);
    if (isFinite(empIdNum)) {
      empId = String(empIdNum); // Store as text string
    } else {
      empId = empId.trim(); // Keep as text but trimmed
    }
    const name  = safeCell(row, iName);
    const type  = safeCell(row, iType);
    const pctVal = toNumberSafe(safeCell(row, iPct));
    const amtVal = toNumberSafe(safeCell(row, iAmt));
    const curr   = iCurr === -1 ? "" : safeCell(row, iCurr);
    
    // No conversion: keep the ISO date string exactly as provided
    const eff = effKey;
    out.push([empId, name, eff, type, isFinite(pctVal) ? pctVal : "", isFinite(amtVal) ? amtVal : "", curr]);
  });
  
    // Write to sheet
    writeToSheet(targetSheetName, out, [
      { range: [2, 3], format: "@" }, // Effective date as text
      { range: [2, 5], format: "0.########" }, // Percent
      { range: [2, 6], format: "#,##0.00" } // Amount
    ]);
    
    Logger.log(`Successfully imported ${targetSheetName}`);
  } catch (error) {
    Logger.log(`Error in importBobBonusHistoryLatest: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error importing Bonus History: ${error.message}`);
    throw error;
  }
}

/**
 * Imports compensation history from Bob API, keeping only the latest entry per employee
 */
function importBobCompHistoryLatest() {
  try {
    const reportId = BOB_REPORT_IDS.COMP_HISTORY;
    const targetSheetName = SHEET_NAMES.COMP_HISTORY;
    
    Logger.log(`Starting import of ${targetSheetName}...`);
    
    // Fetch data from Bob API
    const rows = fetchBobReport(reportId);
    
    // Cache normalized header for performance
    const header = rows[0];
    const normalizedHeader = normalizeHeader(header);
    
    const iEmpId   = findColCached(normalizedHeader, header, ["Emp ID", "Employee ID", "Employee Id"]);
    const iName    = findColCached(normalizedHeader, header, ["Emp Name", "Display name", "Display Name", "Name"]);
    const iEffDate = findColCached(normalizedHeader, header, ["History effective date", "Effective date", "Effective Date", "Effective"]);
    const iBase    = findColCached(normalizedHeader, header, ["History base salary", "Base salary", "Base Salary", "Base pay", "Salary"]);
    const iCurr    = findColCached(normalizedHeader, header, ["History base salary currency", "Base salary currency", "Currency"]);
    const iReason  = findColCached(normalizedHeader, header, ["History reason", "Reason", "Change reason"]);
  
  // Keep latest row per Emp ID by Effective date
  const latest = new Map(); // empId -> { row, ymd }
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.length === 0) continue;
    
    let empId = safeCell(row, iEmpId);
    // Ensure Employee ID is stored as text for consistent XLOOKUP matching
    const empIdNum = toNumberSafe(empId);
    if (isFinite(empIdNum)) {
      empId = String(empIdNum); // Store as text string
    } else {
      empId = empId.trim(); // Keep as text but trimmed
    }
    
    const effStr = safeCell(row, iEffDate);
    const ymd = toYmd(effStr);
    if (!empId || !ymd) continue;
    
    const existing = latest.get(empId);
    if (!existing || ymd > existing.ymd) latest.set(empId, { row, ymd });
  }
  
  // Build output (includes currency)
  const outHeader = ["Emp ID", "Emp Name", "Effective date", "Base salary", "Base salary currency", "History reason"];
  const out = [outHeader];
  
  latest.forEach(({ row, ymd }) => {
    let empId = safeCell(row, iEmpId);
    // Ensure Employee ID is stored as text for consistent XLOOKUP matching
    const empIdNum = toNumberSafe(empId);
    if (isFinite(empIdNum)) {
      empId = String(empIdNum); // Store as text string
    } else {
      empId = empId.trim(); // Keep as text but trimmed
    }
    const name   = safeCell(row, iName);
    const base   = toNumberSafe(safeCell(row, iBase)); // numeric
    const curr   = safeCell(row, iCurr);               // text
    const reason = safeCell(row, iReason);
    const effDate = parseDateSmart(ymd);
    out.push([empId, name, effDate, isFinite(base) ? base : "", curr, reason]);
  });
  
    // Write to sheet & format
    writeToSheet(targetSheetName, out, [
      { range: [2, 3], format: "yyyy-mm-dd" }, // Effective date
      { range: [2, 4], format: "#,##0.00" } // Base salary
      // Column 5 (currency) left as plain text
    ]);
    
    Logger.log(`Successfully imported ${targetSheetName}`);
  } catch (error) {
    Logger.log(`Error in importBobCompHistoryLatest: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error importing Comp History: ${error.message}`);
    throw error;
  }
}

/**
 * Imports full compensation history from Bob API (all time) and extracts:
 * - Last promotion date per employee
 * - Last increase % (calculated from last 2 compensation changes)
 * - Change reason for the last increase
 * Returns one row per employee
 */
function importBobFullCompHistory() {
  try {
    const reportId = BOB_REPORT_IDS.FULL_COMP_HISTORY;
    const targetSheetName = SHEET_NAMES.FULL_COMP_HISTORY;
    
    Logger.log(`Starting import of ${targetSheetName}...`);
    
    // Fetch data from Bob API
    const rows = fetchBobReport(reportId);
    
    // Cache normalized header for performance
    const header = rows[0];
    const normalizedHeader = normalizeHeader(header);
    
    const iEmpId   = findColCached(normalizedHeader, header, ["Emp ID", "Employee ID", "Employee Id"]);
    const iName    = findColCached(normalizedHeader, header, ["Emp Name", "Display name", "Display Name", "Name"]);
    const iEffDate = findColCached(normalizedHeader, header, ["History effective date", "Effective date", "Effective Date", "Effective"]);
    const iBase    = findColCached(normalizedHeader, header, ["History base salary", "Base salary", "Base Salary", "Base pay", "Salary"]);
    const iCurr    = findColCached(normalizedHeader, header, ["History base salary currency", "Base salary currency", "Currency"]);
    const iReason  = findColCached(normalizedHeader, header, ["History reason", "Reason", "Change reason"]);
    
    // Track all employees and their compensation history
    const employees = new Map(); // empId -> { name }
    const compHistory = new Map(); // empId -> Array of { row, ymd, base, currency, reason } sorted by date
    const promotions = new Map(); // empId -> { row, ymd } for promotion entries only
    
    // First pass: Collect all compensation history entries
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length === 0) continue;
      
      let empId = safeCell(row, iEmpId);
      const empIdNum = toNumberSafe(empId);
      if (isFinite(empIdNum)) {
        empId = String(empIdNum);
      } else {
        empId = empId.trim();
      }
      
      if (!empId) continue;
      
      // Track all employees (for names and to ensure we have one row per employee)
      const name = safeCell(row, iName);
      if (!employees.has(empId)) {
        employees.set(empId, { name: name || "" });
      }
      
      // Get compensation data
      const effStr = safeCell(row, iEffDate);
      const ymd = toYmd(effStr);
      const base = toNumberSafe(safeCell(row, iBase));
      const currency = safeCell(row, iCurr);
      const reason = safeCell(row, iReason);
      
      if (!ymd) continue;
      
      // Store compensation history entry
      if (!compHistory.has(empId)) {
        compHistory.set(empId, []);
      }
      compHistory.get(empId).push({ row, ymd, base, currency, reason });
      
      // Check if this is a promotion entry
      const reasonLower = reason.toLowerCase();
      if (reasonLower.includes("promotion")) {
        const existing = promotions.get(empId);
        if (!existing || ymd > existing.ymd) {
          promotions.set(empId, { row, ymd });
        }
      }
    }
    
    // Sort compensation history by date (most recent first) for each employee
    compHistory.forEach((entries, empId) => {
      entries.sort((a, b) => {
        if (b.ymd > a.ymd) return 1;
        if (b.ymd < a.ymd) return -1;
        return 0;
      });
    });
    
    // Build output: One row per employee with last promotion date, last increase %, and change reason
    const outHeader = ["Emp ID", "Emp Name", "Last Promotion Date", "Last Increase %", "Change Reason"];
    const out = [outHeader];
    
    employees.forEach((employee, empId) => {
      // Get last promotion date
      const promotion = promotions.get(empId);
      let promotionDate = "";
      
      if (promotion) {
        const effStr = safeCell(promotion.row, iEffDate);
        const ymd = toYmd(effStr);
        if (ymd) {
          promotionDate = parseDateSmart(ymd);
        }
      }
      
      // Calculate last increase % from last 2 compensation changes
      let lastIncreasePct = "";
      let changeReason = "";
      
      const history = compHistory.get(empId);
      if (history && history.length >= 2) {
        // Get the last 2 entries (most recent first)
        const lastEntry = history[0];
        const previousEntry = history[1];
        
        const lastBase = lastEntry.base;
        const prevBase = previousEntry.base;
        const lastCurrency = (lastEntry.currency || "").trim();
        const prevCurrency = (previousEntry.currency || "").trim();
        
        // Only calculate if both have valid base salaries AND same currency
        // Skip calculation if currencies differ (e.g., relocation between countries)
        if (isFinite(lastBase) && isFinite(prevBase) && prevBase > 0 && 
            lastCurrency && prevCurrency && lastCurrency === prevCurrency) {
          const increasePct = ((lastBase - prevBase) / prevBase) * 100;
          // Only report positive increases (exclude negative changes like employment status changes or corrections)
          if (increasePct >= 0) {
            lastIncreasePct = increasePct.toFixed(2);
            changeReason = lastEntry.reason || "";
          }
          // If negative (decrease), leave blank
        }
        // If currencies differ or missing, leave blank (handles relocation cases)
      }
      // If employee has less than 2 entries (new hire/ineligible), leave blank
      
      out.push([empId, employee.name, promotionDate, lastIncreasePct, changeReason]);
    });
    
    // Write to sheet & format
    writeToSheet(targetSheetName, out, [
      { range: [2, 3], format: "yyyy-mm-dd" }, // Last Promotion Date
      { range: [2, 4], format: "0.00" } // Last Increase % as percentage
    ]);
    
    Logger.log(`Successfully imported ${targetSheetName} - ${out.length - 1} employees, ${promotions.size} with promotion dates`);
  } catch (error) {
    Logger.log(`Error in importBobFullCompHistory: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error importing Full Comp History: ${error.message}`);
    throw error;
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Fetches a report from Bob API and returns parsed CSV rows
 * @param {string} reportId - Bob report ID
 * @returns {Array<Array>} Parsed CSV rows
 * @throws {Error} If fetch fails or data is invalid
 */
function fetchBobReport(reportId) {
  const apiUrl = `https://api.hibob.com/v1/company/reports/${reportId}/download?format=csv&locale=en-CA`;
  const creds = getBobCredentials();
  const basicAuth = Utilities.base64Encode(`${creds.id}:${creds.key}`);
  
  const res = UrlFetchApp.fetch(apiUrl, {
    method: "get",
    headers: { 
      accept: "text/csv", // Fixed: use text/csv for CSV endpoints
      authorization: `Basic ${basicAuth}` 
    },
    muteHttpExceptions: true
  });
  
  if (res.getResponseCode() !== 200) {
    throw new Error(`Failed to fetch CSV: ${res.getResponseCode()} - ${res.getContentText()}`);
  }
  
  const rows = Utilities.parseCsv(res.getContentText());
  if (!rows || rows.length === 0) {
    throw new Error("CSV contained no data.");
  }
  
  return rows;
}

/**
 * Gets Bob API credentials from Script Properties
 * @returns {{id: string, key: string}} Object containing BOB_ID and BOB_KEY
 * @throws {Error} If credentials are missing
 */
function getBobCredentials() {
  const id = PropertiesService.getScriptProperties().getProperty("BOB_ID");
  const key = PropertiesService.getScriptProperties().getProperty("BOB_KEY");
  
  if (!id || !key) {
    throw new Error("Missing BOB_ID or BOB_KEY in Script Properties. Please set these in Project Settings → Script Properties.");
  }
  
  return { id, key };
}

/**
 * Normalizes a header row for efficient searching
 * @param {Array<string>} headerRow - Array of header strings
 * @returns {Array<string>} Normalized header array
 */
function normalizeHeader(headerRow) {
  const norm = (s) => String(s || "").toLowerCase().replace(/\s+/g, " ").trim();
  return headerRow.map(norm);
}

/**
 * Finds a column index using cached normalized header
 * @param {Array<string>} normalizedHeader - Pre-normalized header array
 * @param {Array<string>} originalHeader - Original header for error messages
 * @param {Array<string>} aliases - Array of possible column names to match
 * @returns {number} Column index (0-based)
 * @throws {Error} If no matching column is found
 */
function findColCached(normalizedHeader, originalHeader, aliases) {
  const norm = (s) => String(s || "").toLowerCase().replace(/\s+/g, " ").trim();
  
  for (const alias of aliases) {
    const i = normalizedHeader.indexOf(norm(alias));
    if (i !== -1) return i;
  }
  
  throw new Error(`Could not find any of the columns [${aliases.join(", ")}]. Available headers: ${originalHeader.join(" | ")}`);
}

/**
 * Finds a column index optionally using cached normalized header
 * @param {Array<string>} normalizedHeader - Pre-normalized header array
 * @param {Array<string>} originalHeader - Original header (unused but kept for consistency)
 * @param {Array<string>} aliases - Array of possible column names to match
 * @returns {number} Column index (0-based) or -1 if not found
 */
function findColOptionalCached(normalizedHeader, originalHeader, aliases) {
  const norm = (s) => String(s || "").toLowerCase().replace(/\s+/g, " ").trim();
  
  for (const alias of aliases) {
    const i = normalizedHeader.indexOf(norm(alias));
    if (i !== -1) return i;
  }
  
  return -1;
}

/**
 * Writes data to a sheet with optional formatting
 * @param {string} sheetName - Name of the sheet
 * @param {Array<Array>} data - 2D array of data to write
 * @param {Array<Object>} formats - Array of format objects: {range: [row, col], format: string}
 */
function writeToSheet(sheetName, data, formats = []) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  
  sheet.clearContents();
  const dataRange = sheet.getRange(1, 1, data.length, data[0].length);
  dataRange.setValues(data);
  
  // Delete extra empty rows - use getMaxRows() to ensure we delete all extra rows
  const maxRows = sheet.getMaxRows();
  const dataRows = data.length;
  if (maxRows > dataRows) {
    // Delete all rows after the data
    sheet.deleteRows(dataRows + 1, maxRows - dataRows);
  }
  
  // Apply default formatting: Roboto 10, auto-justified (auto-resize columns)
  dataRange.setFontFamily("Roboto");
  dataRange.setFontSize(10);
  sheet.autoResizeColumns(1, data[0].length);
  
  // Format Employee ID column as plain text (number-plaintext) for XLOOKUP compatibility
  const header = data[0];
  const empIdHeaderIndex = header.findIndex(h => 
    h && (h.toString().toLowerCase().includes("employee id") || 
         h.toString().toLowerCase().includes("emp id"))
  );
  if (empIdHeaderIndex >= 0 && data.length > 1) {
    const empIdCol = empIdHeaderIndex + 1; // Convert to 1-based
    const empIdRange = sheet.getRange(2, empIdCol, data.length - 1, 1);
    empIdRange.setNumberFormat("@"); // Set as text format
    // Ensure values are stored as plain text by setting them again as strings
    const empIdValues = empIdRange.getValues();
    empIdRange.setValues(empIdValues.map(row => [String(row[0])]));
  }
  
  const numRows = data.length - 1;
  if (numRows > 0 && formats.length > 0) {
    formats.forEach(({ range, format }) => {
      if (format) {
        sheet.getRange(range[0], range[1], numRows, 1).setNumberFormat(format);
      }
    });
  }
}

/**
 * Safely extracts a cell value from a row
 * @param {Array} row - Array representing a row of data
 * @param {number} idx - Column index (0-based)
 * @returns {string} Trimmed string value or empty string
 */
function safeCell(row, idx) {
  return idx === -1 ? "" : (row[idx] == null ? "" : String(row[idx]).trim());
}

/**
 * Safely converts a value to a number, removing non-numeric characters
 * @param {*} val - Value to convert
 * @returns {number} Numeric value or NaN
 */
function toNumberSafe(val) {
  if (val == null || val === "") return NaN;
  return Number(String(val).replace(/[^\d.-]/g, ""));
}

/**
 * Converts a column number to letter notation (1=A, 2=B, 27=AA, etc.)
 * @param {number} col - Column number (1-based)
 * @returns {string} Column letter(s)
 */
function columnToLetter(col) {
  let letter = "";
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

/**
 * Parses a date string into a Date object, handling multiple formats
 * @param {string} s - Date string
 * @returns {Date} Date object
 */
function parseDateSmart(s) {
  if (!s) return s;
  
  let m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  
  m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(s);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  
  return new Date(s);
}

/**
 * Extracts YYYY-MM-DD format from a date string
 * @param {string} s - Date string
 * @returns {string} YYYY-MM-DD format or empty string
 */
function toYmd(s) {
  if (!s) return "";
  
  let m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  
  m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(s);
  if (m) return `${m[3]}-${m[2]}-${m[1]}`;
  
  return "";
}

/**
 * Categorizes employee tenure based on days
 * @param {number} days - Number of days of tenure
 * @returns {string} Tenure category (e.g., "6 Months", "1 Year", "4 Years+")
 */
function categorizeTenure(days) {
  if (!days || days < 0 || !isFinite(days)) return "";
  
  if (days >= 1460) {
    return "4 Years+";
  } else if (days >= 1095) {
    return "3 Years";
  } else if (days >= 730) {
    return "2 Years";
  } else if (days >= 545) {
    return "1.5 Years";
  } else if (days >= 365) {
    return "1 Year";
  } else if (days >= 180) {
    return "6 Months";
  } else {
    return "Less than 6 Months";
  }
}

