/**
 * Bob Performance Module - Google Apps Script
 * 
 * This module combines:
 * - Bob Salary Data imports (Base Data, Bonus History, Compensation History)
 * - HiBob Performance Report automation (credentials management, file upload handler)
 * 
 * All functionality is integrated into a single, optimized codebase.
 * 
 * @author MR
 * @version 2.0.0
 * @published 2025-11-10
 * @since 2024
 * 
 * Requirements:
 * - Bob API credentials (BOB_ID, BOB_KEY) in Script Properties
 * - Google Sheets API access
 * - Appropriate permissions for Bob API endpoints
 */

// ============================================================================
// CONSTANTS - Salary Data
// ============================================================================

const BOB_REPORT_IDS = {
  BASE_DATA: "31048356",
  BONUS_HISTORY: "31054302",
  COMP_HISTORY: "31054312",
  FULL_COMP_HISTORY: "31168524",
  PERF_RATINGS: "31172066"
};

const SHEET_NAMES = {
  BASE_DATA: "Base Data",
  BONUS_HISTORY: "Bonus History",
  COMP_HISTORY: "Comp History",
  FULL_COMP_HISTORY: "Full Comp History",
  PERF_RATINGS: "Performance Ratings",
  PERF_REPORT: "Bob Perf Report"
};

const WRITE_COLS = 23; // Column W - limit for Base Data sheet
const ALLOWED_EMP_TYPES = new Set(["Permanent", "Regular Full-Time"]);

// ============================================================================
// CONSTANTS - Performance Reports
// ============================================================================

const PERF_SHEET_ID = '1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA';

// ============================================================================
// UI FUNCTIONS - Menu
// ============================================================================

/**
 * Creates unified menu when spreadsheet is opened
 * Combines Salary Data and Performance Report functions
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Bob Performance Module')
    .addItem('📊 Import Base Data', 'importBobDataSimpleWithLookup')
    .addItem('💰 Import Bonus History', 'importBobBonusHistoryLatest')
    .addItem('📈 Import Compensation History', 'importBobCompHistoryLatest')
    .addItem('📊 Import Full Comp History', 'importBobFullCompHistory')
    .addItem('⭐ Import Performance Ratings', 'importBobPerformanceRatings')
    .addSeparator()
    .addItem('📥 Import All Data', 'importAllBobData')
    .addSeparator()
    .addItem('🔧 Build Summary Sheet', 'buildSummarySheet')
    .addItem('🔄 Update Slicers Only', 'updateSlicersOnly')
    .addItem('🤖 Generate Manager Blurbs', 'generateManagerBlurbs')
    .addItem('🎯 Create AI Readiness Mapping', 'createAIReadinessMapping')
    .addSeparator()
    .addItem('🌐 Launch Web Interface', 'launchWebInterface')
    .addItem('❓ Instructions & Help', 'showInstructions')
    .addToUi();
}

// ============================================================================
// SALARY DATA FUNCTIONS
// ============================================================================

/**
 * Runs all import functions in sequence
 */
function importAllBobData() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Import All Data',
      'This will import Base Data, Bonus History, Compensation History, Full Comp History, and Performance Ratings. This may take a few minutes. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) return;
    
    Logger.log('Starting full data import...');
    importBobDataSimpleWithLookup();
    Logger.log('1/5: Base Data imported');
    
    importBobBonusHistoryLatest();
    Logger.log('2/5: Bonus History imported');
    
    importBobCompHistoryLatest();
    Logger.log('3/5: Compensation History imported');
    
    importBobFullCompHistory();
    Logger.log('4/5: Full Comp History imported');
    
    importBobPerformanceRatings();
    Logger.log('5/5: Performance Ratings imported');
    
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
 */
function convertTenureToArrayFormula() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    let headerRow = 1, tenureCol = -1, startDateCol = -1;
    
    for (let r = 1; r <= 30; r++) {
      const row = sheet.getRange(r, 1, 1, 20).getValues()[0];
      const tenureIdx = row.findIndex(cell => cell && String(cell).toLowerCase().includes("tenure"));
      const startDateIdx = row.findIndex(cell => 
        cell && (String(cell).toLowerCase().includes("start date") || String(cell).toLowerCase().includes("startdate"))
      );
      
      if (tenureIdx !== -1) {
        headerRow = r;
        tenureCol = tenureIdx + 1;
        if (startDateIdx !== -1) startDateCol = startDateIdx + 1;
        break;
      }
    }
    
    if (tenureCol === -1) {
      SpreadsheetApp.getUi().alert('Error', 'Could not find "Tenure" column in the active sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    if (startDateCol === -1) startDateCol = 3; // Default to column C
    
    const startDateColLetter = columnToLetter(startDateCol);
    const tenureColLetter = columnToLetter(tenureCol);
    const firstDataRow = headerRow + 1;
    const lastRow = sheet.getLastRow();
    
    if (lastRow < firstDataRow) {
      SpreadsheetApp.getUi().alert('Error', 'No data rows found below the header.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const tenureRange = sheet.getRange(firstDataRow, tenureCol, lastRow - firstDataRow + 1, 1);
    tenureRange.clearContent();
    
    const arrayFormula = `=ARRAYFORMULA(IF(${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastRow}="", "", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastRow})>=1460, "4 Years+", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastRow})>=1095, "3 Years", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastRow})>=730, "2 Years", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastRow})>=545, "1.5 Years", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastRow})>=365, "1 Year", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastRow})>=180, "6 Months", "Less than 6 Months"))))))))`;
    
    sheet.getRange(firstDataRow, tenureCol).setFormula(arrayFormula);
    Logger.log(`Converted tenure formulas to array formula in column ${tenureColLetter}`);
    SpreadsheetApp.getUi().alert('Success', `Converted tenure formulas to array formula in column ${tenureColLetter}!`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    Logger.log(`Error in convertTenureToArrayFormula: ${error.message}`);
    SpreadsheetApp.getUi().alert('Error', `Failed to convert tenure formulas: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    throw error;
  }
}

/**
 * Imports base employee data from Bob API
 */
function importBobDataSimpleWithLookup() {
  try {
    const reportId = BOB_REPORT_IDS.BASE_DATA;
    const sheetName = SHEET_NAMES.BASE_DATA;
    const bonusSheetName = SHEET_NAMES.BONUS_HISTORY;
    
    Logger.log(`Starting import of ${sheetName}...`);
    const rows = fetchBobReport(reportId);
    const srcHeader = rows[0];
    const normalizedHeader = normalizeHeader(srcHeader);
    
    const idxEmpId = findColCached(normalizedHeader, srcHeader, ["Employee ID", "Emp ID", "Employee Id"]);
    const idxJobLevel = findColCached(normalizedHeader, srcHeader, ["Job Level", "Job level"]);
    const idxBasePay = findColCached(normalizedHeader, srcHeader, ["Base Pay", "Base salary", "Base Salary"]);
    const idxEmpType = findColCached(normalizedHeader, srcHeader, ["Employment Type", "Employment type"]);
    const idxStartDate = findColCached(normalizedHeader, srcHeader, ["Start Date", "Start date", "Original start date", "Original Start Date"]);
    const idxTermination = findColCached(normalizedHeader, srcHeader, ["Termination Date", "Termination date"]);
    
    let header = [...srcHeader, "Variable Type", "Variable %"];
    const out = [header];
  
    for (let r = 1; r < rows.length; r++) {
      const src = rows[r];
      if (!src || !Array.isArray(src) || src.length === 0) continue;
      
      const row = src.slice();
      const empType = safeCell(row, idxEmpType);
      if (!ALLOWED_EMP_TYPES.has(empType)) continue;
      
      const empId = safeCell(row, idxEmpId);
      const jobLvl = safeCell(row, idxJobLevel);
      if (!empId || !jobLvl) continue;
      
      const empIdNum = toNumberSafe(empId);
      row[idxEmpId] = isFinite(empIdNum) ? String(empIdNum) : empId.trim();
      
      const basePayNum = toNumberSafe(safeCell(row, idxBasePay));
      if (!isFinite(basePayNum) || basePayNum === 0) continue;
      
      row[idxBasePay] = basePayNum;
      row.push("", ""); // Variable Type, Variable %
      out.push(row);
    }
    
    Logger.log(`Processed ${out.length - 1} rows for ${sheetName}`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    const sliced = out.map(r => r.slice(0, WRITE_COLS));
    
    const rowsToClear = Math.max(sheet.getMaxRows(), sliced.length);
    sheet.getRange(1, 1, rowsToClear, WRITE_COLS).clearContent();
    
    const dataRange = sheet.getRange(1, 1, sliced.length, sliced[0].length);
    dataRange.setValues(sliced);
    
    const maxRows = sheet.getMaxRows();
    const dataRows = sliced.length;
    if (maxRows > dataRows) {
      sheet.deleteRows(dataRows + 1, maxRows - dataRows);
    }
    
    dataRange.setFontFamily("Roboto");
    dataRange.setFontSize(10);
    sheet.autoResizeColumns(1, sliced[0].length);
    
    const numRows = sliced.length - 1;
    const firstDataRow = 2;
    
    if (numRows > 0) {
      if (idxEmpId + 1 <= WRITE_COLS) {
        const empIdRange = sheet.getRange(firstDataRow, idxEmpId + 1, numRows, 1);
        empIdRange.setNumberFormat("@");
        const empIdValues = empIdRange.getValues();
        empIdRange.setValues(empIdValues.map(row => [String(row[0])]));
      }
      
      if (idxBasePay + 1 <= WRITE_COLS) {
        sheet.getRange(firstDataRow, idxBasePay + 1, numRows, 1).setNumberFormat("#,##0.00");
      }
      
      const bonusSheet = ss.getSheetByName(bonusSheetName);
      if (bonusSheet) {
        const vtCol = header.indexOf("Variable Type") + 1;
        const vpCol = header.indexOf("Variable %") + 1;
        const empIdColLetter = columnToLetter(idxEmpId + 1);
        
        if (vtCol > 0 && vtCol <= WRITE_COLS) {
          const vtFormulas = Array.from({ length: numRows }, (_, i) =>
            [`=XLOOKUP($${empIdColLetter}${firstDataRow + i}, '${bonusSheetName}'!A:A, '${bonusSheetName}'!D:D, "")`]
          );
          sheet.getRange(firstDataRow, vtCol, numRows, 1).setFormulas(vtFormulas);
        }
        
        if (vpCol > 0 && vpCol <= WRITE_COLS) {
          const vpFormulas = Array.from({ length: numRows }, (_, i) =>
            [`=XLOOKUP($${empIdColLetter}${firstDataRow + i}, '${bonusSheetName}'!A:A, '${bonusSheetName}'!E:E, "")`]
          );
          sheet.getRange(firstDataRow, vpCol, numRows, 1).setFormulas(vpFormulas);
        }
      }
      
      if (idxStartDate >= 0 && idxStartDate < header.length) {
        const startDateColLetter = columnToLetter(idxStartDate + 1);
        const tenureCategoryHeaderIndex = header.indexOf("Tenure Category");
        let tenureCol = tenureCategoryHeaderIndex !== -1 ? tenureCategoryHeaderIndex + 1 : 
                       (header.length < WRITE_COLS ? header.length + 1 : WRITE_COLS + 1);
        
        if (tenureCategoryHeaderIndex === -1) {
          sheet.getRange(1, tenureCol).setValue("Tenure Category");
        }
        
        const lastRow = sheet.getLastRow();
        if (lastRow >= firstDataRow) {
          sheet.getRange(firstDataRow, tenureCol, lastRow - firstDataRow + 1, 1).clearContent();
        }
        
        const lastDataRow = firstDataRow + numRows - 1;
        const arrayFormula = `=ARRAYFORMULA(IF(${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastDataRow}="", "", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastDataRow})>=1460, "4 Years+", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastDataRow})>=1095, "3 Years", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastDataRow})>=730, "2 Years", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastDataRow})>=545, "1.5 Years", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastDataRow})>=365, "1 Year", IF((TODAY()-${startDateColLetter}${firstDataRow}:${startDateColLetter}${lastDataRow})>=180, "6 Months", "Less than 6 Months"))))))))`;
        sheet.getRange(firstDataRow, tenureCol).setFormula(arrayFormula);
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
    const rows = fetchBobReport(reportId);
    const header = rows[0];
    const normalizedHeader = normalizeHeader(header);
    
    const iEmpId = findColCached(normalizedHeader, header, ["Employee ID", "Emp ID", "Employee Id"]);
    const iName = findColCached(normalizedHeader, header, ["Display name", "Emp Name", "Display Name", "Name"]);
    const iEffDate = findColCached(normalizedHeader, header, ["Effective date", "Effective Date", "Effective"]);
    const iType = findColCached(normalizedHeader, header, ["Variable type", "Variable Type", "Type"]);
    const iPct = findColCached(normalizedHeader, header, ["Commission/Bonus %", "Variable %", "Commission %", "Bonus %"]);
    const iAmt = findColCached(normalizedHeader, header, ["Amount", "Variable Amount", "Commission/Bonus Amount"]);
    const iCurr = findColOptionalCached(normalizedHeader, header, [
      "Variable Amount currency", "Variable Amount Currency",
      "Amount currency", "Amount Currency", "Currency", "Currency code", "Currency Code"
    ]);
  
    const latest = new Map();
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length === 0) continue;
      
      let empId = safeCell(row, iEmpId);
      const empIdNum = toNumberSafe(empId);
      empId = isFinite(empIdNum) ? String(empIdNum) : empId.trim();
      
      const effRaw = safeCell(row, iEffDate);
      const effKey = (effRaw.match(/^\d{4}-\d{2}-\d{2}/) || [])[0];
      
      if (!empId || !effKey) continue;
      
      const existing = latest.get(empId);
      if (!existing || effKey > existing.effKey) {
        latest.set(empId, { row, effKey });
      }
    }
  
    const outHeader = ["Employee ID", "Display name", "Effective date", "Variable type", "Commission/Bonus %", "Amount", "Amount currency"];
    const out = [outHeader];
  
    latest.forEach(({ row, effKey }) => {
      let empId = safeCell(row, iEmpId);
      const empIdNum = toNumberSafe(empId);
      empId = isFinite(empIdNum) ? String(empIdNum) : empId.trim();
      const name = safeCell(row, iName);
      const type = safeCell(row, iType);
      const pctVal = toNumberSafe(safeCell(row, iPct));
      const amtVal = toNumberSafe(safeCell(row, iAmt));
      const curr = iCurr === -1 ? "" : safeCell(row, iCurr);
      out.push([empId, name, effKey, type, isFinite(pctVal) ? pctVal : "", isFinite(amtVal) ? amtVal : "", curr]);
    });
  
    writeToSheet(targetSheetName, out, [
      { range: [2, 3], format: "@" },
      { range: [2, 5], format: "0.########" },
      { range: [2, 6], format: "#,##0.00" }
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
    const rows = fetchBobReport(reportId);
    const header = rows[0];
    const normalizedHeader = normalizeHeader(header);
    
    const iEmpId = findColCached(normalizedHeader, header, ["Emp ID", "Employee ID", "Employee Id"]);
    const iName = findColCached(normalizedHeader, header, ["Emp Name", "Display name", "Display Name", "Name"]);
    const iEffDate = findColCached(normalizedHeader, header, ["History effective date", "Effective date", "Effective Date", "Effective"]);
    const iBase = findColCached(normalizedHeader, header, ["History base salary", "Base salary", "Base Salary", "Base pay", "Salary"]);
    const iCurr = findColCached(normalizedHeader, header, ["History base salary currency", "Base salary currency", "Currency"]);
    const iReason = findColCached(normalizedHeader, header, ["History reason", "Reason", "Change reason"]);
  
    const latest = new Map();
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length === 0) continue;
      
      let empId = safeCell(row, iEmpId);
      const empIdNum = toNumberSafe(empId);
      empId = isFinite(empIdNum) ? String(empIdNum) : empId.trim();
      
      const effStr = safeCell(row, iEffDate);
      const ymd = toYmd(effStr);
      if (!empId || !ymd) continue;
      
      const existing = latest.get(empId);
      if (!existing || ymd > existing.ymd) latest.set(empId, { row, ymd });
    }
  
    const outHeader = ["Emp ID", "Emp Name", "Effective date", "Base salary", "Base salary currency", "History reason"];
    const out = [outHeader];
  
    latest.forEach(({ row, ymd }) => {
      let empId = safeCell(row, iEmpId);
      const empIdNum = toNumberSafe(empId);
      empId = isFinite(empIdNum) ? String(empIdNum) : empId.trim();
      const name = safeCell(row, iName);
      const base = toNumberSafe(safeCell(row, iBase));
      const curr = safeCell(row, iCurr);
      const reason = safeCell(row, iReason);
      const effDate = parseDateSmart(ymd);
      out.push([empId, name, effDate, isFinite(base) ? base : "", curr, reason]);
    });
  
    writeToSheet(targetSheetName, out, [
      { range: [2, 3], format: "yyyy-mm-dd" },
      { range: [2, 4], format: "#,##0.00" }
    ]);
    
    Logger.log(`Successfully imported ${targetSheetName}`);
  } catch (error) {
    Logger.log(`Error in importBobCompHistoryLatest: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error importing Comp History: ${error.message}`);
    throw error;
  }
}

/**
 * Imports full compensation history from Bob API
 */
function importBobFullCompHistory() {
  try {
    const reportId = BOB_REPORT_IDS.FULL_COMP_HISTORY;
    const targetSheetName = SHEET_NAMES.FULL_COMP_HISTORY;
    
    Logger.log(`Starting import of ${targetSheetName}...`);
    const rows = fetchBobReport(reportId);
    const header = rows[0];
    const normalizedHeader = normalizeHeader(header);
    
    const iEmpId = findColCached(normalizedHeader, header, ["Emp ID", "Employee ID", "Employee Id"]);
    const iName = findColCached(normalizedHeader, header, ["Emp Name", "Display name", "Display Name", "Name"]);
    const iEffDate = findColCached(normalizedHeader, header, ["History effective date", "Effective date", "Effective Date", "Effective"]);
    const iBase = findColCached(normalizedHeader, header, ["History base salary", "Base salary", "Base Salary", "Base pay", "Salary"]);
    const iCurr = findColCached(normalizedHeader, header, ["History base salary currency", "Base salary currency", "Currency"]);
    const iReason = findColCached(normalizedHeader, header, ["History reason", "Reason", "Change reason"]);
    
    const employees = new Map();
    const compHistory = new Map();
    const promotions = new Map();
    
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length === 0) continue;
      
      let empId = safeCell(row, iEmpId);
      const empIdNum = toNumberSafe(empId);
      empId = isFinite(empIdNum) ? String(empIdNum) : empId.trim();
      
      if (!empId) continue;
      
      const name = safeCell(row, iName);
      if (!employees.has(empId)) {
        employees.set(empId, { name: name || "" });
      }
      
      const effStr = safeCell(row, iEffDate);
      const ymd = toYmd(effStr);
      const base = toNumberSafe(safeCell(row, iBase));
      const currency = safeCell(row, iCurr);
      const reason = safeCell(row, iReason);
      
      if (!ymd) continue;
      
      if (!compHistory.has(empId)) {
        compHistory.set(empId, []);
      }
      compHistory.get(empId).push({ row, ymd, base, currency, reason });
      
      const reasonLower = reason.toLowerCase();
      if (reasonLower.includes("promotion")) {
        const existing = promotions.get(empId);
        if (!existing || ymd > existing.ymd) {
          promotions.set(empId, { row, ymd });
        }
      }
    }
    
    compHistory.forEach((entries) => {
      entries.sort((a, b) => {
        if (b.ymd > a.ymd) return 1;
        if (b.ymd < a.ymd) return -1;
        return 0;
      });
    });
    
    const outHeader = ["Emp ID", "Emp Name", "Last Promotion Date", "Last Increase Date", "Last Increase %", "Change Reason"];
    const out = [outHeader];
    
    employees.forEach((employee, empId) => {
      const promotion = promotions.get(empId);
      let promotionDate = "";
      
      if (promotion) {
        const effStr = safeCell(promotion.row, iEffDate);
        const ymd = toYmd(effStr);
        if (ymd) promotionDate = parseDateSmart(ymd);
      }
      
      let lastIncreaseDate = "", lastIncreasePct = "", changeReason = "";
      const history = compHistory.get(empId);
      
      if (history && history.length >= 2) {
        const lastEntry = history[0];
        const previousEntry = history[1];
        const lastBase = lastEntry.base;
        const prevBase = previousEntry.base;
        const lastCurrency = (lastEntry.currency || "").trim();
        const prevCurrency = (previousEntry.currency || "").trim();
        
        if (isFinite(lastBase) && isFinite(prevBase) && prevBase > 0 && 
            lastCurrency && prevCurrency && lastCurrency === prevCurrency) {
          const increasePct = ((lastBase - prevBase) / prevBase) * 100;
          if (increasePct >= 0) {
            lastIncreasePct = increasePct.toFixed(2);
            changeReason = lastEntry.reason || "";
            // Get the date of the increase (effective date of the last entry)
            if (lastEntry.ymd) {
              lastIncreaseDate = parseDateSmart(lastEntry.ymd);
            }
          }
        }
      }
      
      out.push([empId, employee.name, promotionDate, lastIncreaseDate, lastIncreasePct, changeReason]);
    });
    
    writeToSheet(targetSheetName, out, [
      { range: [2, 3], format: "yyyy-mm-dd" },  // Last Promotion Date
      { range: [2, 4], format: "yyyy-mm-dd" },  // Last Increase Date
      { range: [2, 5], format: "0.00" }         // Last Increase %
    ]);
    
    Logger.log(`Successfully imported ${targetSheetName} - ${out.length - 1} employees, ${promotions.size} with promotion dates`);
  } catch (error) {
    Logger.log(`Error in importBobFullCompHistory: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error importing Full Comp History: ${error.message}`);
    throw error;
  }
}

/**
 * Imports historical performance ratings from Bob API
 */
function importBobPerformanceRatings() {
  try {
    const reportId = BOB_REPORT_IDS.PERF_RATINGS;
    const targetSheetName = SHEET_NAMES.PERF_RATINGS;
    
    Logger.log(`Starting import of ${targetSheetName}...`);
    const rows = fetchBobReport(reportId);
    
    if (!rows || rows.length === 0) {
      throw new Error("No data returned from report");
    }
    
    // Use the header from the report
    const header = rows[0];
    
    // Import all rows as-is (no transformation needed for performance ratings)
    const out = [header];
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (row && row.length > 0) {
        out.push(row);
      }
    }
    
    // Write to sheet with formatting
    writeToSheet(targetSheetName, out, [
      // Format date columns if they exist (will be detected automatically)
    ]);
    
    Logger.log(`Successfully imported ${targetSheetName} - ${out.length - 1} rows`);
    SpreadsheetApp.getUi().alert(`Successfully imported ${out.length - 1} performance rating records to ${targetSheetName}`);
  } catch (error) {
    Logger.log(`Error in importBobPerformanceRatings: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error importing Performance Ratings: ${error.message}`);
    throw error;
  }
}

// ============================================================================
// PERFORMANCE REPORT FUNCTIONS
// ============================================================================

/**
 * Prompts user to set HiBob login credentials
 */
function setHiBobCredentials() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  const emailResponse = ui.prompt('Set HiBob Email', 'Enter your HiBob email address:', ui.ButtonSet.OK_CANCEL);
  if (emailResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const email = emailResponse.getResponseText().trim();
  if (!email) {
    ui.alert('Error', 'Email cannot be empty.', ui.ButtonSet.OK);
    return;
  }
  
  const passwordResponse = ui.prompt('Set HiBob Password', 'Enter your JumpCloud password:', ui.ButtonSet.OK_CANCEL);
  if (passwordResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const password = passwordResponse.getResponseText().trim();
  if (!password) {
    ui.alert('Error', 'Password cannot be empty.', ui.ButtonSet.OK);
    return;
  }
  
  props.setProperty('HIBOB_EMAIL', email);
  props.setProperty('HIBOB_PASSWORD', password);
  
  ui.alert('Success', 'Credentials saved successfully!\n\nEmail: ' + email, ui.ButtonSet.OK);
  Logger.log('HiBob credentials updated for: ' + email);
}

/**
 * Shows the status of stored credentials
 */
function viewCredentialsStatus() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  const email = props.getProperty('HIBOB_EMAIL');
  const hasPassword = props.getProperty('HIBOB_PASSWORD') ? 'Yes' : 'No';
  
  const status = 'HiBob Credentials Status\n\n' +
    'Email: ' + (email || 'Not set') + '\n' +
    'Password: ' + hasPassword + '\n\n' +
    (email ? 'Credentials are ready to use.' : 'Please set credentials using "Set HiBob Credentials" menu item.');
  
  ui.alert('Credentials Status', status, ui.ButtonSet.OK);
}

/**
 * Gets stored HiBob credentials
 */
function getHiBobCredentials() {
  const props = PropertiesService.getScriptProperties();
  const email = props.getProperty('HIBOB_EMAIL');
  const password = props.getProperty('HIBOB_PASSWORD');
  
  if (!email || !password) return null;
  return { email, password };
}

/**
 * Triggers a message about downloading performance report
 */
function triggerPerformanceReportDownload() {
  const ui = SpreadsheetApp.getUi();
  const creds = getHiBobCredentials();
  const hasCreds = creds ? 'Yes' : 'No';
  
  const message = 'Performance Report Download\n\n' +
    'Credentials stored: ' + hasCreds + '\n\n' +
    'To download a report:\n' +
    '1. Open terminal on your local machine\n' +
    '2. Navigate to project directory\n' +
    '3. Run: python3 hibob_report_downloader.py\n' +
    '4. Enter the report name when prompted\n\n' +
    (creds ? 
      '✅ Credentials are stored and ready to use.\nThe Python script can fetch them automatically.' :
      '⚠️ Please set credentials first using "Set HiBob Credentials"');
  
  ui.alert('Download Performance Report', message, ui.ButtonSet.OK);
}

/**
 * Generate Manager Blurbs using AI-powered summarization
 * Provides instructions for running the Python script
 */
function generateManagerBlurbs() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body {
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
          padding: 20px;
          max-width: 700px;
          margin: 0 auto;
        }
        h2 {
          color: #9825ff;
          margin-top: 0;
        }
        .step {
          background: #f8f9fa;
          padding: 15px;
          margin: 10px 0;
          border-radius: 8px;
          border-left: 4px solid #9825ff;
        }
        .step-number {
          font-weight: bold;
          color: #9825ff;
        }
        .code {
          background: #2d2d2d;
          color: #f8f8f2;
          padding: 10px;
          border-radius: 5px;
          font-family: 'Courier New', monospace;
          margin: 10px 0;
          overflow-x: auto;
          font-size: 13px;
        }
        .info {
          background: #d1ecf1;
          border-left: 4px solid #0c5460;
          padding: 10px;
          border-radius: 5px;
          margin: 10px 0;
        }
        .warning {
          background: #fff3cd;
          border-left: 4px solid #ffc107;
          padding: 10px;
          border-radius: 5px;
          margin: 10px 0;
        }
        ul {
          line-height: 1.6;
        }
      </style>
    </head>
    <body>
      <h2>🤖 Generate Manager Blurbs</h2>
      
      <div class="info">
        ℹ️ <strong>About:</strong> This feature uses AI (BART model) to summarize manager feedback into concise 50-60 word performance review blurbs. 
        Blurbs are stored in a hidden "Manager Blurbs" sheet and automatically referenced in the Summary sheet.
      </div>
      
      <div class="step">
        <span class="step-number">Step 1:</span> Ensure you have the required Python libraries
        <div class="code">pip install transformers torch google-auth google-api-python-client</div>
      </div>
      
      <div class="step">
        <span class="step-number">Step 2:</span> Make sure Bob Perf Report contains manager feedback
        <ul style="margin: 10px 0; padding-left: 20px;">
          <li>Import the Performance Report (via Web Interface)</li>
          <li>Verify the sheet contains columns with manager comments</li>
          <li>Expected columns: leadership strength/improvement, AI readiness, performance comments, etc.</li>
        </ul>
      </div>
      
      <div class="step">
        <span class="step-number">Step 3:</span> Run the Manager Blurb Generator script
        <div class="code">cd scripts/python<br>python3 manager_blurb_generator.py</div>
        <p style="font-size: 12px; color: #666; margin-top: 5px;">
          ⏱️ Note: First run may take 1-2 minutes to download the AI model
        </p>
      </div>
      
      <div class="warning">
        ⚠️ <strong>Processing Time:</strong> The script will process all employees and generate blurbs using AI. 
        This may take 2-5 minutes depending on the number of employees and your machine's performance.
      </div>
      
      <div class="step">
        <span class="step-number">Step 4:</span> Refresh the Summary sheet
        <ul style="margin: 10px 0; padding-left: 20px;">
          <li>Run "Build Summary Sheet" from the menu</li>
          <li>The "Manager Blurb" column will auto-populate from the hidden sheet</li>
          <li>If no blurb is found, it will show "No blurb generated"</li>
        </ul>
      </div>
      
      <div class="info">
        💡 <strong>Blurb Format:</strong>
        <ul style="margin: 10px 0; padding-left: 20px;">
          <li>Starts with action verbs (Delivered, Demonstrated, etc.)</li>
          <li>Highlights AI adoption where relevant</li>
          <li>Ends with development/growth focus</li>
          <li>Gender-neutral, performance-review tone</li>
          <li>~50-60 words, free of filler and leadership jargon</li>
        </ul>
      </div>
      
      <p style="text-align: center; color: #666; margin-top: 30px;">
        <button onclick="google.script.host.close()" style="background: #9825ff; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-size: 14px;">
          ✓ Got it
        </button>
      </p>
    </body>
    </html>
  `)
    .setWidth(750)
    .setHeight(650);
  
  SpreadsheetApp.getUi().showModalDialog(html, '🤖 Manager Blurb Generator');
}

/**
 * Update Slicers Only
 * 
 * Updates slicer data ranges to match current Summary sheet row count
 * without rebuilding the entire sheet. Safe to run - preserves all data,
 * formatting, and existing slicer filters.
 */
function updateSlicersOnly() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = ss.getSheetByName("Summary");
    
    if (!summarySheet) {
      SpreadsheetApp.getUi().alert(
        "Error",
        "Summary sheet not found. Please create it first by running 'Build Summary Sheet'.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Get the current number of data rows
    // Assuming header is in row 20, data starts in row 21
    const lastRow = summarySheet.getLastRow();
    const dataStartRow = 20;
    
    if (lastRow <= dataStartRow) {
      SpreadsheetApp.getUi().alert(
        "Error",
        "No data found in Summary sheet. Please run 'Build Summary Sheet' first.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const numDataRows = lastRow - dataStartRow;
    
    Logger.log(`Updating slicers for ${numDataRows} data rows (rows ${dataStartRow + 1} to ${lastRow})`);
    
    // Update slicer ranges
    updateSlicerRanges(summarySheet, numDataRows, dataStartRow - 1);
    
    SpreadsheetApp.getUi().alert(
      "Success",
      `Slicer ranges updated successfully!\n\n` +
      `Data rows: ${numDataRows}\n` +
      `Range: Row ${dataStartRow + 1} to ${lastRow}\n\n` +
      `All slicer filters and formatting preserved.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    Logger.log(`✓ Slicers updated successfully for ${numDataRows} rows`);
    
  } catch (error) {
    Logger.log(`Error updating slicers: ${error.message}`);
    SpreadsheetApp.getUi().alert(
      "Error",
      `Failed to update slicers: ${error.message}\n\n` +
      `Check that the Summary sheet exists and has data.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

/**
 * Create AI Readiness Mapping Sheet
 * 
 * Creates a reference sheet with Department → AI Readiness Category mapping
 * and validation dropdown for leadership to assess employees
 */
function createAIReadinessMapping() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    // Check if sheet already exists
    let mappingSheet = ss.getSheetByName("AI Readiness Mapping");
    
    if (mappingSheet) {
      const response = ui.alert(
        "Sheet Exists",
        "AI Readiness Mapping sheet already exists. Do you want to recreate it? This will overwrite existing data.",
        ui.ButtonSet.YES_NO
      );
      
      if (response !== ui.Button.YES) {
        return;
      }
      
      ss.deleteSheet(mappingSheet);
    }
    
    // Create new mapping sheet
    mappingSheet = ss.insertSheet("AI Readiness Mapping");
    
    // Define function-to-category mapping
    const mappingData = [
      ["Department/Function", "AI Readiness Category"],
      ["Engineering", "AI-Accelerated Developer"],
      ["Product", "AI-Forward Solutions Engineer"],
      ["Customer Success", "AI Outcome Manager"],
      ["Marketing", "AI Content Intelligence Strategist"],
      ["Sales", "AI-Augmented Revenue Accelerator"],
      ["Human Resources", "AI-Powered People Strategist"],
      ["People Ops", "AI-Powered People Strategist"],
      ["HR", "AI-Powered People Strategist"],
      ["Finance", "AI-Driven Financial Intelligence Analyst"],
      ["General & Administrative", "AI Operations Optimizer"],
      ["G&A", "AI Operations Optimizer"],
      ["Operations", "AI Operations Optimizer"],
      ["Support", "AI-Enhanced Support Specialist"],
      ["Service Delivery", "AI-Enhanced Support Specialist"],
      ["Data", "AI-Native Insights Engineer"],
      ["Analytics", "AI-Native Insights Engineer"],
      ["Data & Analytics", "AI-Native Insights Engineer"],
      ["Design", "AI-Augmented Experience Designer"],
      ["UX", "AI-Augmented Experience Designer"],
      ["Information Security", "AI-Powered Security Guardian"],
      ["Security", "AI-Powered Security Guardian"],
      ["Legal", "AI-Assisted Legal Intelligence Analyst"],
      ["Compliance", "AI-Assisted Legal Intelligence Analyst"]
    ];
    
    // Write mapping data
    const mappingRange = mappingSheet.getRange(1, 1, mappingData.length, 2);
    mappingRange.setValues(mappingData);
    
    // Format header
    const headerRange = mappingSheet.getRange(1, 1, 1, 2);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#9825ff");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontFamily("Roboto");
    headerRange.setFontSize(11);
    
    // Format data rows
    const dataRange = mappingSheet.getRange(2, 1, mappingData.length - 1, 2);
    dataRange.setFontFamily("Roboto");
    dataRange.setFontSize(10);
    
    // Alternate row colors for readability
    for (let i = 2; i <= mappingData.length; i++) {
      const rowRange = mappingSheet.getRange(i, 1, 1, 2);
      if (i % 2 === 0) {
        rowRange.setBackground("#f3f3f3");
      }
    }
    
    // Auto-resize columns
    mappingSheet.autoResizeColumns(1, 2);
    
    // Freeze header row
    mappingSheet.setFrozenRows(1);
    
    // Add instructions in column D
    const instructions = [
      ["Instructions"],
      ["This sheet maps each Department/Function to its AI Readiness Category."],
      [""],
      ["The Summary sheet's 'AI Readiness Category' column uses VLOOKUP to auto-populate based on employee department."],
      [""],
      ["To assess employees:"],
      ["1. Review the AI Readiness Categories documentation (see docs/AI_READINESS_CATEGORIES.md)"],
      ["2. Add a 'AI Readiness Assessment' column to Summary sheet"],
      ["3. Use data validation with these options: 'AI-Ready', 'AI-Capable', 'Not AI-Ready'"],
      ["4. Leadership evaluates each employee against their function's criteria"],
      [""],
      ["Category Descriptions:"],
      ["• AI-Ready: Demonstrates 3+ criteria, measurable impact, proactive adoption"],
      ["• AI-Capable: Demonstrates 2 criteria, uses AI but limited impact, needs coaching"],
      ["• Not AI-Ready: 0-1 criteria, resistant to AI, manual-first approach"],
      [""],
      ["📖 Full criteria available in: docs/AI_READINESS_CATEGORIES.md"]
    ];
    
    const instructionsRange = mappingSheet.getRange(1, 4, instructions.length, 1);
    instructionsRange.setValues(instructions);
    instructionsRange.setFontFamily("Roboto");
    instructionsRange.setFontSize(9);
    instructionsRange.setWrap(true);
    
    // Format instructions header
    const instructionsHeader = mappingSheet.getRange(1, 4, 1, 1);
    instructionsHeader.setFontWeight("bold");
    instructionsHeader.setBackground("#ff9901");
    instructionsHeader.setFontColor("#ffffff");
    instructionsHeader.setFontSize(10);
    
    // Set column widths
    mappingSheet.setColumnWidth(1, 250);
    mappingSheet.setColumnWidth(2, 350);
    mappingSheet.setColumnWidth(4, 600);
    
    ui.alert(
      "Success",
      "AI Readiness Mapping sheet created successfully!\n\n" +
      "Next steps:\n" +
      "1. Review/customize department mappings if needed\n" +
      "2. Run 'Build Summary Sheet' to populate the AI Readiness Category column\n" +
      "3. Add an 'AI Readiness Assessment' column for leadership evaluation\n\n" +
      "📖 Full category criteria: docs/AI_READINESS_CATEGORIES.md",
      ui.ButtonSet.OK
    );
    
    Logger.log("AI Readiness Mapping sheet created successfully");
    
  } catch (error) {
    Logger.log(`Error creating AI Readiness Mapping sheet: ${error.message}`);
    SpreadsheetApp.getUi().alert(
      "Error",
      `Failed to create AI Readiness Mapping sheet: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

/**
 * Launches the web interface for Performance Report automation
 * Opens a dialog with instructions and a link to the web interface
 */
function launchWebInterface() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body {
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
          padding: 20px;
          max-width: 600px;
          margin: 0 auto;
        }
        h2 {
          color: #667eea;
          margin-top: 0;
        }
        .step {
          background: #f8f9fa;
          padding: 15px;
          margin: 10px 0;
          border-radius: 8px;
          border-left: 4px solid #667eea;
        }
        .step-number {
          font-weight: bold;
          color: #667eea;
        }
        .code {
          background: #2d2d2d;
          color: #f8f8f2;
          padding: 10px;
          border-radius: 5px;
          font-family: 'Courier New', monospace;
          margin: 10px 0;
          overflow-x: auto;
        }
        .button {
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          color: white;
          padding: 12px 24px;
          border: none;
          border-radius: 8px;
          font-size: 16px;
          font-weight: 600;
          cursor: pointer;
          text-decoration: none;
          display: inline-block;
          margin: 10px 5px;
          transition: transform 0.2s;
        }
        .button:hover {
          transform: translateY(-2px);
        }
        .button-secondary {
          background: #6c757d;
        }
        .warning {
          background: #fff3cd;
          border-left-color: #ffc107;
          padding: 10px;
          border-radius: 5px;
          margin: 10px 0;
        }
        .success {
          background: #d4edda;
          border-left-color: #28a745;
          padding: 10px;
          border-radius: 5px;
          margin: 10px 0;
        }
      </style>
    </head>
    <body>
      <h2>🚀 Launch Web Interface</h2>
      
      <div class="step">
        <span class="step-number">Step 1:</span> Start the web server on your local machine
        <div class="code">./start_web_app.sh</div>
        <p style="font-size: 12px; color: #666; margin-top: 5px;">
          Or on Windows: <code>start_web_app.bat</code>
        </p>
      </div>
      
      <div class="warning">
        ⚠️ <strong>Important:</strong> The web interface runs on your local machine (localhost). 
        Make sure the server is running before clicking the button below.
      </div>
      
      <div class="step">
        <span class="step-number">Step 2:</span> Open the web interface
        <br><br>
        <a href="http://localhost:5001" target="_blank" class="button">
          🌐 Open Web Interface
        </a>
        <a href="http://127.0.0.1:5001" target="_blank" class="button button-secondary">
          🔗 Alternative URL
        </a>
      </div>
      
      <div class="success">
        ✅ <strong>Features:</strong>
        <ul style="margin: 5px 0; padding-left: 20px;">
          <li>Visual progress tracking</li>
          <li>Checkbox selection for reports</li>
          <li>Real-time status updates</li>
          <li>No terminal interaction needed</li>
        </ul>
      </div>
      
      <div class="step">
        <strong>Default Port:</strong> 5001<br>
        <strong>Custom Port:</strong> Run <code>./start_web_app.sh 8080</code> to use port 8080
      </div>
      
      <p style="text-align: center; margin-top: 20px;">
        <button onclick="google.script.host.close()" class="button button-secondary">
          Close
        </button>
      </p>
    </body>
    </html>
  `)
    .setWidth(650)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Launch Web Interface');
}

/**
 * Shows instructions for using the Performance Report automation
 */
function showPerformanceReportInstructions() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body {
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
          padding: 20px;
          max-width: 700px;
          margin: 0 auto;
          line-height: 1.6;
        }
        h2 {
          color: #667eea;
          margin-top: 0;
          border-bottom: 2px solid #667eea;
          padding-bottom: 10px;
        }
        h3 {
          color: #764ba2;
          margin-top: 20px;
        }
        .method {
          background: #f8f9fa;
          padding: 15px;
          margin: 15px 0;
          border-radius: 8px;
          border-left: 4px solid #667eea;
        }
        .method.recommended {
          border-left-color: #28a745;
          background: #d4edda;
        }
        .method-title {
          font-weight: bold;
          font-size: 18px;
          margin-bottom: 10px;
        }
        .method.recommended .method-title::before {
          content: "⭐ ";
        }
        .step {
          margin: 8px 0;
          padding-left: 20px;
        }
        .code {
          background: #2d2d2d;
          color: #f8f8f2;
          padding: 10px;
          border-radius: 5px;
          font-family: 'Courier New', monospace;
          margin: 10px 0;
          overflow-x: auto;
          font-size: 14px;
        }
        .info-box {
          background: #e3f2fd;
          border-left: 4px solid #2196f3;
          padding: 12px;
          margin: 15px 0;
          border-radius: 5px;
        }
        .button {
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          color: white;
          padding: 10px 20px;
          border: none;
          border-radius: 8px;
          font-size: 14px;
          font-weight: 600;
          cursor: pointer;
          margin-top: 10px;
        }
        .button:hover {
          transform: translateY(-2px);
        }
        ul {
          margin: 10px 0;
          padding-left: 25px;
        }
        li {
          margin: 5px 0;
        }
      </style>
    </head>
    <body>
      <h2>📖 HiBob Performance Report Automation</h2>
      
      <div class="info-box">
        <strong>Purpose:</strong> Download performance reports from HiBob (no API available) 
        and automatically upload them to Google Sheets for performance cycle reporting.
      </div>
      
      <div class="method recommended">
        <div class="method-title">Method 1: Web Interface (Recommended)</div>
        <div class="step">1. Click <strong>"🚀 Launch Web Interface"</strong> from the menu</div>
        <div class="step">2. Start the web server on your local machine:</div>
        <div class="code">./start_web_app.sh</div>
        <div class="step">3. Click the button in the dialog to open the web interface</div>
        <div class="step">4. Enter report name and click "Start Automation"</div>
        <div class="step">5. Select report using checkboxes and watch progress</div>
        
        <strong>Features:</strong>
        <ul>
          <li>✅ Visual progress tracking with progress bar</li>
          <li>✅ Checkbox selection for reports</li>
          <li>✅ Real-time status updates</li>
          <li>✅ Runs in background (no popup windows)</li>
          <li>✅ No terminal interaction needed</li>
        </ul>
      </div>
      
      <div class="method">
        <div class="method-title">Method 2: Terminal (Alternative)</div>
        <div class="step">1. Set credentials: <strong>"Set HiBob Credentials"</strong> from menu</div>
        <div class="step">2. Open terminal and navigate to project directory</div>
        <div class="step">3. Run:</div>
        <div class="code">python3 hibob_report_downloader.py</div>
        <div class="step">4. Enter report name when prompted</div>
      </div>
      
      <h3>🔐 Credentials Setup</h3>
      <div class="info-box">
        <strong>Option 1:</strong> Store in Apps Script (Recommended)
        <ul>
          <li>Click "Set HiBob Credentials" from menu</li>
          <li>Enter your HiBob email and password</li>
          <li>Credentials are stored securely in Apps Script</li>
          <li>Python script can fetch them automatically</li>
        </ul>
        
        <strong>Option 2:</strong> Use config.json file
        <ul>
          <li>Edit <code>config.json</code> with your credentials</li>
          <li>File is gitignored (not committed to repository)</li>
        </ul>
      </div>
      
      <h3>📋 How It Works</h3>
      <ol>
        <li>Automation logs into HiBob via JumpCloud SSO</li>
        <li>Navigates to Performance Cycles page</li>
        <li>Searches for the specified report</li>
        <li>Downloads the Excel report</li>
        <li>Reads the Excel file content</li>
        <li>Uploads data to "Bob Perf Report" sheet in Google Sheets</li>
        <li>Formats the header row (blue background, bold text)</li>
      </ol>
      
      <h3>⚙️ Local Setup (First Time Only)</h3>
      <div class="code">
# Install Python dependencies<br>
pip3 install -r requirements.txt<br><br>
# Install Playwright browser<br>
playwright install chromium<br><br>
# For web interface, use the startup script:<br>
./start_web_app.sh
      </div>
      
      <h3>🔗 Google Sheets Integration</h3>
      <div class="info-box">
        Reports are automatically uploaded to:<br>
        <strong>Sheet:</strong> "Bob Perf Report"<br>
        <strong>Sheet ID:</strong> 1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA<br><br>
        Uses Google Sheets API with service account for seamless automation.
      </div>
      
      <p style="text-align: center; margin-top: 30px;">
        <button onclick="google.script.host.close()" class="button">
          Close
        </button>
      </p>
    </body>
    </html>
  `)
    .setWidth(750)
    .setHeight(700);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Performance Report Instructions');
}

// ============================================================================
// WEB APP HANDLERS (doGet, doPost)
function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'getCredentials') {
      return getCredentialsAPI(e);
    }
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: 'Invalid action'
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log('Error in doGet: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle POST requests from the Python automation
 */
function doPost(e) {
  try {
    Logger.log('Received POST request');
    
    if (!e || !e.postData) {
      return createResponse(400, 'No data received');
    }
    
    // Check if it's JSON data (new method) or multipart file (old method)
    const contentType = e.postData.type || '';
    
    if (contentType.includes('application/json')) {
      // New method: JSON data with Excel content
      try {
        const jsonData = JSON.parse(e.postData.contents);
        
        if (!jsonData.data || !Array.isArray(jsonData.data)) {
          Logger.log('Invalid JSON data structure');
          return createResponse(400, 'Invalid JSON data. Expected "data" array.');
        }
        
        const sheetName = jsonData.sheet_name || SHEET_NAMES.PERF_REPORT;
        const filename = jsonData.filename || 'report.xlsx';
        
        Logger.log('JSON data received: ' + jsonData.data.length + ' rows, filename: ' + filename);
        
        // Check if data is too large
        if (jsonData.data.length > 100000) {
          Logger.log('Warning: Very large dataset: ' + jsonData.data.length + ' rows');
        }
        
        const result = importDataToSheet(jsonData.data, sheetName);
        
        if (result.success) {
          return createResponse(200, result.message, {
            rowsImported: result.rowsImported,
            sheetName: sheetName,
            sheetId: PERF_SHEET_ID
          });
        } else {
          Logger.log('Import failed: ' + result.message);
          return createResponse(500, result.message);
        }
      } catch (parseError) {
        Logger.log('JSON parse error: ' + parseError.toString());
        Logger.log('Content type: ' + contentType);
        Logger.log('Content length: ' + (e.postData.contents ? e.postData.contents.length : 0));
        return createResponse(400, 'Failed to parse JSON: ' + parseError.toString());
      }
    } else {
      // Old method: multipart file upload (for backward compatibility)
      const boundary = extractBoundary(contentType);
      if (!boundary) {
        return createResponse(400, 'Invalid content type. Expected application/json or multipart/form-data');
      }
      
      const fileData = extractFileFromMultipart(e.postData.contents, boundary);
      if (!fileData) {
        return createResponse(400, 'No file found in request');
      }
      
      Logger.log('File received: ' + fileData.filename);
      const result = importFileToSheet(fileData);
      
      if (result.success) {
        return createResponse(200, result.message, {
          rowsImported: result.rowsImported,
          sheetName: SHEET_NAMES.PERF_REPORT,
          sheetId: PERF_SHEET_ID
        });
      } else {
        return createResponse(500, result.message);
      }
    }
  } catch (error) {
    Logger.log('Error in doPost: ' + error.toString());
    return createResponse(500, 'Internal error: ' + error.toString());
  }
}

/**
 * API endpoint to get credentials (called by Python script)
 */
function getCredentialsAPI(e) {
  try {
    const creds = getHiBobCredentials();
    
    if (!creds) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: 'Credentials not set. Please use "Set HiBob Credentials" menu item in Google Sheets.'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      email: creds.email,
      password: creds.password
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log('Error in getCredentialsAPI: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Import data array directly to Google Sheet (for Performance Reports)
 * This is the new method that receives parsed Excel data as JSON
 */
function importDataToSheet(data, sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(PERF_SHEET_ID);
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (sheet) {
      sheet.clear();
      Logger.log('Cleared existing sheet: ' + sheetName);
    } else {
      sheet = spreadsheet.insertSheet(sheetName);
      Logger.log('Created new sheet: ' + sheetName);
    }
    
    if (!data || data.length === 0) {
      return {
        success: false,
        message: 'No data provided'
      };
    }
    
    const numRows = data.length;
    const numCols = data[0].length;
    
    // Write data to sheet
    sheet.getRange(1, 1, numRows, numCols).setValues(data);
    
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, numCols);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    
    // Auto-resize columns
    for (let i = 1; i <= numCols; i++) {
      sheet.autoResizeColumn(i);
    }
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    Logger.log('Successfully imported ' + numRows + ' rows to ' + sheetName);
    
    return {
      success: true,
      message: 'Successfully imported report to Google Sheet',
      rowsImported: numRows
    };
  } catch (error) {
    Logger.log('Error importing to sheet: ' + error.toString());
    return {
      success: false,
      message: 'Error importing to sheet: ' + error.toString()
    };
  }
}

/**
 * Import file to Google Sheet (for Performance Reports)
 * This is the old method for backward compatibility with CSV files
 */
function importFileToSheet(fileData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(PERF_SHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAMES.PERF_REPORT);
    
    if (sheet) {
      sheet.clear();
      Logger.log('Cleared existing sheet: ' + SHEET_NAMES.PERF_REPORT);
    } else {
      sheet = spreadsheet.insertSheet(SHEET_NAMES.PERF_REPORT);
      Logger.log('Created new sheet: ' + SHEET_NAMES.PERF_REPORT);
    }
    
    let data = [];
    const filename = fileData.filename.toLowerCase();
    
    if (filename.endsWith('.csv')) {
      data = parseCSV(fileData.content);
    } else if (filename.endsWith('.xlsx') || filename.endsWith('.xls')) {
      return {
        success: false,
        message: 'Excel file format not yet supported. Please download as CSV format.'
      };
    } else {
      data = parseCSV(fileData.content);
    }
    
    if (!data || data.length === 0) {
      return {
        success: false,
        message: 'No data found in file or unable to parse file format'
      };
    }
    
    const numRows = data.length;
    const numCols = data[0].length;
    
    sheet.getRange(1, 1, numRows, numCols).setValues(data);
    
    const headerRange = sheet.getRange(1, 1, 1, numCols);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    
    for (let i = 1; i <= numCols; i++) {
      sheet.autoResizeColumn(i);
    }
    
    sheet.setFrozenRows(1);
    
    Logger.log('Successfully imported ' + numRows + ' rows to ' + SHEET_NAMES.PERF_REPORT);
    
    return {
      success: true,
      message: 'Successfully imported report to Google Sheet',
      rowsImported: numRows
    };
  } catch (error) {
    Logger.log('Error importing to sheet: ' + error.toString());
    return {
      success: false,
      message: 'Error importing to sheet: ' + error.toString()
    };
  }
}

/**
 * Extract boundary from content type
 */
function extractBoundary(contentType) {
  const match = contentType.match(/boundary=([^;]+)/);
  return match ? match[1].trim() : null;
}

/**
 * Extract file data from multipart form data
 */
function extractFileFromMultipart(contents, boundary) {
  try {
    const parts = contents.split('--' + boundary);
    
    for (let i = 0; i < parts.length; i++) {
      const part = parts[i];
      
      if (part.indexOf('filename=') !== -1) {
        const filenameMatch = part.match(/filename="([^"]+)"/);
        const filename = filenameMatch ? filenameMatch[1] : 'unknown';
        
        const contentTypeMatch = part.match(/Content-Type: ([^\r\n]+)/);
        const contentType = contentTypeMatch ? contentTypeMatch[1].trim() : 'application/octet-stream';
        
        const contentStart = part.indexOf('\r\n\r\n');
        if (contentStart !== -1) {
          const fileContent = part.substring(contentStart + 4);
          return {
            filename: filename,
            contentType: contentType,
            content: fileContent.trim()
          };
        }
      }
    }
    
    return null;
  } catch (error) {
    Logger.log('Error extracting file: ' + error.toString());
    return null;
  }
}

/**
 * Parse CSV content into 2D array
 */
function parseCSV(content) {
  try {
    const lines = content.split(/\r?\n/);
    const data = [];
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (line) {
        const row = parseCSVLine(line);
        data.push(row);
      }
    }
    
    return data;
  } catch (error) {
    Logger.log('Error parsing CSV: ' + error.toString());
    return [];
  }
}

/**
 * Parse a single CSV line handling quotes and escaped characters
 */
function parseCSVLine(line) {
  const result = [];
  let cell = '';
  let inQuotes = false;
  
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    const nextChar = i < line.length - 1 ? line[i + 1] : null;
    
    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        cell += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      result.push(cell);
      cell = '';
    } else {
      cell += char;
    }
  }
  
  result.push(cell);
  return result;
}

/**
 * Create HTTP response
 */
function createResponse(statusCode, message, additionalData) {
  const response = {
    status: statusCode,
    message: message
  };
  
  if (additionalData) {
    Object.assign(response, additionalData);
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================================
// HELPER FUNCTIONS - Shared
// ============================================================================

/**
 * Fetches a report from Bob API and returns parsed CSV rows
 */
function fetchBobReport(reportId) {
  const apiUrl = `https://api.hibob.com/v1/company/reports/${reportId}/download?format=csv&locale=en-CA`;
  const creds = getBobCredentials();
  const basicAuth = Utilities.base64Encode(`${creds.id}:${creds.key}`);
  
  const res = UrlFetchApp.fetch(apiUrl, {
    method: "get",
    headers: { 
      accept: "text/csv",
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
 */
function normalizeHeader(headerRow) {
  const norm = (s) => String(s || "").toLowerCase().replace(/\s+/g, " ").trim();
  return headerRow.map(norm);
}

/**
 * Finds a column index using cached normalized header
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
 */
function writeToSheet(sheetName, data, formats = []) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  
  sheet.clearContents();
  const dataRange = sheet.getRange(1, 1, data.length, data[0].length);
  dataRange.setValues(data);
  
  const maxRows = sheet.getMaxRows();
  const dataRows = data.length;
  if (maxRows > dataRows) {
    sheet.deleteRows(dataRows + 1, maxRows - dataRows);
  }
  
  dataRange.setFontFamily("Roboto");
  dataRange.setFontSize(10);
  sheet.autoResizeColumns(1, data[0].length);
  
  const header = data[0];
  const empIdHeaderIndex = header.findIndex(h => 
    h && (h.toString().toLowerCase().includes("employee id") || h.toString().toLowerCase().includes("emp id"))
  );
  if (empIdHeaderIndex >= 0 && data.length > 1) {
    const empIdCol = empIdHeaderIndex + 1;
    const empIdRange = sheet.getRange(2, empIdCol, data.length - 1, 1);
    empIdRange.setNumberFormat("@");
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
 */
function safeCell(row, idx) {
  return idx === -1 ? "" : (row[idx] == null ? "" : String(row[idx]).trim());
}

/**
 * Safely converts a value to a number, removing non-numeric characters
 */
function toNumberSafe(val) {
  if (val == null || val === "") return NaN;
  return Number(String(val).replace(/[^\d.-]/g, ""));
}

/**
 * Converts a column number to letter notation (1=A, 2=B, 27=AA, etc.)
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
 * Test function - can be run manually to verify setup
 */
function testSheetAccess() {
  try {
    const spreadsheet = SpreadsheetApp.openById(PERF_SHEET_ID);
    Logger.log('Successfully accessed spreadsheet: ' + spreadsheet.getName());
    
    let sheet = spreadsheet.getSheetByName(SHEET_NAMES.PERF_REPORT);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAMES.PERF_REPORT);
      Logger.log('Created new sheet: ' + SHEET_NAMES.PERF_REPORT);
    } else {
      Logger.log('Found existing sheet: ' + SHEET_NAMES.PERF_REPORT);
    }
    
    sheet.getRange('A1').setValue('Test successful at ' + new Date());
    return 'Test passed!';
  } catch (error) {
    Logger.log('Test failed: ' + error.toString());
    return 'Test failed: ' + error.toString();
  }
}

// ============================================================================
// SUMMARY SHEET BUILDER
// ============================================================================

/**
 * Builds the Summary sheet with all data, slicers, and chart
 * Replaces all formulas with Apps Script-generated data
 */
function buildSummarySheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = ss.getSheetByName("Summary");
    
    if (!summarySheet) {
      throw new Error("Summary sheet not found. Please create it first.");
    }
    
    Logger.log("Updating Summary sheet...");
    
    // Get source data
    const baseDataSheet = ss.getSheetByName(SHEET_NAMES.BASE_DATA);
    const fullCompSheet = ss.getSheetByName(SHEET_NAMES.FULL_COMP_HISTORY);
    const bonusSheet = ss.getSheetByName(SHEET_NAMES.BONUS_HISTORY);
    const perfReportSheet = ss.getSheetByName(SHEET_NAMES.PERF_REPORT);
    
    if (!baseDataSheet) {
      throw new Error("Base Data sheet not found. Please import Base Data first.");
    }
    
    // Get all source data
    const baseData = baseDataSheet.getDataRange().getValues();
    const baseHeaders = baseData[0];
    const baseDataRows = baseData.slice(1);
    
    const fullCompData = fullCompSheet ? fullCompSheet.getDataRange().getValues() : [];
    const fullCompHeaders = fullCompData.length > 0 ? fullCompData[0] : [];
    const fullCompRows = fullCompData.slice(1);
    
    const bonusData = bonusSheet ? bonusSheet.getDataRange().getValues() : [];
    const bonusHeaders = bonusData.length > 0 ? bonusData[0] : [];
    const bonusRows = bonusData.slice(1);
    
    const perfReportData = perfReportSheet ? perfReportSheet.getDataRange().getValues() : [];
    const perfReportHeaders = perfReportData.length > 0 ? perfReportData[0] : [];
    const perfReportRows = perfReportData.slice(1);
    
    // Get Performance Ratings sheet for AYR and H1 lookups
    const perfRatingsSheet = ss.getSheetByName(SHEET_NAMES.PERF_RATINGS);
    const perfRatingsData = perfRatingsSheet ? perfRatingsSheet.getDataRange().getValues() : [];
    const perfRatingsHeaders = perfRatingsData.length > 0 ? perfRatingsData[0] : [];
    const perfRatingsRows = perfRatingsData.slice(1);
    
    // Find column indices in Base Data
    const baseCols = {
      empId: findColumnIndex(baseHeaders, ["Emp ID", "Employee ID", "Employee Id"]),
      empName: findColumnIndex(baseHeaders, ["Emp Name", "Display name", "Display Name", "Name"]),
      startDate: findColumnIndex(baseHeaders, ["Start Date", "Start date"]),
      jobTitle: findColumnIndex(baseHeaders, ["Job title", "Job Title"]),
      jobLevel: findColumnIndex(baseHeaders, ["Job Level", "Job level", "Level"]),
      manager: findColumnIndex(baseHeaders, ["Manager"]),
      department: findColumnIndex(baseHeaders, ["Department"]),
      elt: findColumnIndex(baseHeaders, ["ELT"]),
      location: findColumnIndex(baseHeaders, ["Site", "Location"]), // Prefer Site for FX lookup
      terminationDate: findColumnIndex(baseHeaders, ["Termination date", "Termination Date"]),
      activeStatus: findColumnIndex(baseHeaders, ["Active/Inactive", "Status"]),
      baseSalary: findColumnIndex(baseHeaders, ["Base salary", "Base Salary", "Base Pay"]),
      variableType: findColumnIndex(baseHeaders, ["Variable Type"]),
      variablePct: findColumnIndex(baseHeaders, ["Variable %"]),
      noticePeriod: findColumnIndex(baseHeaders, ["Notice Period"])
    };
    
    // Find column indices in Full Comp History
    const fullCompCols = {
      empId: findColumnIndex(fullCompHeaders, ["Emp ID", "Employee ID"]),
      lastPromoDate: findColumnIndex(fullCompHeaders, ["Last Promotion Date"]),
      lastIncreaseDate: findColumnIndex(fullCompHeaders, ["Last Increase Date"]),
      lastIncreasePct: findColumnIndex(fullCompHeaders, ["Last Increase %"]),
      changeReason: findColumnIndex(fullCompHeaders, ["Change Reason"])
    };
    
    // Find column indices in Bonus History
    const bonusCols = {
      empId: findColumnIndex(bonusHeaders, ["Emp ID", "Employee ID"]),
      variableType: findColumnIndex(bonusHeaders, ["Variable Type"]),
      variablePct: findColumnIndex(bonusHeaders, ["Variable %"])
    };
    
    // Find column indices in Performance Report
    const perfCols = {
      empId: findColumnIndex(perfReportHeaders, ["employee id", "employee", "emp id", "emp id (employee)"]),
      performanceQ23: findColumnIndex(perfReportHeaders, ["How has the employee performed", "Q2&Q3", "performance"]),
      potentialQ23: findColumnIndex(perfReportHeaders, ["What is this employee's potential", "potential"]),
      promotionQ1: findColumnIndex(perfReportHeaders, ["Is this employee on track to be promoted", "promotion", "Q1 2026"])
    };
    
    // Check if promotion is in column T (20) - it's a numeric value (3=Ready, 2=Uncertain, 1=Not Ready)
    // Column T is index 19 (0-based) or column 20 (1-based)
    if (perfCols.promotionQ1 < 0 && perfReportHeaders.length > 19) {
      // Try column T directly (index 19)
      perfCols.promotionQ1 = 19;
    }
    
    // Find column indices in Performance Ratings (for AYR and H1)
    const perfRatingsCols = {
      empId: findColumnIndex(perfRatingsHeaders, ["Emp ID", "Employee ID", "Employee Id", "employee id"]),
      ayr2024: findColumnIndex(perfRatingsHeaders, ["AYR 2024", "AYR2024", "Annual Year Review 2024", "2024 Rating"]),
      h12025: findColumnIndex(perfRatingsHeaders, ["H1 2025", "H12025", "Half Year 2025", "H1 Rating"])
    };
    
    // Build lookup maps
    const baseDataMap = new Map();
    baseDataRows.forEach(row => {
      const empId = normalizeEmpId(safeCell(row, baseCols.empId));
      if (empId && safeCell(row, baseCols.activeStatus).toLowerCase() === "active") {
        baseDataMap.set(empId, row);
      }
    });
    
    const fullCompMap = new Map();
    fullCompRows.forEach(row => {
      const empId = normalizeEmpId(safeCell(row, fullCompCols.empId));
      if (empId) {
        fullCompMap.set(empId, row);
      }
    });
    
    const bonusMap = new Map();
    bonusRows.forEach(row => {
      const empId = normalizeEmpId(safeCell(row, bonusCols.empId));
      if (empId) {
        bonusMap.set(empId, row);
      }
    });
    
    const perfMap = new Map();
    perfReportRows.forEach(row => {
      const empId = normalizeEmpId(safeCell(row, perfCols.empId));
      if (empId) {
        perfMap.set(empId, row);
      }
    });
    
    // Build Performance Ratings map for AYR and H1 lookups
    const perfRatingsMap = new Map();
    perfRatingsRows.forEach(row => {
      const empId = normalizeEmpId(safeCell(row, perfRatingsCols.empId));
      if (empId) {
        perfRatingsMap.set(empId, row);
      }
    });
    
    // Build Manager Blurbs lookup map
    const managerBlurbsSheet = ss.getSheetByName("Manager Blurbs");
    const managerBlurbsMap = new Map();
    if (managerBlurbsSheet) {
      const blurbsData = managerBlurbsSheet.getDataRange().getValues();
      // Skip header row, assuming format: [Emp ID, Blurb]
      blurbsData.slice(1).forEach(row => {
        const empId = normalizeEmpId(row[0]);
        const blurb = row[1] || "";
        if (empId) {
          managerBlurbsMap.set(empId, blurb);
        }
      });
      Logger.log(`Loaded ${managerBlurbsMap.size} manager blurbs`);
    } else {
      Logger.log("Manager Blurbs sheet not found - blurbs will be empty");
    }
    
    // Build AI Readiness Category lookup map
    const aiReadinessMappingSheet = ss.getSheetByName("AI Readiness Mapping");
    const aiReadinessMap = new Map();
    if (aiReadinessMappingSheet) {
      const mappingData = aiReadinessMappingSheet.getDataRange().getValues();
      // Skip header row, assuming format: [Department/Function, AI Readiness Category]
      mappingData.slice(1).forEach(row => {
        const dept = String(row[0] || "").trim();
        const category = String(row[1] || "").trim();
        if (dept && category) {
          aiReadinessMap.set(dept, category);
        }
      });
      Logger.log(`Loaded ${aiReadinessMap.size} AI readiness mappings`);
    } else {
      Logger.log("AI Readiness Mapping sheet not found - categories will be 'Not Mapped'");
    }
    
    // Build data table
    const headerRow = [
      "Emp ID", "Emp Name", "Start Date", "Tenure", "Job Title", "Level", "Manager", 
      "Department", "ELT", "Location", "AYR 2024 Rating", "H1 2025", "Q2/Q3 Rating", 
      "Promotion", "Last Promotion Date", "Last Increase Date", "Last Increase %", 
      "Change Reason", "Base Salary (Local)", "Base Salary (USD)", "Variable Type", 
      "Variable %", "Notice Period", "Manager Blurb", "AI Readiness Category"
    ];
    
    const dataRows = [];
    const activeEmpIds = Array.from(baseDataMap.keys());
    
      activeEmpIds.forEach(empId => {
      const baseRow = baseDataMap.get(empId);
      const fullCompRow = fullCompMap.get(empId);
      const bonusRow = bonusMap.get(empId);
      const perfRow = perfMap.get(empId);
      const perfRatingsRow = perfRatingsMap.get(empId);
      
      // Emp ID
      const empIdVal = empId;
      
      // Emp Name
      const empName = safeCell(baseRow, baseCols.empName);
      
      // Start Date
      const startDateStr = safeCell(baseRow, baseCols.startDate);
      const startDate = parseDate(startDateStr);
      
      // Tenure
      const tenure = calculateTenure(startDate);
      
      // Job Title through Location
      const jobTitle = safeCell(baseRow, baseCols.jobTitle);
      const level = safeCell(baseRow, baseCols.jobLevel);
      const manager = safeCell(baseRow, baseCols.manager);
      const department = safeCell(baseRow, baseCols.department);
      const elt = safeCell(baseRow, baseCols.elt);
      // Location - use Site for FX lookup
      const location = safeCell(baseRow, baseCols.location);
      
      // AYR 2024 and H1 2025 from Performance Ratings sheet
      let ayrRating = "";
      let h1Rating = "";
      if (perfRatingsRow) {
        ayrRating = safeCell(perfRatingsRow, perfRatingsCols.ayr2024);
        h1Rating = safeCell(perfRatingsRow, perfRatingsCols.h12025);
      }
      
      // Q2/Q3 Rating and Promotion from Performance Report
      let q23Rating = "";
      let promotion = "";
      
      if (perfRow) {
        // Extract performance and potential ratings for Q2/Q3
        const perfText = safeCell(perfRow, perfCols.performanceQ23);
        const potentialText = safeCell(perfRow, perfCols.potentialQ23);
        
        if (perfText && potentialText) {
          const perfVal = getRatingValue(perfText);
          const potentialVal = getRatingValue(potentialText);
          q23Rating = mapPerformanceRating(perfVal, potentialVal);
        }
        
        // Extract promotion status - column T has numeric values: 3=Ready, 2=Uncertain, 1=Not Ready
        const promoValue = safeCell(perfRow, perfCols.promotionQ1);
        if (promoValue !== null && promoValue !== undefined && promoValue !== "") {
          const promoNum = toNumberSafe(promoValue);
          if (promoNum === 3) {
            promotion = "Ready";
          } else if (promoNum === 2) {
            promotion = "Uncertain";
          } else if (promoNum === 1) {
            promotion = "Not Ready";
          } else {
            // Fallback: try text parsing if not numeric
            const promoText = String(promoValue).toLowerCase();
            if (promoText.includes("ready") && !promoText.includes("not ready")) {
              promotion = "Ready";
            } else if (promoText.includes("uncertain")) {
              promotion = "Uncertain";
            } else {
              promotion = "Not Ready";
            }
          }
        } else {
          promotion = "Not Ready";
        }
      } else {
        promotion = "Not Ready";
      }
      
      // Last Promotion Date, Last Increase Date, Last Increase %, Change Reason
      let lastPromoDate = "";
      let lastIncreaseDate = "";
      let lastIncreasePct = "";
      let changeReason = "";
      
      if (fullCompRow) {
        lastPromoDate = safeCell(fullCompRow, fullCompCols.lastPromoDate);
        lastIncreaseDate = safeCell(fullCompRow, fullCompCols.lastIncreaseDate);
        lastIncreasePct = safeCell(fullCompRow, fullCompCols.lastIncreasePct);
        changeReason = safeCell(fullCompRow, fullCompCols.changeReason);
      }
      
      // Base Salary from Base Data
      const baseSalary = safeCell(baseRow, baseCols.baseSalary);
      const baseSalaryNum = toNumberSafe(baseSalary);
      const fxRate = getFXRate(location);
      const baseSalaryUSD = isFinite(baseSalaryNum) && fxRate ? baseSalaryNum * fxRate : "";
      
      // Variable Type and Variable % - prefer Bonus History over Base Data
      let variableType = "";
      let variablePct = "";
      
      // Try Bonus History first
      if (bonusRow) {
        variableType = safeCell(bonusRow, bonusCols.variableType);
        variablePct = safeCell(bonusRow, bonusCols.variablePct);
      }
      
      // Fallback to Base Data if not found in Bonus History
      if (!variableType && baseCols.variableType >= 0) {
        variableType = safeCell(baseRow, baseCols.variableType);
      }
      if (!variablePct && baseCols.variablePct >= 0) {
        variablePct = safeCell(baseRow, baseCols.variablePct);
      }
      
      // Notice Period - if active employee has termination date, they're on notice
      // We'll set this as value after building data rows
      const noticePeriod = "";
      
      // Manager Blurb - lookup from Manager Blurbs sheet
      const managerBlurb = managerBlurbsMap.get(empId) || "No blurb generated";
      
      // AI Readiness Category - lookup from AI Readiness Mapping sheet by department
      const aiReadinessCategory = aiReadinessMap.get(department) || "Not Mapped";
      
      dataRows.push([
        empIdVal, empName, startDate, tenure, jobTitle, level, manager, department, elt, location,
        ayrRating, h1Rating, q23Rating, promotion, lastPromoDate, lastIncreaseDate, 
        lastIncreasePct, changeReason, baseSalary, baseSalaryUSD, variableType, variablePct, noticePeriod, managerBlurb, aiReadinessCategory
      ]);
    });
    
    // Write data starting at row 20
    const startRow = 20;
    const dataRange = summarySheet.getRange(startRow, 1, dataRows.length + 1, headerRow.length);
    dataRange.setValues([headerRow, ...dataRows]);
    
    // Set Notice Period values (column 23) - return termination date from Base Data
    if (dataRows.length > 0) {
      const noticePeriodCol = 23;
      const noticePeriodValues = dataRows.map((_, i) => {
        const empId = dataRows[i][0]; // Emp ID is first column
        const baseRow = baseDataMap.get(normalizeEmpId(empId));
        if (baseRow && baseCols.terminationDate >= 0) {
          const termDate = safeCell(baseRow, baseCols.terminationDate);
          if (termDate) {
            // Parse and return as Date object
            return [parseDate(termDate)];
          }
        }
        return [""];
      });
      const noticePeriodRange = summarySheet.getRange(startRow + 1, noticePeriodCol, dataRows.length, 1);
      noticePeriodRange.setValues(noticePeriodValues);
      noticePeriodRange.setNumberFormat("dd-mmm-yy");
      
      // Manager Blurb (column 24) and AI Readiness Category (column 25) are already set as values in dataRows
      // Just apply text wrapping for better readability
      const managerBlurbCol = 24;
      const managerBlurbRange = summarySheet.getRange(startRow + 1, managerBlurbCol, dataRows.length, 1);
      managerBlurbRange.setWrap(true);
      
      const aiReadinessCol = 25;
      const aiReadinessRange = summarySheet.getRange(startRow + 1, aiReadinessCol, dataRows.length, 1);
      aiReadinessRange.setWrap(true);
    }
    
    // Format header row
    const headerRange = summarySheet.getRange(startRow, 1, 1, headerRow.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#9825ff");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontFamily("Roboto");
    headerRange.setFontSize(10);
    
    // Format data rows
    const dataRange2 = summarySheet.getRange(startRow + 1, 1, dataRows.length, headerRow.length);
    dataRange2.setFontFamily("Roboto");
    dataRange2.setFontSize(10);
    
    // Format specific columns
    const empIdCol = 1;
    const startDateCol = 3;
    const lastPromoDateCol = 15;
    const lastIncreaseDateCol = 16;
    const lastIncreasePctCol = 17;
    const baseSalaryCol = 19;
    const baseSalaryUSDCol = 20;
    const variablePctCol = 22;
    
    if (dataRows.length > 0) {
      // Emp ID as text
      summarySheet.getRange(startRow + 1, empIdCol, dataRows.length, 1).setNumberFormat("@");
      
      // Dates - format as dates, not text
      if (startDateCol > 0) {
        const startDateRange = summarySheet.getRange(startRow + 1, startDateCol, dataRows.length, 1);
        startDateRange.setNumberFormat("dd-mmm-yy");
        // Convert date strings to actual date values
        const startDateValues = startDateRange.getValues();
        const convertedDates = startDateValues.map(row => {
          if (row[0] && row[0] instanceof Date) {
            return [row[0]];
          } else if (row[0] && typeof row[0] === 'string') {
            const parsed = new Date(row[0]);
            return isNaN(parsed.getTime()) ? [""] : [parsed];
          }
          return [""];
        });
        startDateRange.setValues(convertedDates);
      }
      if (lastPromoDateCol > 0) {
        const promoDateRange = summarySheet.getRange(startRow + 1, lastPromoDateCol, dataRows.length, 1);
        promoDateRange.setNumberFormat("dd-mmm-yy");
        // Convert date strings to actual date values
        const promoDateValues = promoDateRange.getValues();
        const convertedPromoDates = promoDateValues.map(row => {
          if (row[0] && row[0] instanceof Date) {
            return [row[0]];
          } else if (row[0] && typeof row[0] === 'string' && row[0].trim()) {
            const parsed = new Date(row[0]);
            return isNaN(parsed.getTime()) ? [""] : [parsed];
          }
          return [""];
        });
        promoDateRange.setValues(convertedPromoDates);
      }
      if (lastIncreaseDateCol > 0) {
        const increaseDateRange = summarySheet.getRange(startRow + 1, lastIncreaseDateCol, dataRows.length, 1);
        increaseDateRange.setNumberFormat("dd-mmm-yy");
        // Convert date strings to actual date values
        const increaseDateValues = increaseDateRange.getValues();
        const convertedIncreaseDates = increaseDateValues.map(row => {
          if (row[0] && row[0] instanceof Date) {
            return [row[0]];
          } else if (row[0] && typeof row[0] === 'string' && row[0].trim()) {
            const parsed = new Date(row[0]);
            return isNaN(parsed.getTime()) ? [""] : [parsed];
          }
          return [""];
        });
        increaseDateRange.setValues(convertedIncreaseDates);
      }
      
      // Percentages - ensure they're numbers
      if (lastIncreasePctCol > 0) {
        const pctRange = summarySheet.getRange(startRow + 1, lastIncreasePctCol, dataRows.length, 1);
        pctRange.setNumberFormat("0.00");
        const pctValues = pctRange.getValues();
        const convertedPcts = pctValues.map(row => {
          if (row[0] === "" || row[0] == null) return [""];
          const num = toNumberSafe(row[0]);
          return isFinite(num) ? [num] : [""];
        });
        pctRange.setValues(convertedPcts);
      }
      if (variablePctCol > 0) {
        const varPctRange = summarySheet.getRange(startRow + 1, variablePctCol, dataRows.length, 1);
        varPctRange.setNumberFormat("0.00");
        const varPctValues = varPctRange.getValues();
        const convertedVarPcts = varPctValues.map(row => {
          if (row[0] === "" || row[0] == null) return [""];
          const num = toNumberSafe(row[0]);
          return isFinite(num) ? [num] : [""];
        });
        varPctRange.setValues(convertedVarPcts);
      }
      
      // Currency - ensure they're numbers
      if (baseSalaryCol > 0) {
        const salaryRange = summarySheet.getRange(startRow + 1, baseSalaryCol, dataRows.length, 1);
        salaryRange.setNumberFormat("#,##0");
        const salaryValues = salaryRange.getValues();
        const convertedSalaries = salaryValues.map(row => {
          if (row[0] === "" || row[0] == null) return [""];
          const num = toNumberSafe(row[0]);
          return isFinite(num) ? [num] : [""];
        });
        salaryRange.setValues(convertedSalaries);
      }
      if (baseSalaryUSDCol > 0) {
        const usdRange = summarySheet.getRange(startRow + 1, baseSalaryUSDCol, dataRows.length, 1);
        usdRange.setNumberFormat("$#,##0.00");
        const usdValues = usdRange.getValues();
        const convertedUSD = usdValues.map(row => {
          if (row[0] === "" || row[0] == null) return [""];
          const num = toNumberSafe(row[0]);
          return isFinite(num) ? [num] : [""];
        });
        usdRange.setValues(convertedUSD);
      }
    }
    
    // Add borders to all data rows (from row 21) - #d9d2e9
    if (dataRows.length > 0) {
      const dataRowsRange = summarySheet.getRange(startRow + 1, 1, dataRows.length, headerRow.length);
      dataRowsRange.setBorder(true, true, true, true, true, true, "#d9d2e9", SpreadsheetApp.BorderStyle.SOLID);
    }
    
    // Add helper column in column AH (34) for SUBTOTAL - based on Employee ID (column A), not Rating
    const helperCol = 34;
    if (dataRows.length > 0) {
      const helperFormulas = dataRows.map((_, i) => [`=SUBTOTAL(103,$A${startRow + 1 + i})`]);
      const helperRange = summarySheet.getRange(startRow + 1, helperCol, dataRows.length, 1);
      helperRange.setFormulas(helperFormulas);
      
      // Format helper column: middle aligned, Roboto 10, #cccccc
      helperRange.setVerticalAlignment("middle");
      helperRange.setFontFamily("Roboto");
      helperRange.setFontSize(10);
      helperRange.setFontColor("#cccccc");
      
      // Update chart data formulas - only update the range references
      updateChartDataFormulas(summarySheet, dataRows.length, startRow, helperCol);
      
      // Update slicer ranges to match new data size
      updateSlicerRanges(summarySheet, dataRows.length, startRow - 1);
    }
    
    Logger.log(`Summary sheet updated successfully with ${dataRows.length} employees`);
    SpreadsheetApp.getUi().alert("Success", `Summary sheet updated successfully with ${dataRows.length} employees`, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`Error building Summary sheet: ${error.message}`);
    SpreadsheetApp.getUi().alert("Error", `Failed to build Summary sheet: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    throw error;
  }
}

/**
 * Helper function to find column index by header name
 */
function findColumnIndex(headers, possibleNames) {
  if (!headers || headers.length === 0) return -1;
  const normalizedHeaders = headers.map(h => String(h).toLowerCase().trim());
  for (const name of possibleNames) {
    const normalizedName = name.toLowerCase().trim();
    const index = normalizedHeaders.findIndex(h => h.includes(normalizedName) || normalizedName.includes(h));
    if (index >= 0) return index;
  }
  return -1;
}

/**
 * Normalize employee ID
 */
function normalizeEmpId(empId) {
  if (!empId) return "";
  const num = toNumberSafe(empId);
  return isFinite(num) ? String(num) : String(empId).trim();
}

/**
 * Parse date string to Date object
 */
function parseDate(dateStr) {
  if (!dateStr) return "";
  try {
    return new Date(dateStr);
  } catch (e) {
    return "";
  }
}

/**
 * Calculate tenure from start date
 */
function calculateTenure(startDate) {
  if (!startDate || !(startDate instanceof Date)) return "";
  const today = new Date();
  const daysDiff = Math.floor((today - startDate) / (1000 * 60 * 60 * 24));
  
  if (daysDiff >= 1460) return "4 Years+";
  if (daysDiff >= 1095) return "3 Years";
  if (daysDiff >= 730) return "2 Years";
  if (daysDiff >= 545) return "1.5 Years";
  if (daysDiff >= 365) return "1 Year";
  if (daysDiff >= 180) return "6 Months";
  return "Less than 6 Months";
}

/**
 * Get rating value from text (Exceeds=3, Meets=2, Low=1)
 */
function getRatingValue(text) {
  if (!text) return 0;
  const lower = text.toLowerCase();
  if (lower.includes("exceed") || lower === "3") return 3;
  if (lower.includes("meet") || lower === "2") return 2;
  if (lower.includes("low") || lower === "1") return 1;
  return 0;
}

/**
 * Map performance and potential to rating (HH, HM, etc.)
 */
function mapPerformanceRating(perfVal, potentialVal) {
  const combined = String(perfVal) + String(potentialVal);
  const mapping = {
    "33": "HH",
    "32": "HM",
    "31": "HL",
    "23": "MH",
    "22": "MM",
    "21": "ML",
    "13": "NI",
    "12": "NI",
    "11": "NI"
  };
  return mapping[combined] || "";
}

/**
 * Get FX rate for location
 */
function getFXRate(location) {
  if (!location) return 1;
  const locationLower = location.toLowerCase();
  if (locationLower.includes("india")) return 0.0125;
  if (locationLower.includes("usa") || locationLower.includes("united states")) return 1;
  if (locationLower.includes("uk") || locationLower.includes("united kingdom")) return 1.37;
  return 1;
}

/**
 * Update existing slicer ranges on the Summary sheet
 * Uses the simple SpreadsheetApp approach with "reset then set" workaround
 * This preserves slicer formatting and position, only updating the data range
 */
function updateSlicerRanges(sheet, numRows, dataStartRow) {
  try {
    Logger.log(`=== Starting Slicer Range Update ===`);
    Logger.log(`Sheet: ${sheet.getName()}`);
    Logger.log(`Updating slicer ranges: ${numRows} rows starting at row ${dataStartRow + 1}`);
    Logger.log(`New range will be: Row ${dataStartRow + 1} to ${dataStartRow + numRows + 1}`);
    
    // Get all slicers on this sheet
    const slicers = sheet.getSlicers();
    
    if (slicers.length === 0) {
      Logger.log("⚠️ No slicers found. Skipping slicer range update.");
      SpreadsheetApp.getUi().alert(
        "No Slicers Found",
        "No slicers were found on this sheet. Please create slicers first before updating their ranges.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    Logger.log(`Found ${slicers.length} slicers on sheet`);
    
    // Define the new range for slicers
    // From row dataStartRow+1 (header row) to the last data row, all columns
    const lastRow = dataStartRow + numRows + 1;
    const lastColumn = sheet.getLastColumn();
    const newRange = sheet.getRange(dataStartRow + 1, 1, numRows + 1, lastColumn);
    
    Logger.log(`New range: ${newRange.getA1Notation()}`);
    
    // Workaround for "common rows" error:
    // Reset all slicers to a tiny range first, then set to desired range
    const resetRange = sheet.getRange(1, 1, 1, 1); // Single cell outside slicer coverage
    
    Logger.log("Step 1: Resetting all slicers to A1...");
    slicers.forEach((slicer, idx) => {
      try {
        slicer.setRange(resetRange);
        Logger.log(`  ✓ Reset slicer ${idx + 1}`);
      } catch (e) {
        Logger.log(`  ⚠️ Could not reset slicer ${idx + 1}: ${e.message}`);
      }
    });
    
    Logger.log("Step 2: Setting all slicers to new range...");
    slicers.forEach((slicer, idx) => {
      try {
        slicer.setRange(newRange);
        Logger.log(`  ✓ Updated slicer ${idx + 1} to ${newRange.getA1Notation()}`);
      } catch (e) {
        Logger.log(`  ⚠️ Could not update slicer ${idx + 1}: ${e.message}`);
        throw e;
      }
    });
    
    Logger.log(`✓ Successfully updated ${slicers.length} slicer ranges`);
    
    SpreadsheetApp.getUi().alert(
      "Success",
      `Updated ${slicers.length} slicer ranges to:\n${newRange.getA1Notation()}\n\n` +
      `(Rows ${dataStartRow + 1} to ${lastRow})`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log(`❌ Error updating slicer ranges: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    
    SpreadsheetApp.getUi().alert(
      "Error Updating Slicers",
      `Failed to update slicer ranges: ${error.message}\n\n` +
      `This can happen if slicers have different ranges. Try recreating the slicers.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    throw error;
  }
}

/**
 * Create slicers on the Summary sheet using Sheets API
 */
function createSlicers(sheet, numRows, dataStartRow) {
  try {
    const spreadsheetId = sheet.getParent().getId();
    const sheetId = sheet.getSheetId();
    
    // Slicer configuration: {col, sourceRange, title, color, position}
    // Column mapping: A=1 (Emp ID), B=2 (Emp Name), C=3 (Start Date), D=4 (Tenure), E=5 (Job Title), 
    // F=6 (Level), G=7 (Manager), H=8 (Department), I=9 (ELT), J=10 (Location), K=11 (AYR), 
    // L=12 (H1), M=13 (Q2/Q3), N=14 (Promotion), etc.
    const slicers = [
      {col: 9, range: `I${dataStartRow + 1}:I${dataStartRow + numRows}`, title: "ELT", color: "#9825ff", row: 4, colPos: 1},
      {col: 6, range: `F${dataStartRow + 1}:F${dataStartRow + numRows}`, title: "Level", color: "#9825ff", row: 4, colPos: 2},
      {col: 4, range: `D${dataStartRow + 1}:D${dataStartRow + numRows}`, title: "Tenure", color: "#9825ff", row: 4, colPos: 3},
      {col: 8, range: `H${dataStartRow + 1}:H${dataStartRow + numRows}`, title: "Department", color: "#9825ff", row: 4, colPos: 4},
      {col: 11, range: `K${dataStartRow + 1}:K${dataStartRow + numRows}`, title: "AYR 2024 Rating", color: "#9825ff", row: 4, colPos: 5},
      {col: 21, range: `U${dataStartRow + 1}:U${dataStartRow + numRows}`, title: "Variable Type", color: "#9825ff", row: 4, colPos: 6},
      {col: 7, range: `G${dataStartRow + 1}:G${dataStartRow + numRows}`, title: "Manager", color: "#9825ff", row: 4, colPos: 7},
      {col: 12, range: `L${dataStartRow + 1}:L${dataStartRow + numRows}`, title: "H1 2025", color: "#9825ff", row: 4, colPos: 8},
      {col: 10, range: `J${dataStartRow + 1}:J${dataStartRow + numRows}`, title: "Location", color: "#9825ff", row: 4, colPos: 9},
      {col: 13, range: `M${dataStartRow + 1}:M${dataStartRow + numRows}`, title: "Q2/Q3 Rating", color: "#ff9901", row: 4, colPos: 10},
      {col: 14, range: `N${dataStartRow + 1}:N${dataStartRow + numRows}`, title: "Promotion", color: "#ff9901", row: 4, colPos: 11},
      {col: 23, range: `W${dataStartRow + 1}:W${dataStartRow + numRows}`, title: "Notice Period", color: "#ff0100", row: 4, colPos: 12}
    ];
    
    const requests = [];
    
    slicers.forEach((slicer, idx) => {
      // Calculate position for slicer (starting at row 4, column based on colPos)
      const startRow = slicer.row - 1; // 0-indexed
      const startCol = slicer.colPos - 1; // 0-indexed
      const endRow = startRow + 10; // Height of slicer
      const endCol = startCol; // Width of slicer
      
      const slicerRequest = {
        addSlicer: {
          slicer: {
            title: slicer.title,
            sourceRange: {
              sheetId: sheetId,
              startRowIndex: dataStartRow, // Header row
              endRowIndex: dataStartRow + numRows,
              startColumnIndex: slicer.col - 1,
              endColumnIndex: slicer.col
            },
            position: {
              overlayPosition: {
                anchorCell: {
                  sheetId: sheetId,
                  rowIndex: startRow,
                  columnIndex: startCol
                },
                offsetXPixels: 0,
                offsetYPixels: 0,
                widthPixels: 120,
                heightPixels: 200
              }
            },
            style: {
              theme: "LIGHT",
              headerStyle: {
                backgroundColor: {
                  red: parseInt(slicer.color.substring(1, 3), 16) / 255,
                  green: parseInt(slicer.color.substring(3, 5), 16) / 255,
                  blue: parseInt(slicer.color.substring(5, 7), 16) / 255
                },
                textFormat: {
                  foregroundColor: {red: 1, green: 1, blue: 1},
                  fontFamily: "Roboto",
                  fontSize: 14,
                  bold: true
                }
              }
            }
          }
        }
      };
      
      requests.push(slicerRequest);
    });
    
    if (requests.length > 0) {
      try {
        Sheets.Spreadsheets.batchUpdate({
          requests: requests
        }, spreadsheetId);
        Logger.log(`Created ${requests.length} slicers successfully`);
      } catch (apiError) {
        Logger.log(`Sheets API error creating slicers: ${apiError.message}`);
        // Fallback to placeholder cells
        createSlicerPlaceholders(sheet, slicers);
      }
    }
  } catch (error) {
    Logger.log(`Error creating slicers: ${error.message}`);
    // Fallback to placeholder cells - using correct column positions
    const slicers = [
      {row: 4, col: 1, title: "ELT", color: "#9825ff"},
      {row: 4, col: 2, title: "Level", color: "#9825ff"},
      {row: 4, col: 3, title: "Tenure", color: "#9825ff"},
      {row: 4, col: 4, title: "Department", color: "#9825ff"},
      {row: 4, col: 5, title: "AYR 2024 Rating", color: "#9825ff"},
      {row: 4, col: 6, title: "Variable Type", color: "#9825ff"},
      {row: 4, col: 7, title: "Manager", color: "#9825ff"},
      {row: 4, col: 8, title: "H1 2025", color: "#9825ff"},
      {row: 4, col: 9, title: "Location", color: "#9825ff"},
      {row: 4, col: 10, title: "Q2/Q3 Rating", color: "#ff9901"},
      {row: 4, col: 11, title: "Promotion", color: "#ff9901"},
      {row: 4, col: 12, title: "Notice Period", color: "#ff0100"}
    ];
    createSlicerPlaceholders(sheet, slicers);
  }
}

/**
 * Fallback: Create placeholder cells for slicers
 */
function createSlicerPlaceholders(sheet, slicers) {
  slicers.forEach((slicer) => {
    const cell = sheet.getRange(slicer.row, slicer.col);
    cell.setValue(`All - ${slicer.title}`);
    cell.setFontFamily("Roboto");
    cell.setFontSize(14);
    cell.setBackground(slicer.color);
    cell.setFontColor("#ffffff");
    cell.setFontWeight("bold");
  });
}

/**
 * Update chart data formulas - only update range references in SUMPRODUCT
 */
function updateChartDataFormulas(sheet, numDataRows, dataStartRow, helperCol) {
  // Calculate the last data row
  const lastDataRow = dataStartRow + numDataRows;
  
  // Chart data is in rows 2-7 (I2:I7), column I (9)
  // Only update the existing formulas to use the correct last row
  const chartRows = [2, 3, 4, 5, 6, 7]; // HH, HM, MH, MM, ML, NI
  
  chartRows.forEach(row => {
    const cell = sheet.getRange(row, 9); // Column I
    const currentFormula = cell.getFormula();
    
    if (currentFormula && currentFormula.includes("SUMPRODUCT")) {
      // Replace the range references with the correct last row
      // Pattern: ($AH$21:$AH$XXX=1)*($M$21:$M$XXX=H2)
      const updatedFormula = currentFormula.replace(
        /\$AH\$\d+:\$AH\$\d+/g, 
        `$AH$${dataStartRow + 1}:$AH$${lastDataRow}`
      ).replace(
        /\$M\$\d+:\$M\$\d+/g,
        `$M$${dataStartRow + 1}:$M$${lastDataRow}`
      );
      
      cell.setFormula(updatedFormula);
    }
  });
  
  // Update employee count in I17 if it exists
  const countCell = sheet.getRange(17, 9);
  const countFormula = countCell.getFormula();
  if (countFormula && countFormula.includes("COUNTIF")) {
    const updatedCountFormula = countFormula.replace(
      /AH\$?\d+:AH\$?\d+/g,
      `AH$${dataStartRow + 1}:AH$${lastDataRow}`
    );
    countCell.setFormula(updatedCountFormula);
  }
}

/**
 * Create rating distribution chart using Sheets API
 */
function createRatingDistributionChart(sheet, chartStartRow, chartStartCol, numRatings, dataStartRow) {
  try {
    const spreadsheetId = sheet.getParent().getId();
    const sheetId = sheet.getSheetId();
    
    // Color mapping for ratings (RGB values 0-1)
    const ratingColors = {
      "HH": {red: 78/255, green: 167/255, blue: 46/255}, // #4ea72e
      "HM": {red: 152/255, green: 37/255, blue: 255/255}, // #9825ff
      "MH": {red: 152/255, green: 37/255, blue: 255/255}, // #9825ff
      "MM": {red: 152/255, green: 37/255, blue: 255/255}, // #9825ff
      "ML": {red: 255/255, green: 153/255, blue: 1/255}, // #ff9901
      "NI": {red: 244/255, green: 199/255, blue: 195/255} // #f4c7c3
    };
    
    // Create a single series for the count data
    const series = [{
      series: {
        sourceRange: {
          sheets: [{
            sourceRange: {
              sheetId: sheetId,
              startRowIndex: chartStartRow,
              endRowIndex: chartStartRow + 1 + numRatings,
              startColumnIndex: chartStartCol, // Count column (G)
              endColumnIndex: chartStartCol + 1
            }
          }]
        }
      },
      targetAxis: "LEFT_AXIS",
      type: "COLUMN"
    }];
    
    // Chart position - overlay on the data area (rows 1-15, columns F-T)
    const chartPosition = {
      overlayPosition: {
        anchorCell: {
          sheetId: sheetId,
          rowIndex: 0, // Row 1
          columnIndex: chartStartCol - 1 // Column F
        },
        offsetXPixels: 0,
        offsetYPixels: 0,
        widthPixels: 800,
        heightPixels: 400
      }
    };
    
    // Create the chart
    const chartRequest = {
      addChart: {
        chart: {
          spec: {
            title: "Rating Distribution",
            basicChart: {
              chartType: "COLUMN",
              legendPosition: "BOTTOM_LEGEND",
              axis: [
                {
                  position: "BOTTOM_AXIS",
                  title: "Rating"
                },
                {
                  position: "LEFT_AXIS",
                  title: "Count"
                }
              ],
              domains: [{
                domain: {
                  sourceRange: {
                    sheets: [{
                      sourceRange: {
                        sheetId: sheetId,
                        startRowIndex: chartStartRow,
                        endRowIndex: chartStartRow + 1 + numRatings,
                        startColumnIndex: chartStartCol - 1, // Rating column (F)
                        endColumnIndex: chartStartCol
                      }
                    }]
                  }
                }
              }],
              series: series,
              headerCount: 1
            }
          },
          position: chartPosition
        }
      }
    };
    
    try {
      Sheets.Spreadsheets.batchUpdate({
        requests: [chartRequest]
      }, spreadsheetId);
      Logger.log("Chart created successfully");
    } catch (apiError) {
      Logger.log(`Sheets API error creating chart: ${apiError.message}`);
      Logger.log("Note: Make sure Sheets API is enabled in Apps Script project");
      Logger.log("Go to: Resources > Advanced Google services > Enable Google Sheets API");
      // Chart creation will be skipped, but data is still available
    }
  } catch (error) {
    Logger.log(`Error creating chart: ${error.message}`);
    if (error.message.includes("Sheets is not defined")) {
      Logger.log("Sheets API not available. Enable it in: Resources > Advanced Google services");
    }
    // Chart creation will be skipped, but data is still available
  }
}

/**
 * Apply conditional formatting to the Summary sheet
 */
function applyConditionalFormatting(sheet, numRows, dataStartRow) {
  const ratingCol = 13; // Q2/Q3 Rating column
  const promotionCol = 14; // Promotion column
  
  // Format HH ratings (green #b7d7a8)
  const hhRange = sheet.getRange(dataStartRow + 1, ratingCol, numRows, 1);
  const hhRule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([hhRange])
    .whenTextEqualTo("HH")
    .setBackground("#b7d7a8")
    .build();
  
  // Format Promotion Ready (blue #92d1ea)
  const promoReadyRange = sheet.getRange(dataStartRow + 1, promotionCol, numRows, 1);
  const promoReadyRule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([promoReadyRange])
    .whenTextEqualTo("Ready")
    .setBackground("#92d1ea")
    .build();
  
  // Format NI ratings (red #f4c7c3)
  const niRange = sheet.getRange(dataStartRow + 1, ratingCol, numRows, 1);
  const niRule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([niRange])
    .whenTextEqualTo("NI")
    .setBackground("#f4c7c3")
    .build();
  
  const rules = [hhRule, promoReadyRule, niRule];
  sheet.setConditionalFormatRules(rules);
}

