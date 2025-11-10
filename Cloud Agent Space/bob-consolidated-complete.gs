/**
 * ================================
 * HiBob Data Updater - CONSOLIDATED
 * Version: 2.0
 * Date: 2025-10-29
 * ================================
 *
 * This consolidated file combines:
 * - CONFIG: Global constants and configuration
 * - UTILITIES: Shared helper functions
 * - MAIN: Core functionality (Fields, Lists, Employees, Uploader)
 * - HISTORY: History table uploader
 * - SALARY DATA: Salary data imports
 * - PERFORMANCE: Performance report automation
 * - DOCUMENTATION: Comprehensive guide generator
 */

// ============================================================================
// CONFIG - Global Configuration
// ============================================================================

/** ----------- GLOBAL CONFIG ------------- */
const CONFIG = Object.freeze({
  /** HiBob API base (switch to sandbox if needed) */
  HIBOB_BASE_URL: 'https://api.hibob.com',

  /** Sheet names */
  META_SHEET: 'Bob Fields Meta Data',
  LISTS_SHEET: 'Bob Lists',
  EMPLOYEES_SHEET: 'Employees',
  UPDATES_SHEET: 'Uploader',
  HISTORY_SHEET: 'History Uploader',
  DOCS_SHEET: 'Bob Updater Guide',
  
  /** Salary Data Sheets */
  BASE_DATA: 'Base Data',
  BONUS_HISTORY: 'Bonus History',
  COMP_HISTORY: 'Comp History',
  FULL_COMP_HISTORY: 'Full Comp History',
  PERF_REPORT: 'Bob Perf Report',

  /** API Rate Limits */
  PUTS_PER_MINUTE: 10,
  RETRY_BACKOFF_MS: 300,
  
  /** Logging & Debugging */
  LOG_VERBOSE: true,
  
  /** Employee Search Settings */
  DEFAULT_EMPLOYMENT_STATUS: 'Active',
  SEARCH_HUMANREADABLE: 'REPLACE',
  
  /** UI Colors */
  COLORS: {
    HEADER: '#4285F4',
    HEADER_TEXT: '#FFFFFF',
    INPUT_REQUIRED: '#FFF3CD',
    INPUT_OPTIONAL: '#F8F9FA',
    SUCCESS: '#D9EAD3',
    WARNING: '#FFF2CC',
    ERROR: '#F4CCCC',
    INFO: '#CFE2F3',
    SECTION_HEADER: '#E8EAED'
  }
});

/** ----------- Salary Data Report IDs ------------- */
const BOB_REPORT_IDS = {
  BASE_DATA: "31048356",
  BONUS_HISTORY: "31054302",
  COMP_HISTORY: "31054312",
  FULL_COMP_HISTORY: "31168524"
};

const WRITE_COLS = 23; // Column W - limit for Base Data sheet
const ALLOWED_EMP_TYPES = new Set(["Permanent", "Regular Full-Time"]);

/** ----------- Performance Reports ------------- */
const PERF_SHEET_ID = '1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA';

/** ----------- Batch Upload Constants ------------- */
const BATCH_SIZE = 45;
const TRIGGER_INTERVAL = 5;
const MAX_EXECUTION_TIME = 330000;
const PUT_DELAY_MS = Math.ceil(60000 / CONFIG.PUTS_PER_MINUTE);

/** ----------- Aliases for convenience ------------- */
const BASE = CONFIG.HIBOB_BASE_URL;
const SHEET_FIELDS = CONFIG.META_SHEET;
const SHEET_LISTS = CONFIG.LISTS_SHEET;
const SHEET_EMPLOYEES = CONFIG.EMPLOYEES_SHEET;
const SHEET_UPLOADER = CONFIG.UPDATES_SHEET;

// ============================================================================
// CREDENTIAL MANAGEMENT
// ============================================================================

/**
 * Store HiBob service user credentials
 * 
 * @param {string} id - Service user ID/email
 * @param {string} token - Service user API token
 */
function setBobServiceUser(id, token) {
  if (!id || !token) {
    throw new Error('Both service user ID and token are required.');
  }
  
  const props = PropertiesService.getScriptProperties();
  props.setProperty('HIBOB_SERVICE_USER_ID', id);
  props.setProperty('HIBOB_SERVICE_USER_TOKEN', token);
  
  Logger.log('✅ HiBob service user credentials saved successfully.');
  Logger.log(`   User ID: ${id}`);
  Logger.log('   Token: ••••••••• (hidden)');
  
  try {
    const { auth } = getCreds_();
    Logger.log('✅ Credentials validated and Base64 auth header generated.');
  } catch (e) {
    Logger.log('⚠️ Warning: Could not validate credentials: ' + e.message);
  }
}

/**
 * Retrieve stored credentials and ready-to-use Base64 Authorization header
 *
 * @returns {Object} { id, token, auth }
 * @throws {Error} If credentials are not configured
 */
function getCreds_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('HIBOB_SERVICE_USER_ID');
  const token = props.getProperty('HIBOB_SERVICE_USER_TOKEN');

  if (!id || !token) {
    throw new Error(
      '❌ Missing HiBob credentials.\n\n' +
      'Please run: setBobServiceUser("<SERVICE_USER_ID>", "<SERVICE_USER_TOKEN>")\n\n' +
      'To get your credentials:\n' +
      '1. Log in to HiBob as an admin\n' +
      '2. Go to Settings > API > Service Users\n' +
      '3. Create or copy existing service user credentials'
    );
  }

  const auth = Utilities.base64Encode(`${id}:${token}`);
  
  if (CONFIG.LOG_VERBOSE) {
    Logger.log(`🔐 Using HiBob service user: ${id}`);
  }

  return { id, token, auth };
}

/**
 * Reset stored credentials
 */
function resetBobCredentials() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('HIBOB_SERVICE_USER_ID');
  props.deleteProperty('HIBOB_SERVICE_USER_TOKEN');
  Logger.log('🧹 Cleared stored HiBob credentials.');
  Logger.log('Run setBobServiceUser() to set new credentials.');
}

/**
 * View current credentials (token partially hidden)
 */
function viewBobCredentials() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('HIBOB_SERVICE_USER_ID');
  const token = props.getProperty('HIBOB_SERVICE_USER_TOKEN');
  
  Logger.log('📋 Current HiBob Credentials:');
  Logger.log({
    'Service User ID': id || '(not set)',
    'Service User Token': token ? '••••••••••••' + token.slice(-4) + ' (hidden)' : '(not set)',
    'Status': (id && token) ? '✅ Configured' : '❌ Not configured'
  });
  
  if (!id || !token) {
    Logger.log('\n⚠️ Credentials not configured. Run setBobServiceUser() first.');
  }
}

/**
 * Test API connection with current credentials
 */
function testBobConnection() {
  try {
    const { auth, id } = getCreds_();
    Logger.log(`🧪 Testing connection with service user: ${id}`);
    
    const url = `${CONFIG.HIBOB_BASE_URL}/v1/company/people/fields`;
    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      headers: { 
        Authorization: `Basic ${auth}`, 
        Accept: 'application/json' 
      }
    });
    
    const code = resp.getResponseCode();
    
    if (code === 200) {
      Logger.log('✅ Connection successful! API responded with 200 OK.');
      const data = JSON.parse(resp.getContentText());
      Logger.log(`📊 Retrieved ${Array.isArray(data) ? data.length : 'unknown'} fields from HiBob.`);
    } else if (code === 401) {
      Logger.log('❌ Authentication failed (401 Unauthorized).');
      Logger.log('Please check your service user credentials.');
    } else if (code === 403) {
      Logger.log('❌ Access forbidden (403). Service user may lack permissions.');
    } else {
      Logger.log(`⚠️ Unexpected response: HTTP ${code}`);
      Logger.log(resp.getContentText().slice(0, 200));
    }
    
  } catch (e) {
    Logger.log('❌ Connection test failed: ' + e.message);
  }
}

/**
 * Legacy function - Get Bob API credentials (for backward compatibility)
 * Maps to new credential system
 */
function getBobCredentials() {
  const creds = getCreds_();
  return { id: creds.id, key: creds.token };
}

// ============================================================================
// UTILITIES - Shared Helper Functions
// ============================================================================

/**
 * Get existing sheet or create if not exists
 * @param {string} sheetName - Name of the sheet
 * @returns {Sheet} The sheet object
 */
function getOrCreateSheet_(sheetName) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  return sheet;
}

/**
 * Safely create filter on a sheet (ignores errors if filter exists)
 * @param {Sheet} sheet - The sheet to add filter to
 */
function safeCreateFilter_(sheet) {
  try {
    if (sheet.getFilter()) {
      sheet.getFilter().remove();
    }
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow > 0 && lastCol > 0) {
      sheet.getRange(1, 1, lastRow, lastCol).createFilter();
    }
  } catch (e) {
    Logger.log('Note: Could not create filter - ' + e.message);
  }
}

/**
 * Auto-fit ALL columns for better readability
 * @param {Sheet} sheet - The sheet to auto-fit
 */
function autoFitAllColumns_(sheet) {
  try {
    const lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      for (let i = 1; i <= lastCol; i++) {
        sheet.autoResizeColumn(i);
        // Add a bit of padding
        const currentWidth = sheet.getColumnWidth(i);
        sheet.setColumnWidth(i, Math.min(currentWidth + 20, 500));
      }
    }
  } catch (e) {
    Logger.log('Note: Could not auto-fit columns - ' + e.message);
  }
}

/**
 * Auto-fit specific columns
 * @param {Sheet} sheet - The sheet to auto-fit
 * @param {number} numCols - Number of columns to auto-fit
 */
function autoFit_(sheet, numCols) {
  try {
    for (let i = 1; i <= numCols; i++) {
      sheet.autoResizeColumn(i);
      // Add padding
      const currentWidth = sheet.getColumnWidth(i);
      sheet.setColumnWidth(i, Math.min(currentWidth + 20, 500));
    }
  } catch (e) {
    Logger.log('Note: Could not auto-fit columns - ' + e.message);
  }
}

/**
 * Format header row with consistent styling
 * @param {Sheet} sheet - The sheet
 * @param {number} row - Header row number
 * @param {number} numCols - Number of columns in header
 */
function formatHeaderRow_(sheet, row, numCols) {
  const headerRange = sheet.getRange(row, 1, 1, numCols);
  headerRange
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setFontWeight('bold')
    .setFontSize(11);
}

/**
 * Format a cell as a required input field
 * @param {Range} range - The cell range
 * @param {string} placeholder - Optional placeholder text
 */
function formatRequiredInput_(range, placeholder) {
  range
    .setBackground(CONFIG.COLORS.INPUT_REQUIRED)
    .setBorder(true, true, true, true, false, false, '#F9CB9C', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  if (placeholder) {
    range.setNote('REQUIRED: ' + placeholder);
  }
}

/**
 * Format a cell as an optional input field
 * @param {Range} range - The cell range
 * @param {string} note - Optional note text
 */
function formatOptionalInput_(range, note) {
  range
    .setBackground(CONFIG.COLORS.INPUT_OPTIONAL)
    .setBorder(true, true, true, true, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  
  if (note) {
    range.setNote('OPTIONAL: ' + note);
  }
}

/**
 * Format section header
 * @param {Range} range - The range to format
 * @param {string} text - Header text
 */
function formatSectionHeader_(range, text) {
  range
    .setValue(text)
    .setBackground(CONFIG.COLORS.SECTION_HEADER)
    .setFontWeight('bold')
    .setFontSize(12)
    .setBorder(false, false, true, false, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

/**
 * Safely convert value to string
 * @param {*} val - Value to convert
 * @returns {string} String representation or empty string
 */
function safe(val) {
  if (val === null || val === undefined) return '';
  return String(val);
}

/**
 * Convert boolean to string or return blank
 * @param {*} val - Value to convert
 * @returns {string} 'true', 'false', or ''
 */
function boolOrBlank(val) {
  if (val === true) return 'true';
  if (val === false) return 'false';
  return '';
}

/**
 * Convert object to JSON string or return blank
 * @param {*} obj - Object to stringify
 * @returns {string} JSON string or ''
 */
function jsonOrBlank(obj) {
  if (!obj) return '';
  try {
    return JSON.stringify(obj);
  } catch (e) {
    return '';
  }
}

/**
 * Normalize blank/null values to empty string
 * @param {*} val - Value to normalize
 * @returns {string} The value or empty string
 */
function normalizeBlank_(val) {
  if (val === null || val === undefined || val === '') return '';
  const str = String(val).trim();
  return str.toLowerCase() === 'null' ? '' : str;
}

/**
 * Try to parse JSON string safely
 * @param {string} str - JSON string to parse
 * @returns {Object|null} Parsed object or null
 */
function tryParseJson_(str) {
  if (!str) return null;
  try {
    return JSON.parse(str);
  } catch (e) {
    return null;
  }
}

/**
 * Show toast notification
 * @param {string} msg - Message to display
 * @param {string} title - Optional title
 * @param {number} timeout - Optional timeout in seconds
 */
function toast_(msg, title, timeout) {
  SpreadsheetApp.getActive().toast(msg, title || 'Bob Updater', timeout || 5);
}

/**
 * Read fields metadata from Bob Fields Meta Data sheet
 * @returns {Array} Array of field objects
 */
function readFields_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.META_SHEET);
  if (!sh) {
    throw new Error('Sheet "' + CONFIG.META_SHEET + '" not found. Run "1. Pull Fields" first.');
  }
  
  const data = sh.getDataRange().getValues();
  if (data.length < 3) {
    throw new Error('Sheet "' + CONFIG.META_SHEET + '" is empty. Run "1. Pull Fields" first.');
  }
  
  // Header is on row 2 (index 1), row 1 (index 0) is the title
  const header = data[1].map(function(x) { return String(x || '').trim(); });
  const iId = header.indexOf('id');
  const iName = header.indexOf('name');
  const iPath = header.indexOf('jsonPath');
  const iType = header.indexOf('type');
  const iCalc = header.indexOf('calculated');
  const iTypeData = header.indexOf('typeData (raw JSON)');
  
  const fields = [];
  // Data starts at row 3 (index 2)
  for (var r = 2; r < data.length; r++) {
    var row = data[r];
    var typeData = null;
    if (iTypeData >= 0) {
      typeData = tryParseJson_(row[iTypeData]);
    }
    
    fields.push({
      id: iId >= 0 ? safe(row[iId]) : '',
      name: iName >= 0 ? safe(row[iName]) : '',
      jsonPath: iPath >= 0 ? safe(row[iPath]) : '',
      type: iType >= 0 ? safe(row[iType]) : '',
      calculated: iCalc >= 0 ? safe(row[iCalc]) : '',
      typeData: typeData || {}
    });
  }
  
  return fields;
}

/**
 * Build mapping of CIQ ID to Bob ID from Employees sheet
 * @returns {Object} Map of CIQ -> Bob ID
 */
function buildCiqToBobMap_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.EMPLOYEES_SHEET);
  if (!sh) {
    throw new Error('Sheet "' + CONFIG.EMPLOYEES_SHEET + '" not found. Run "3. Pull Employees" first.');
  }
  
  const data = sh.getDataRange().getValues();
  if (data.length < 8) return {};
  
  // Header is at row 7 (index 6)
  const header = data[6].map(function(x) { return String(x || '').trim(); });
  const iCiq = header.indexOf('CIQ ID');
  const iBob = header.indexOf('Bob ID');
  
  if (iCiq < 0 || iBob < 0) {
    throw new Error('Could not find "CIQ ID" or "Bob ID" columns in Employees sheet.');
  }
  
  const map = {};
  for (var r = 7; r < data.length; r++) {
    const ciq = String(data[r][iCiq] || '').trim();
    const bob = String(data[r][iBob] || '').trim();
    if (ciq && bob) {
      map[ciq] = bob;
    }
  }
  
  return map;
}

/**
 * Build reverse mapping of Bob ID to CIQ ID
 * @returns {Object} Map of Bob ID -> CIQ ID
 */
function buildBobToCiqMap_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.EMPLOYEES_SHEET);
  if (!sh) return {};
  
  const data = sh.getDataRange().getValues();
  if (data.length < 8) return {};
  
  const header = data[6].map(function(x) { return String(x || '').trim(); });
  const iCiq = header.indexOf('CIQ ID');
  const iBob = header.indexOf('Bob ID');
  
  if (iCiq < 0 || iBob < 0) return {};
  
  const map = {};
  for (var r = 7; r < data.length; r++) {
    const ciq = String(data[r][iCiq] || '').trim();
    const bob = String(data[r][iBob] || '').trim();
    if (ciq && bob) {
      map[bob] = ciq;
    }
  }
  
  return map;
}

/**
 * Build mapping of list label to ID
 * @param {string} listName - Name of the list
 * @returns {Object} Map of label -> ID (both original and lowercase)
 */
function buildListLabelToId_(listName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.LISTS_SHEET);
  if (!sh) return {};
  
  const vals = sh.getDataRange().getValues();
  if (vals.length < 3) return {};
  
  // Header is on row 2 (index 1)
  const head = vals[1].map(function(x) { return String(x || '').trim(); });
  const iList = head.indexOf('listName');
  const iValId = head.indexOf('valueId');
  const iValLbl = head.indexOf('valueLabel');
  
  if (iList < 0 || iValId < 0 || iValLbl < 0) return {};
  
  const map = {};
  // Data starts at row 3 (index 2)
  for (var r = 2; r < vals.length; r++) {
    if (String(vals[r][iList] || '').trim() === listName) {
      const id = String(vals[r][iValId] || '').trim();
      const lbl = String(vals[r][iValLbl] || '').trim();
      if (lbl && id) {
        map[lbl] = id;
        map[lbl.toLowerCase()] = id;
      }
    }
  }
  
  return map;
}

/**
 * Build list name to values map
 * @returns {Object} Map of listName -> { id -> {id, label} }
 */
function buildListNameMap_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.LISTS_SHEET);
  if (!sh) return {};
  
  const vals = sh.getDataRange().getValues();
  if (vals.length < 3) return {};
  
  // Header is on row 2 (index 1), row 1 (index 0) is the title
  const head = vals[1].map(function(x) { return String(x || '').trim(); });
  const iList = head.indexOf('listName');
  const iValId = head.indexOf('valueId');
  const iValLbl = head.indexOf('valueLabel');
  
  if (iList < 0 || iValId < 0 || iValLbl < 0) return {};
  
  const map = {};
  // Data starts at row 3 (index 2)
  for (var r = 2; r < vals.length; r++) {
    const listName = String(vals[r][iList] || '').trim();
    const valId = String(vals[r][iValId] || '').trim();
    const valLbl = String(vals[r][iValLbl] || '').trim();
    
    if (listName && valId) {
      if (!map[listName]) map[listName] = {};
      map[listName][valId] = { id: valId, label: valLbl };
    }
  }
  
  return map;
}

/**
 * Get list of available sites from Bob Lists
 * @returns {Array<string>} Array of site names
 */
function getSitesList_() {
  const listMap = buildListNameMap_();
  if (listMap['Site']) {
    return Object.values(listMap['Site']).map(function(v) { return v.label; }).filter(Boolean);
  }
  return [];
}

/**
 * Get list of available locations from Bob Lists
 * @returns {Array<string>} Array of location names
 */
function getLocationsList_() {
  const listMap = buildListNameMap_();
  if (listMap['Location']) {
    return Object.values(listMap['Location']).map(function(v) { return v.label; }).filter(Boolean);
  }
  return [];
}

/**
 * Write upload result to Uploader sheet
 * @param {Sheet} sheet - The uploader sheet
 * @param {number} row - Row number
 * @param {string} status - Status message
 * @param {number|string} code - HTTP code
 * @param {string} error - Error message
 * @param {string} verified - Verified value
 */
function writeUploaderResult_(sheet, row, status, code, error, verified) {
  sheet.getRange(row, 5).setValue(status);
  sheet.getRange(row, 6).setValue(code);
  sheet.getRange(row, 7).setValue(error);
  sheet.getRange(row, 8).setValue(verified);
  
  const statusCell = sheet.getRange(row, 5);
  if (status === 'COMPLETED') {
    statusCell.setBackground(CONFIG.COLORS.SUCCESS);
  } else if (status === 'SKIP') {
    statusCell.setBackground(CONFIG.COLORS.WARNING);
  } else if (status === 'FAILED') {
    statusCell.setBackground(CONFIG.COLORS.ERROR);
  } else if (status.indexOf('Processing') >= 0) {
    statusCell.setBackground(CONFIG.COLORS.INFO);
  }
}

/**
 * Build PUT request body for updating a field
 * @param {string} jsonPath - Field's JSON path
 * @param {*} value - Value to set
 * @returns {Object} Request body object
 */
function buildPutBody_(jsonPath, value) {
  const parts = jsonPath.replace(/^root\./, '').split('.');
  const body = {};
  
  var current = body;
  for (var i = 0; i < parts.length - 1; i++) {
    current[parts[i]] = {};
    current = current[parts[i]];
  }
  
  current[parts[parts.length - 1]] = value;
  return body;
}

/**
 * Read back a field value after PUT to verify
 * @param {string} auth - Base64 auth header
 * @param {string} bobId - Employee Bob ID
 * @param {string} jsonPath - Field's JSON path
 * @returns {string} The field value or empty string
 */
function readBackField_(auth, bobId, jsonPath) {
  try {
    const url = CONFIG.HIBOB_BASE_URL + '/v1/people/' + encodeURIComponent(bobId);
    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      headers: { 
        Authorization: 'Basic ' + auth, 
        Accept: 'application/json' 
      }
    });
    
    if (resp.getResponseCode() !== 200) return '';
    
    const person = JSON.parse(resp.getContentText());
    return String(getVal_(person, jsonPath) || '');
  } catch (e) {
    return '';
  }
}

/**
 * Get value from nested object using dot notation path
 * @param {Object} obj - Object to search
 * @param {string} path - Dot-notation path
 * @returns {*} The value or empty string
 */
function getVal_(obj, path) {
  if (!obj || !path) return '';
  
  if (path.indexOf('.') === -1 && path.indexOf('/') === -1) {
    const val = obj[path];
    if (val !== undefined && val !== null && val !== '') {
      return val;
    }
  }
  
  const parts = path.split('.');
  var current = obj;
  for (var i = 0; i < parts.length; i++) {
    if (current == null) break;
    current = current[parts[i]];
  }
  if (current !== undefined && current !== null && current !== '') {
    if (typeof current === 'object' && current !== null && 'value' in current) {
      return current.value;
    }
    return current;
  }
  
  const slashPath = '/' + path.replace(/\./g, '/');
  const slashNode = obj[slashPath];
  if (slashNode !== undefined && slashNode !== null) {
    if (typeof slashNode === 'object' && 'value' in slashNode) {
      return slashNode.value;
    }
    if (slashNode !== '') {
      return slashNode;
    }
  }
  
  if (path.indexOf('root.') !== 0) {
    const withRoot = getVal_(obj, 'root.' + path);
    if (withRoot) return withRoot;
  }
  
  return '';
}

/**
 * Get value with human-readable fallback
 * @param {Object} obj - Object to search
 * @param {string} path - Dot-notation path
 * @param {Object} idToLabelMap - Map of ID to label
 * @returns {*} The value or empty string
 */
function getValWithHumanReadable_(obj, path, idToLabelMap) {
  if (!obj || !path) return '';
  
  if (path.indexOf('custom.') >= 0) {
    const pathParts = path.replace(/^root\./, '').split('.');
    
    if (obj.humanReadable) {
      var hrValue = obj.humanReadable;
      for (var i = 0; i < pathParts.length; i++) {
        if (hrValue == null) break;
        hrValue = hrValue[pathParts[i]];
      }
      if (hrValue && hrValue !== '') {
        return hrValue;
      }
    }
    
    const machineValue = getVal_(obj, path);
    if (machineValue && idToLabelMap && idToLabelMap[machineValue]) {
      return idToLabelMap[machineValue];
    }
    if (machineValue) return machineValue;
  }
  
  return getVal_(obj, path);
}

/**
 * Sanitize field path for API requests
 * @param {string} path - Field path
 * @returns {string} Sanitized path
 */
function sanitiseFieldPath_(path) {
  if (String(path || '').indexOf('custom.') >= 0) {
    if (path.indexOf('root.') !== 0) {
      return 'root.' + path;
    }
    return path;
  }
  return String(path || '').replace(/^root\./, '');
}

/**
 * Build map of list ID to label
 * @param {string} listName - Name of the list
 * @returns {Object} Map of ID -> label
 */
function buildListIdToLabelMap_(listName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.LISTS_SHEET);
  if (!sh) return {};
  const vals = sh.getDataRange().getValues();
  if (vals.length < 3) return {};
  // Header is on row 2 (index 1)
  const head = vals[1].map(function(x) { return String(x || '').trim(); });
  const iList = head.indexOf('listName');
  const iValId = head.indexOf('valueId');
  const iValLbl = head.indexOf('valueLabel');
  if (iList < 0 || iValId < 0 || iValLbl < 0) return {};
  const map = {};
  // Data starts at row 3 (index 2)
  for (var r = 2; r < vals.length; r++) {
    if (String(vals[r][iList] || '').trim() === listName) {
      const id = String(vals[r][iValId] || '').trim();
      const lbl = String(vals[r][iValLbl] || '').trim();
      if (id) map[id] = lbl || id;
    }
  }
  return map;
}

// ============================================================================
// MENU STRUCTURE
// ============================================================================

/**
 * Creates unified menu when spreadsheet is opened
 * Combines all functionality: Salary Data, Performance Reports, Field Uploader
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Bob')
    // Documentation
    .addItem('>> View Documentation', 'createDocumentationSheet')
    .addSeparator()
    
    // Test Connection
    .addItem('>> Test API Connection', 'testApiConnection')
    .addSeparator()
    
    // SETUP Section
    .addSubMenu(ui.createMenu('SETUP')
      .addItem('1. Pull Fields', 'pullFields')
      .addItem('2. Pull Lists', 'pullLists')
      .addItem('3. Pull Employees', 'setupAndPullEmployees')
      .addItem('   >> Refresh Employees (Keep Filters)', 'pullEmployees')
      .addSeparator()
      .addItem('4. Setup Field Uploader', 'setupUploader')
      .addItem('5. Select Field to Update', 'showFieldSelector')
      .addSeparator()
      .addSubMenu(ui.createMenu('6. History Tables')
        .addItem('Setup History Uploader', 'setupHistoryUploader')
        .addItem('Generate Columns for Table', 'generateHistoryColumns')
      )
    )
    .addSeparator()
    
    // VALIDATE Section  
    .addSubMenu(ui.createMenu('VALIDATE')
      .addItem('Validate Field Upload Data', 'validateUploadData')
      .addItem('Validate History Upload Data', 'validateHistoryUpload')
    )
    .addSeparator()
    
    // UPLOAD Section
    .addSubMenu(ui.createMenu('UPLOAD')
      .addItem('Quick Upload (<40 rows)', 'runQuickUpload')
      .addItem('Batch Upload (40-1000+ rows)', 'runBatchUpload')
      .addItem('Retry Failed Rows Only', 'retryFailedRows')
    )
    .addSeparator()
    
    // MONITORING Section
    .addSubMenu(ui.createMenu('MONITORING')
      .addItem('Check Batch Status', 'checkBatchStatus')
    )
    .addSeparator()
    
    // CONTROL Section
    .addSubMenu(ui.createMenu('CONTROL')
      .addItem('Stop Batch Upload', 'clearBatchUpload')
    )
    .addSeparator()
    
    // CLEANUP Section
    .addSubMenu(ui.createMenu('CLEANUP')
      .addItem('Clear All Upload Data', 'clearAllUploadData')
    )
    .addSeparator()
    
    // Salary Data (Legacy)
    .addSubMenu(ui.createMenu('Salary Data')
      .addItem('Import Base Data', 'importBobDataSimpleWithLookup')
      .addItem('Import Bonus History', 'importBobBonusHistoryLatest')
      .addItem('Import Compensation History', 'importBobCompHistoryLatest')
      .addItem('Import Full Comp History', 'importBobFullCompHistory')
      .addSeparator()
      .addItem('Import All Data', 'importAllBobData')
      .addSeparator()
      .addItem('Convert Tenure to Array Formula', 'convertTenureToArrayFormula'))
    .addSeparator()
    
    // Performance Reports
    .addSubMenu(ui.createMenu('Performance Reports')
      .addItem('🚀 Launch Web Interface', 'launchWebInterface')
      .addSeparator()
      .addItem('Set HiBob Credentials', 'setHiBobCredentials')
      .addItem('View Credentials Status', 'viewCredentialsStatus')
      .addSeparator()
      .addItem('📖 Instructions & Help', 'showPerformanceReportInstructions'))
    
    .addToUi();
}

// ============================================================================
// FIELD UPLOADER - Core Functions
// ============================================================================

/**
 * Pulls all available fields from Bob API and stores in metadata sheet
 */
function pullFields() {
  const { auth } = getCreds_();
  const sh = getOrCreateSheet_(SHEET_FIELDS);
  
  const url = `${BASE}/v1/company/people/fields`;
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' }
  });

  const code = resp.getResponseCode();
  const text = resp.getContentText();
  
  if (code !== 200) {
    throw new Error(`Failed to fetch fields (${code}): ${text.slice(0,500)}`);
  }
  
  const data = JSON.parse(text);
  const allFields = Array.isArray(data) ? data : (data.fields || []);

  sh.clear();
  
  sh.getRange('A1').setValue('Field Metadata  Pulled from HiBob')
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setFontWeight('bold')
    .setFontSize(14);
  sh.getRange('A1:I1').mergeAcross();

  const headers = ['id', 'name', 'jsonPath', 'category', 'type', 'calculated', 'description', 'listName', 'typeData (raw JSON)'];
  sh.getRange(2, 1, 1, headers.length).setValues([headers]);
  formatHeaderRow_(sh, 2, headers.length);

  const rows = [];
  allFields.forEach(f => {
    const fieldId = safe(f.id || '');
    const fieldName = safe(f.name || '');
    const jsonPath = safe(f.jsonPath || '');
    const category = safe(f.category || '');
    const fieldType = safe(f.type || '');
    const calculated = safe(f.calculated || '');
    const description = safe(f.description || '');
    const listName = safe(f.typeData?.listName || '');
    const typeDataJson = f.typeData ? JSON.stringify(f.typeData) : '';
    
    rows.push([fieldId, fieldName, jsonPath, category, fieldType, calculated, description, listName, typeDataJson]);
  });

  if (rows.length > 0) {
    sh.getRange(3, 1, rows.length, headers.length).setValues(rows);
  }

  autoFitAllColumns_(sh);
  toast_(`[OK] Fields: ${rows.length} rows pulled from HiBob`);
}

/**
 * Pulls all available lists from Bob API
 */
function pullLists() {
  const { auth } = getCreds_();
  const sh = getOrCreateSheet_(SHEET_LISTS);

  const url = `${BASE}/v1/company/named-lists`;
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' }
  });

  const code = resp.getResponseCode();
  const text = resp.getContentText();
  
  if (code !== 200) {
    throw new Error(`Failed to fetch named lists (${code}): ${text.slice(0,500)}`);
  }
  
  // HiBob returns an object where keys are list names
  const listsData = JSON.parse(text);

  sh.clear();
  
  sh.getRange('A1').setValue('List Values - Pulled from HiBob')
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setFontWeight('bold')
    .setFontSize(14);
  sh.getRange('A1:D1').mergeAcross();

  const headers = ['listName', 'valueId', 'valueLabel', 'extraInfo'];
  sh.getRange(2, 1, 1, headers.length).setValues([headers]);
  formatHeaderRow_(sh, 2, headers.length);

  const rows = [];
  
  // Iterate through each list in the object
  Object.keys(listsData).forEach(listKey => {
    const listObj = listsData[listKey];
    const listName = safe(listObj.name || listKey);
    
    // Try both 'items' and 'values' properties
    const items = listObj.items || listObj.values || [];
    
    if (items.length === 0) {
      rows.push([listName, '', '', '(No items)']);
    } else {
      items.forEach(item => {
        const itemId = safe(item.id || '');
        const itemLabel = safe(item.value || item.label || item.name || '');
        const extraInfo = item.color ? `Color: ${item.color}` : '';
        rows.push([listName, itemId, itemLabel, extraInfo]);
      });
    }
  });

  if (rows.length > 0) {
    sh.getRange(3, 1, rows.length, headers.length).setValues(rows);
  }

  autoFitAllColumns_(sh);
  
  const listCount = Object.keys(listsData).length;
  toast_(`[OK] Lists: ${listCount} lists, ${rows.length} items total`);
}

/**
 * Test API Connection
 */
function testApiConnection() {
  try {
    const { id, auth } = getCreds_();
    Logger.log(`[AUTH] Using HiBob service user: ${id}`);
    
    const url = `${BASE}/v1/company/people/fields`;
    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      headers: { 
        Authorization: `Basic ${auth}`, 
        Accept: 'application/json' 
      }
    });
    
    const code = resp.getResponseCode();
    const text = resp.getContentText();
    
    if (code === 0) {
      SpreadsheetApp.getUi().alert(
        '[ERROR] Connection Failed',
        `Could not connect to HiBob API.\n\n` +
        `Possible causes:\n` +
        ` Network connectivity issues\n` +
        ` Invalid BASE URL\n` +
        ` Service blocked\n\n` +
        `Current BASE URL: ${BASE}\n\n` +
        `Check the Apps Script logs for details.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    if (code >= 500) {
      SpreadsheetApp.getUi().alert(
        '[ERROR] Server Error',
        `HiBob API returned server error.\n\n` +
        `HTTP Status: ${code}\n\n` +
        `This is typically a temporary issue.\n` +
        `Please try again in a few moments.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    if (code >= 400 && code < 500) {
      SpreadsheetApp.getUi().alert(
        '[ERROR] Authentication or Access Error',
        `Could not authenticate with HiBob API.\n\n` +
        `Possible causes:\n` +
        ` Invalid credentials\n` +
        ` Authentication failed (check credentials)\n` +
        ` Service user lacks permissions\n\n` +
        `Current BASE URL: ${BASE}\n` +
        `HTTP Status: ${code}\n\n` +
        `Check the Apps Script logs for details.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    if (code === 200) {
      try {
        const data = JSON.parse(text);
        SpreadsheetApp.getUi().alert(
          '[OK] Connection Successful!',
          `Successfully connected to HiBob API.\n\n` +
          `Service User: ${id}\n` +
          `Base URL: ${BASE}\n` +
          `Fields Retrieved: ${Array.isArray(data) ? data.length : 'unknown'}\n\n` +
          `You can now use "Pull Fields", "Pull Lists", etc.`,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      } catch (e) {
        SpreadsheetApp.getUi().alert(
          '[WARN] Parse Error',
          `API responded with 200 but JSON parse failed.\n\n${e.message}\n\nCheck logs for details.`,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
    } else if (code === 401) {
      SpreadsheetApp.getUi().alert(
        '[ERROR] Authentication Failed',
        `HTTP 401 Unauthorized\n\n` +
        `Your service user credentials are incorrect.\n\n` +
        `Run: setBobServiceUser("id", "token") with correct credentials.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else if (code === 403) {
      SpreadsheetApp.getUi().alert(
        '[ERROR] Access Forbidden',
        `HTTP 403 Forbidden\n\n` +
        `Service user lacks required permissions.\n\n` +
        `Check HiBob admin settings for service user permissions.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        '[ERROR] Unexpected Response',
        `HTTP ${code}\n\n${text.slice(0, 500)}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      '[ERROR] Connection Failed',
      `Error: ${e.message}\n\nCheck Apps Script logs for details.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    Logger.log('Connection test error: ' + e);
  }
}

// Continue with remaining functions... (This file will be very large, so I'll continue in next part)

