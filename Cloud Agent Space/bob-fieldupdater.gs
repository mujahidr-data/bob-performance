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

// ============================================================================
// EMPLOYEE PULL FUNCTIONS
// ============================================================================

/**
 * Sheet 3  Pull Employees with ENHANCED SETUP
 */
function setupAndPullEmployees() {
  const { auth } = getCreds_();
  const sh = getOrCreateSheet_(SHEET_EMPLOYEES);
  
  // Clear everything including old validations AND filters first
  sh.clear();
  sh.clearNotes();
  
  // Remove any existing filters first
  try {
    const filter = sh.getFilter();
    if (filter) {
      filter.remove();
      Logger.log('Removed existing filter');
    }
  } catch (e) {
    Logger.log('Note: No filter to remove - ' + e.message);
  }
  
  const lastRow = sh.getMaxRows();
  const lastCol = sh.getMaxColumns();
  if (lastRow > 0 && lastCol > 0) {
    sh.getRange(1, 1, lastRow, lastCol).clearDataValidations();
  }

  // Build list maps for dropdowns
  const listMap = buildListNameMap_();
  const siteOptions = listMap['site'] ? Object.values(listMap['site']).map(v => v.label).filter(Boolean) : [];
  const locationOptions = listMap['root.field_1696948109629'] ? Object.values(listMap['root.field_1696948109629']).map(v => v.label).filter(Boolean) : [];
  
  // Status is determined by work.isActive field, not a list - use hardcoded values
  const statusOptions = ['Active', 'Inactive'];
  
  // Row 1: Title
  sh.getRange('A1').setValue('Employee Data - Filters & Configuration')
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setFontWeight('bold')
    .setFontSize(14);
  sh.getRange('A1:H1').mergeAcross();

  // Row 2: Instructions
  sh.getRange('A2').setValue('Configure filters below, then data will load at row 8')
    .setFontStyle('italic')
    .setFontColor('#666666');
  sh.getRange('A2:H2').mergeAcross();

  // Row 3: Employment Status filter
  sh.getRange('A3').setValue('Employment Status *')
    .setFontWeight('bold');
  const statusCell = sh.getRange('B3');
  statusCell.setValue('Active');
  formatRequiredInput_(statusCell, 'Active = current employees | Inactive = terminated employees');
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusOptions, true)
    .setAllowInvalid(false)
    .build();
  statusCell.setDataValidation(statusRule);

  // Row 4: Site filter
  sh.getRange('A4').setValue('Site (optional)');
  const siteCell = sh.getRange('B4');
  formatOptionalInput_(siteCell, 'Leave empty to include all sites');
  if (siteOptions.length > 0) {
    const siteRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['', ...siteOptions], true)
      .setAllowInvalid(true)
      .build();
    siteCell.setDataValidation(siteRule);
  }

  // Row 5: Location filter
  sh.getRange('A5').setValue('Location (optional)');
  const locationCell = sh.getRange('B5');
  formatOptionalInput_(locationCell, 'Leave empty to include all locations');
  if (locationOptions.length > 0) {
    const locationRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['', ...locationOptions], true)
      .setAllowInvalid(true)
      .build();
    locationCell.setDataValidation(locationRule);
  }

  // Row 6: Tip
  sh.getRange('A6').setValue('[TIP] Select Active or Inactive status. Leave Site/Location blank to include all.')
    .setFontStyle('italic')
    .setFontColor('#E67C73')
    .setBackground('#FFF3CD');
  sh.getRange('A6:H6').mergeAcross();

  // Row 7: Data headers
  const headers = ['Bob ID', 'CIQ ID', 'Employee Name', 'Site', 'Location', 'Employment Status', 'Employment Type', 'Date of Hire'];
  sh.getRange(7, 1, 1, headers.length).setValues([headers]);
  formatHeaderRow_(sh, 7, headers.length);
  
  // Freeze rows before pulling data
  sh.setFrozenRows(7);
  
  // Now pull the employee data (this will add rows starting at row 8)
  pullEmployeesData_(sh, auth);
}

function pullEmployees() {
  const { auth } = getCreds_();
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_EMPLOYEES);
  
  if (!sh) {
    throw new Error('Employees sheet not found. Run "3. Pull Employees" from SETUP menu first.');
  }
  
  // Check if sheet is set up (has headers at row 7)
  const headerCheck = sh.getRange('A7').getValue();
  if (!headerCheck || headerCheck !== 'Bob ID') {
    throw new Error('Employees sheet not set up properly. Run "3. Pull Employees" from SETUP menu to set up first.');
  }
  
  // Remove existing filter before clearing data
  try {
    const filter = sh.getFilter();
    if (filter) {
      filter.remove();
      Logger.log('Removed existing filter before refresh');
    }
  } catch (e) {
    Logger.log('Note: No filter to remove - ' + e.message);
  }
  
  // Clear only the data rows (row 8 onwards), keep filters, headers, and dropdowns
  const lastRow = sh.getLastRow();
  if (lastRow > 7) {
    sh.getRange(8, 1, lastRow - 7, sh.getMaxColumns()).clearContent();
  }
  
  // Now pull the employee data with current filter settings
  pullEmployeesData_(sh, auth);
}

function pullEmployeesData_(sh, auth) {
  // Read filter values
  const filterStatus = normalizeBlank_(sh.getRange('B3').getValue());
  const filterSite = normalizeBlank_(sh.getRange('B4').getValue());
  const filterLocation = normalizeBlank_(sh.getRange('B5').getValue());

  if (!filterStatus) {
    throw new Error('Employment Status is required (row 3).');
  }

  // Build list maps - only needed for dropdown filters
  const listMap = buildListNameMap_();
  const siteOptions = listMap['site'] ? Object.values(listMap['site']).map(v => v.label).filter(Boolean) : [];
  const locationOptions = listMap['root.field_1696948109629'] ? Object.values(listMap['root.field_1696948109629']).map(v => v.label).filter(Boolean) : [];

  // Determine if we should show inactive employees based on filter
  const showInactive = /^inactive$/i.test(filterStatus);

  // Build search body
  const searchBody = {
    showInactive: showInactive,
    humanReadable: 'REPLACE'
  };

  // Fetch employees using POST /v1/people/search
  const url = `${BASE}/v1/people/search`;
  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    muteHttpExceptions: true,
    headers: { 
      Authorization: `Basic ${auth}`, 
      Accept: 'application/json',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(searchBody)
  });

  const code = resp.getResponseCode();
  const text = resp.getContentText();
  
  if (code !== 200) {
    throw new Error(`Failed to fetch employees (${code}): ${text.slice(0,500)}`);
  }
  
  // Check if response is HTML (authentication error)
  if (text.trim().toLowerCase().startsWith('<!doctype') || text.trim().toLowerCase().startsWith('<html')) {
    throw new Error(`API returned HTML instead of JSON. This usually means authentication failed or wrong endpoint.\n\nHTTP ${code}\n\nResponse preview: ${text.slice(0,300)}`);
  }

  let data;
  try {
    data = JSON.parse(text);
  } catch (parseError) {
    throw new Error(`Failed to parse API response as JSON.\n\nHTTP ${code}\n\nResponse preview: ${text.slice(0,500)}\n\nParse error: ${parseError.message}`);
  }
  const employees = data.employees || [];
  
  Logger.log(`Pull Employees Debug:`);
  Logger.log(`- Filter Status: ${filterStatus}`);
  Logger.log(`- showInactive param: ${showInactive}`);
  Logger.log(`- API returned ${employees.length} employees`);

  // Apply site/location filters after API call
  let filteredEmployees = employees;
  
  if (filterSite) {
    filteredEmployees = filteredEmployees.filter(emp => {
      const siteVal = getVal_(emp, 'work.site') 
                   || getVal_(emp, 'work.siteId')
                   || '';
      return String(siteVal).toLowerCase() === filterSite.toLowerCase();
    });
  }
  
  if (filterLocation) {
    filteredEmployees = filteredEmployees.filter(emp => {
      let locVal = '';
      if (emp.custom && emp.custom.field_1696948109629) {
        locVal = emp.custom.field_1696948109629;
      } else {
        locVal = getVal_(emp, 'custom.field_1696948109629') 
              || getVal_(emp, 'root.field_1696948109629')
              || '';
      }
      return String(locVal).toLowerCase() === filterLocation.toLowerCase();
    });
  }

  const rows = [];
  filteredEmployees.forEach(emp => {
    const bobId = safe(emp.id || getVal_(emp, 'root.id'));
    const ciqId = safe(getVal_(emp, 'work.employeeIdInCompany') || '');
    
    // Get display name
    let displayName = '';
    if (emp.fullName) {
      displayName = emp.fullName;
    } else if (emp.displayName) {
      displayName = emp.displayName;
    } else {
      const first = getVal_(emp, 'root.firstName') || '';
      const last = getVal_(emp, 'root.surname') || '';
      displayName = (first + ' ' + last).trim();
    }
    
    // Get site value
    const siteVal = getVal_(emp, 'work.site') 
                 || getVal_(emp, 'work.siteId')
                 || '';
    
    // Get location value
    let locationVal = '';
    if (emp.custom && emp.custom.field_1696948109629) {
      locationVal = emp.custom.field_1696948109629;
    } else {
      locationVal = getVal_(emp, 'custom.field_1696948109629') 
                 || getVal_(emp, 'root.field_1696948109629')
                 || '';
    }
    
    // Get employment status
    let empStatus = 'Unknown';
    const isActive = getVal_(emp, 'work.isActive');
    const statusFromWork = getVal_(emp, 'work.status');
    const statusFromLifecycle = getVal_(emp, 'lifecycle.status');
    
    if (isActive === true) {
      empStatus = 'Active';
    } else if (isActive === false) {
      empStatus = 'Inactive';
    } else if (statusFromWork) {
      empStatus = statusFromWork;
    } else if (statusFromLifecycle) {
      empStatus = statusFromLifecycle;
    } else {
      empStatus = showInactive ? 'Inactive' : 'Active';
    }
    
    // Get employment type
    let employmentTypeVal = '';
    employmentTypeVal = getVal_(emp, 'payroll.employment.type')
                     || getVal_(emp, 'work.payrollEmploymentType')
                     || getVal_(emp, 'work.employmentType')
                     || '';
    
    const hireDate = safe(getVal_(emp, 'work.startDate') || '');

    rows.push([bobId, ciqId, displayName, siteVal, locationVal, empStatus, employmentTypeVal, hireDate]);
  });

  if (rows.length === 0) {
    let message = 'No employees found matching filters';
    if (!showInactive && employees.length === 0) {
      message = 'No active employees found. Try selecting "Inactive" status to see terminated employees.';
    } else if (employees.length > 0 && rows.length === 0) {
      message = `Found ${employees.length} employees but none matched Site/Location filters. Check filter values.`;
    }
    sh.getRange(8, 1, 1, 8).setValues([[message, '', '', '', '', '', '', '']]);
    toast_(message);
  } else {
    sh.getRange(8, 1, rows.length, 8).setValues(rows);
    toast_(`[OK] Employees: ${rows.length} rows pulled from HiBob`);
  }

  // Auto-fit columns first
  autoFitAllColumns_(sh);
  
  // Create filter starting from row 7 (header row)
  try {
    if (!sh.getFilter()) {
      const lastRow = sh.getLastRow();
      const lastCol = sh.getLastColumn();
      if (lastRow >= 7 && lastCol > 0) {
        sh.getRange(7, 1, lastRow - 6, lastCol).createFilter();
        Logger.log('Created filter on employee data');
      }
    } else {
      Logger.log('Filter already exists, skipping creation');
    }
  } catch (e) {
    Logger.log('Note: Could not create filter - ' + e.message);
  }
}

// ============================================================================
// FIELD UPLOADER - Setup and Selection
// ============================================================================

/**
 * Sheet 4  Field Uploader Setup
 */
function setupUploader() {
  const sh = getOrCreateSheet_(SHEET_UPLOADER);
  
  // Clear everything including old validations
  sh.clear();
  sh.clearNotes();
  const lastRow = sh.getMaxRows();
  const lastCol = sh.getMaxColumns();
  if (lastRow > 0 && lastCol > 0) {
    sh.getRange(1, 1, lastRow, lastCol).clearDataValidations();
  }

  sh.getRange('A1').setValue('Field Uploader - Update Single Field for Multiple Employees')
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setFontWeight('bold')
    .setFontSize(14);
  sh.getRange('A1:H1').mergeAcross();

  sh.getRange('A2').setValue('INSTRUCTIONS:')
    .setFontWeight('bold')
    .setFontSize(11);
  sh.getRange('A2:H2').mergeAcross();

  const instructions = [
    '1. Click "SETUP -> 5. Select Field to Update" to choose which field to update',
    '2. Paste employee CIQ IDs in column A below (starting row 8)',
    '3. Enter new values in column B',
    '4. Validate with "VALIDATE -> Validate Field Upload Data"',
    '5. Upload with "UPLOAD -> Quick Upload" or "Batch Upload"'
  ];

  instructions.forEach((instr, i) => {
    sh.getRange(3 + i, 1).setValue(instr)
      .setFontStyle('italic')
      .setWrap(true);
    sh.getRange(3 + i, 1, 1, 8).mergeAcross();
  });

  const dataHeaders = ['CIQ ID', 'New Value', 'Bob ID', 'Field Path', 'Status', 'Code', 'Error', 'Verified Value'];
  sh.getRange(7, 1, 1, dataHeaders.length).setValues([dataHeaders]);
  formatHeaderRow_(sh, 7, dataHeaders.length);

  sh.setFrozenRows(7);
  autoFitAllColumns_(sh);
  
  toast_('[OK] Field Uploader sheet ready. Use menu to select a field.');
}

function showFieldSelector() {
  const template = HtmlService.createTemplateFromFile('FieldSelector');
  template.fields = getFieldsForSelector();
  
  const html = template.evaluate()
    .setWidth(800)
    .setHeight(600)
    .setTitle('Select Field to Update');
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Field to Update');
}

function getFieldsForSelector() {
  const fields = readFields_();
  return fields
    .filter(f => f.calculated !== 'true' && f.type !== 'object')
    .map(f => ({
      id: f.id,
      name: f.name,
      jsonPath: f.jsonPath,
      type: f.type,
      listName: f.typeData?.listName || ''
    }));
}

function setSelectedField(fieldName) {
  try {
    Logger.log(`[setSelectedField] Called with fieldName: ${fieldName}`);
    
    const fields = readFields_();
    Logger.log(`[setSelectedField] Total fields available: ${fields.length}`);
    
    const field = fields.find(f => f.name === fieldName);
    
    if (!field) {
      Logger.log(`[setSelectedField] Field NOT FOUND: ${fieldName}`);
      throw new Error(`Field not found: ${fieldName}`);
    }
    
    Logger.log(`[setSelectedField] Field found: ${JSON.stringify(field)}`);

    const ul = getOrCreateSheet_(SHEET_UPLOADER);
    
    ul.getRange('D1').setValue(field.name);
    ul.getRange('E1').setValue(field.jsonPath);
    ul.getRange('F1').setValue(field.id);
    ul.getRange('G1').setValue(field.type);
    ul.getRange('H1').setValue(field.typeData?.listName || '');
    
    ul.getRange('D1:H1')
      .setBackground('#E8F0FE')
      .setFontWeight('bold')
      .setBorder(true, true, true, true, false, false, '#4285F4', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    
    SpreadsheetApp.getActive().setActiveSheet(ul);
    SpreadsheetApp.getActive().setActiveRange(ul.getRange('A8'));
    
    Logger.log(`[setSelectedField] Field successfully set in uploader sheet`);
    toast_(`✓ Selected: ${field.name}`);
    
    return `Selected field: ${field.name} (${field.jsonPath})`;
  } catch (error) {
    Logger.log(`[setSelectedField] ERROR: ${error.message}`);
    throw error;
  }
}

// ============================================================================
// VALIDATION
// ============================================================================

function validateUploadData() {
  const ul = getOrCreateSheet_(SHEET_UPLOADER);
  
  const fieldName = normalizeBlank_(ul.getRange('D1').getValue());
  const fieldPath = normalizeBlank_(ul.getRange('E1').getValue());
  
  if (!fieldName || !fieldPath) {
    throw new Error(' No field selected!\n\nUse " SETUP  5. Select Field to Update" first.');
  }
  
  const lastRow = ul.getLastRow();
  if (lastRow < 8) {
    throw new Error(' No data to validate!\n\nAdd employee CIQ IDs and new values starting at row 8.');
  }
  
  const data = ul.getRange(8, 1, lastRow - 7, 2).getValues();
  
  let emptyRows = 0;
  let validRows = 0;
  let issues = [];
  
  data.forEach((row, i) => {
    const ciq = normalizeBlank_(row[0]);
    const newVal = normalizeBlank_(row[1]);
    const rowNum = i + 8;
    
    if (!ciq && !newVal) {
      emptyRows++;
      return;
    }
    
    if (!ciq) {
      issues.push(`Row ${rowNum}: Missing CIQ ID`);
      return;
    }
    
    if (newVal === '') {
      issues.push(`Row ${rowNum}: Missing new value`);
      return;
    }
    
    validRows++;
  });
  
  let msg = ` Validation Results:\n\n`;
  msg += ` Valid rows: ${validRows}\n`;
  msg += ` Empty rows: ${emptyRows}\n`;
  
  if (issues.length > 0) {
    msg += `\n Issues found (${issues.length}):\n`;
    msg += issues.slice(0, 10).join('\n');
    if (issues.length > 10) {
      msg += `\n... and ${issues.length - 10} more issues`;
    }
    
    SpreadsheetApp.getUi().alert(' Validation Issues', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    msg += `\n All data looks good!\n\nReady to upload ${validRows} updates.`;
    SpreadsheetApp.getUi().alert(' Validation Passed', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ============================================================================
// UPLOAD FUNCTIONS
// ============================================================================

function runQuickUpload() {
  const ul = getOrCreateSheet_(SHEET_UPLOADER);
  const { auth } = getCreds_();
  
  const fieldName = normalizeBlank_(ul.getRange('D1').getValue());
  const fieldPath = normalizeBlank_(ul.getRange('E1').getValue());
  const fieldType = normalizeBlank_(ul.getRange('G1').getValue());
  const listName = normalizeBlank_(ul.getRange('H1').getValue());
  
  if (!fieldName || !fieldPath) {
    throw new Error(' No field selected. Use "Select Field to Update" first.');
  }
  
  const lastRow = ul.getLastRow();
  if (lastRow < 8) {
    throw new Error(' No data to upload. Add CIQ IDs and values starting at row 8.');
  }
  
  const dataRows = lastRow - 7;
  if (dataRows > 40) {
    const response = SpreadsheetApp.getUi().alert(
      ' Large Dataset Detected',
      `You have ${dataRows} rows.\n\n` +
      `Quick Upload is best for <40 rows.\n` +
      `For ${dataRows} rows, use "Batch Upload" instead.\n\n` +
      `Continue with Quick Upload anyway?`,
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (response !== SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
  
  const data = ul.getRange(8, 1, dataRows, 2).getValues();
  const ciqToBobMap = buildCiqToBobMap_();
  
  let listMap = null;
  if (listName) {
    listMap = buildListLabelToId_(listName);
  }
  
  let ok = 0, skip = 0, fail = 0;
  let foundTerminated = 0;
  
  for (let i = 0; i < data.length; i++) {
    const rowNum = i + 8;
    const ciq = normalizeBlank_(data[i][0]);
    const rawNew = normalizeBlank_(data[i][1]);
    
    if (!ciq) {
      continue;
    }
    
    ul.getRange(rowNum, 5).setValue(' Processing...');
    SpreadsheetApp.flush();
    
    let bobId = ciqToBobMap[ciq];
    
    if (!bobId) {
      bobId = lookupBobIdFromApi_(auth, ciq);
    }
    
    if (!bobId) {
      writeUploaderResult_(ul, rowNum, 'FAILED', '', `CIQ ${ciq} not found`, '');
      fail++;
      continue;
    }
    
    const wasInOriginal = ciqToBobMap[ciq];
    if (!wasInOriginal && bobId) {
      foundTerminated++;
      ul.getRange(rowNum, 3).setNote(' Found in terminated employees');
    }
    
    ul.getRange(rowNum, 3).setValue(bobId);
    ul.getRange(rowNum, 4).setValue(fieldPath);
    
    let newVal = rawNew;
    if (listMap) {
      const hit = listMap[rawNew] || listMap[String(rawNew).toLowerCase()];
      if (hit) newVal = hit;
    }
    
    const body = buildPutBody_(fieldPath, newVal);
    const putUrl = `${BASE}/v1/people/${encodeURIComponent(bobId)}`;
    let code = 0, text = '';
    
    try {
      const resp = UrlFetchApp.fetch(putUrl, {
        method: 'put',
        muteHttpExceptions: true,
        headers: { 
          Authorization: `Basic ${auth}`, 
          Accept: 'application/json', 
          'Content-Type': 'application/json' 
        },
        payload: JSON.stringify(body)
      });
      code = resp.getResponseCode();
      text = resp.getContentText();
    } catch(e) { 
      code = 0; 
      text = String(e && e.message || e); 
    }
    
    if (code >= 200 && code < 300) {
      const verified = readBackField_(auth, bobId, fieldPath);
      writeUploaderResult_(ul, rowNum, 'COMPLETED', code, '', verified);
      ok++;
    } else if (code === 304) {
      const verified = readBackField_(auth, bobId, fieldPath);
      writeUploaderResult_(ul, rowNum, 'SKIP', code, 'Already correct', verified);
      skip++;
    } else if (code === 404) {
      writeUploaderResult_(ul, rowNum, 'FAILED', code, 'Bob API 404', '');
      fail++;
    } else {
      writeUploaderResult_(ul, rowNum, 'FAILED', code || '', (text||'').slice(0,200), '');
      fail++;
    }
    
    SpreadsheetApp.flush();
    Utilities.sleep(PUT_DELAY_MS);
  }
  
  let msg = ` Upload Complete!\n\n`;
  msg += ` Completed: ${ok}\n`;
  msg += ` Skipped: ${skip}\n`;
  msg += ` Failed: ${fail}`;
  
  if (foundTerminated > 0) {
    msg += `\n\n Found ${foundTerminated} terminated employees!`;
  }
  
  SpreadsheetApp.getUi().alert('Upload Results', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function runBatchUpload() {
  const ul = getOrCreateSheet_(SHEET_UPLOADER);
  
  const fieldName = normalizeBlank_(ul.getRange('D1').getValue());
  const fieldPath = normalizeBlank_(ul.getRange('E1').getValue());
  
  if (!fieldName || !fieldPath) {
    throw new Error(' No field selected. Use "Select Field to Update" first.');
  }
  
  const lastRow = ul.getLastRow();
  if (lastRow < 8) {
    throw new Error(' No data to upload.');
  }
  
  const dataRows = lastRow - 7;
  const estimatedMinutes = Math.ceil(dataRows / BATCH_SIZE) * TRIGGER_INTERVAL;
  
  const response = SpreadsheetApp.getUi().alert(
    ' Start Batch Upload?',
    `This will upload ${dataRows} rows in batches of ${BATCH_SIZE}.\n\n` +
    `Estimated time: ~${estimatedMinutes} minutes\n\n` +
    `The upload will run automatically in the background.\n` +
    `You can close this spreadsheet and check progress later.\n\n` +
    `Continue?`,
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  
  if (response !== SpreadsheetApp.getUi().Button.YES) {
    return;
  }
  
  const props = PropertiesService.getScriptProperties();
  props.setProperty('BATCH_UPLOAD_STATE', JSON.stringify({
    nextRow: 8,
    fieldName: fieldName,
    fieldPath: fieldPath,
    startTime: new Date().toISOString(),
    lastBatchTime: new Date().toISOString(),
    stats: { completed: 0, skipped: 0, failed: 0 }
  }));
  
  createBatchTrigger();
  
  toast_(` Batch upload started! Check progress with " MONITORING  Check Batch Status"`);
}

function processBatch() {
  const props = PropertiesService.getScriptProperties();
  const stateJson = props.getProperty('BATCH_UPLOAD_STATE');
  
  if (!stateJson) {
    return;
  }
  
  const state = JSON.parse(stateJson);
  const ul = getOrCreateSheet_(SHEET_UPLOADER);
  const { auth } = getCreds_();
  
  const fieldPath = state.fieldPath;
  const listName = normalizeBlank_(ul.getRange('H1').getValue());
  
  const lastRow = ul.getLastRow();
  const endRow = Math.min(state.nextRow + BATCH_SIZE - 1, lastRow);
  
  if (state.nextRow > lastRow) {
    clearBatchUpload();
    clearUploadDataAfterBatch_(state.stats);
    return;
  }
  
  const data = ul.getRange(state.nextRow, 1, endRow - state.nextRow + 1, 2).getValues();
  const ciqToBobMap = buildCiqToBobMap_();
  
  let listMap = null;
  if (listName) {
    listMap = buildListLabelToId_(listName);
  }
  
  const stats = state.stats || { completed: 0, skipped: 0, failed: 0 };
  
  for (let i = 0; i < data.length; i++) {
    const rowNum = state.nextRow + i;
    const ciq = normalizeBlank_(data[i][0]);
    const rawNew = normalizeBlank_(data[i][1]);
    
    if (!ciq) continue;
    
    ul.getRange(rowNum, 5).setValue(' Processing...');
    
    let bobId = ciqToBobMap[ciq];
    if (!bobId) bobId = lookupBobIdFromApi_(auth, ciq);
    
    if (!bobId) {
      writeUploaderResult_(ul, rowNum, 'FAILED', '', `CIQ ${ciq} not found`, '');
      stats.failed++;
      continue;
    }
    
    ul.getRange(rowNum, 3).setValue(bobId);
    ul.getRange(rowNum, 4).setValue(fieldPath);
    
    let newVal = rawNew;
    if (listMap) {
      const hit = listMap[rawNew] || listMap[String(rawNew).toLowerCase()];
      if (hit) newVal = hit;
    }
    
    const body = buildPutBody_(fieldPath, newVal);
    const putUrl = `${BASE}/v1/people/${encodeURIComponent(bobId)}`;
    
    try {
      const resp = UrlFetchApp.fetch(putUrl, {
        method: 'put',
        muteHttpExceptions: true,
        headers: { 
          Authorization: `Basic ${auth}`, 
          Accept: 'application/json', 
          'Content-Type': 'application/json' 
        },
        payload: JSON.stringify(body)
      });
      
      const code = resp.getResponseCode();
      
      if (code >= 200 && code < 300) {
        const verified = readBackField_(auth, bobId, fieldPath);
        writeUploaderResult_(ul, rowNum, 'COMPLETED', code, '', verified);
        stats.completed++;
      } else if (code === 304) {
        const verified = readBackField_(auth, bobId, fieldPath);
        writeUploaderResult_(ul, rowNum, 'SKIP', code, 'Already correct', verified);
        stats.skipped++;
      } else {
        writeUploaderResult_(ul, rowNum, 'FAILED', code, resp.getContentText().slice(0,200), '');
        stats.failed++;
      }
    } catch(e) {
      writeUploaderResult_(ul, rowNum, 'FAILED', 0, String(e).slice(0,200), '');
      stats.failed++;
    }
    
    Utilities.sleep(PUT_DELAY_MS);
  }
  
  SpreadsheetApp.flush();
  
  state.nextRow = endRow + 1;
  state.lastBatchTime = new Date().toISOString();
  state.stats = stats;
  
  props.setProperty('BATCH_UPLOAD_STATE', JSON.stringify(state));
}

function lookupBobIdFromApi_(auth, ciq) {
  try {
    const searchUrl = `${BASE}/v1/people/search`;
    const searchBody = { fields: ['internal.customId'], values: [ciq] };
    
    const resp = UrlFetchApp.fetch(searchUrl, {
      method: 'post',
      muteHttpExceptions: true,
      headers: { 
        Authorization: `Basic ${auth}`, 
        Accept: 'application/json',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(searchBody)
    });
    
    if (resp.getResponseCode() === 200) {
      const result = JSON.parse(resp.getContentText());
      if (result.employees && result.employees.length > 0) {
        return result.employees[0].id;
      }
    }
  } catch(e) {
    Logger.log('API search failed for CIQ ' + ciq + ': ' + e);
  }
  return null;
}

function retryFailedRows() {
  const ul = getOrCreateSheet_(SHEET_UPLOADER);
  const { auth } = getCreds_();
  
  const fieldPath = normalizeBlank_(ul.getRange('E1').getValue());
  const listName = normalizeBlank_(ul.getRange('H1').getValue());
  
  if (!fieldPath) {
    throw new Error(' No field selected.');
  }
  
  const lastRow = ul.getLastRow();
  if (lastRow < 8) {
    throw new Error(' No data to retry.');
  }
  
  const data = ul.getRange(8, 1, lastRow - 7, 5).getValues();
  const ciqToBobMap = buildCiqToBobMap_();
  
  let listMap = null;
  if (listName) {
    listMap = buildListLabelToId_(listName);
  }
  
  let ok = 0, skip = 0, fail = 0;
  
  for (let i = 0; i < data.length; i++) {
    const rowNum = i + 8;
    const status = normalizeBlank_(data[i][4]);
    
    if (status !== 'FAILED') continue;
    
    const ciq = normalizeBlank_(data[i][0]);
    const rawNew = normalizeBlank_(data[i][1]);
    
    if (!ciq) continue;
    
    ul.getRange(rowNum, 5).setValue(' Retrying...');
    SpreadsheetApp.flush();
    
    let bobId = ciqToBobMap[ciq];
    if (!bobId) {
      bobId = lookupBobIdFromApi_(auth, ciq);
    }
    
    if (!bobId) { 
      writeUploaderResult_(ul, rowNum, 'FAILED', '', `CIQ ${ciq} still not found`, ''); 
      fail++;
      continue; 
    }
    
    ul.getRange(rowNum, 3).setValue(bobId);
    ul.getRange(rowNum, 4).setValue(fieldPath);
    
    let newVal = rawNew;
    if (listMap) {
      const hit = listMap[rawNew] || listMap[String(rawNew).toLowerCase()];
      if (hit) newVal = hit;
    }
    
    const body = buildPutBody_(fieldPath, newVal);
    const putUrl = `${BASE}/v1/people/${encodeURIComponent(bobId)}`;
    let code = 0, text = '';
    
    try {
      const resp = UrlFetchApp.fetch(putUrl, {
        method: 'put',
        muteHttpExceptions: true,
        headers: { 
          Authorization: `Basic ${auth}`, 
          Accept: 'application/json', 
          'Content-Type': 'application/json' 
        },
        payload: JSON.stringify(body)
      });
      code = resp.getResponseCode();
      text = resp.getContentText();
    } catch(e) { 
      code = 0; 
      text = String(e && e.message || e); 
    }
    
    if (code >= 200 && code < 300) {
      const verified = readBackField_(auth, bobId, fieldPath);
      writeUploaderResult_(ul, rowNum, 'COMPLETED', code, '', verified);
      ok++;
    } else if (code === 304) {
      const verified = readBackField_(auth, bobId, fieldPath);
      writeUploaderResult_(ul, rowNum, 'SKIP', code, 'Already correct', verified);
      skip++;
    } else if (code === 404) {
      writeUploaderResult_(ul, rowNum, 'FAILED', code, 'Bob API 404', '');
      fail++;
    } else {
      writeUploaderResult_(ul, rowNum, 'FAILED', code || '', (text||'').slice(0,200), '');
      fail++;
    }
    
    SpreadsheetApp.flush();
    Utilities.sleep(PUT_DELAY_MS);
  }
  
  let msg = ` Retry Complete!\n\n Completed: ${ok}\n Skipped: ${skip}\n Failed: ${fail}`;
  
  SpreadsheetApp.getUi().alert('Retry Results', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function checkBatchStatus() {
  const props = PropertiesService.getScriptProperties();
  const stateJson = props.getProperty('BATCH_UPLOAD_STATE');
  
  if (!stateJson) {
    SpreadsheetApp.getUi().alert(
      ' No Batch Running',
      'No batch upload is currently running.\n\nStart one with " UPLOAD  Batch Upload"',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const state = JSON.parse(stateJson);
  const ul = getOrCreateSheet_(SHEET_UPLOADER);
  const totalRows = ul.getLastRow() - 7;
  const completed = state.nextRow - 8;
  const progress = Math.round((completed / totalRows) * 100);
  const remaining = totalRows - completed;
  const estimatedMinutes = Math.ceil(remaining / BATCH_SIZE) * TRIGGER_INTERVAL;
  
  const stats = state.stats || { completed: 0, skipped: 0, failed: 0 };
  
  const startTime = new Date(state.startTime || state.lastBatchTime);
  const elapsed = Math.round((new Date() - startTime) / 60000);
  
  SpreadsheetApp.getUi().alert(
    ' Batch Upload Status',
    `Field: ${state.fieldName}\n\n` +
    `Progress: ${completed}/${totalRows} rows (${progress}%)\n` +
    `Remaining: ${remaining} rows\n\n` +
    `Results So Far:\n` +
    ` Completed: ${stats.completed}\n` +
    ` Skipped: ${stats.skipped}\n` +
    ` Failed: ${stats.failed}\n\n` +
    `Time Elapsed: ${elapsed} minutes\n` +
    `Est. Remaining: ~${estimatedMinutes} minutes\n\n` +
    `Last Batch: ${new Date(state.lastBatchTime).toLocaleString()}\n` +
    `Next Batch: In ~${TRIGGER_INTERVAL} minutes`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function createBatchTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processBatch') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  ScriptApp.newTrigger('processBatch')
    .timeBased()
    .everyMinutes(TRIGGER_INTERVAL)
    .create();
}

function clearBatchUpload() {
  const props = PropertiesService.getScriptProperties();
  const hadBatch = !!props.getProperty('BATCH_UPLOAD_STATE');
  
  props.deleteProperty('BATCH_UPLOAD_STATE');
  
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processBatch') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  if (hadBatch) {
    toast_(' Batch upload stopped and cleared.');
  } else {
    toast_(' No batch upload was running.');
  }
}

function clearUploadDataAfterBatch_(stats) {
  const ul = getOrCreateSheet_(SHEET_UPLOADER);
  const lastRow = ul.getLastRow();
  
  if (lastRow <= 7) return;
  
  const dataRows = lastRow - 7;
  
  const response = SpreadsheetApp.getUi().alert(
    ' Batch Upload Complete!',
    `Final Results:\n\n` +
    ` Completed: ${stats.completed}\n` +
    ` Skipped: ${stats.skipped}\n` +
    ` Failed: ${stats.failed}\n\n` +
    `Clear ${dataRows} rows of upload data?\n` +
    `(Field selection will be kept for next upload)`,
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  
  if (response === SpreadsheetApp.getUi().Button.YES) {
    ul.getRange(8, 1, dataRows, ul.getMaxColumns()).clearContent().clearFormat();
    
    if (dataRows > 100) {
      ul.deleteRows(8, dataRows);
    }
    
    toast_(' Upload data cleared. Ready for next upload!');
  }
}

function clearAllUploadData() {
  const ss = SpreadsheetApp.getActive();
  const regularSheet = ss.getSheetByName(SHEET_UPLOADER);
  const historySheet = ss.getSheetByName(CONFIG.HISTORY_SHEET);
  
  let clearedSheets = [];
  let totalRowsCleared = 0;
  
  // Clear regular uploader
  if (regularSheet && regularSheet.getLastRow() > 7) {
    const dataRows = regularSheet.getLastRow() - 7;
    const maxCols = regularSheet.getMaxColumns();
    
    regularSheet.getRange(8, 1, dataRows, maxCols)
      .clearContent()
      .clearFormat()
      .clearNote()
      .clearDataValidations();
    
    if (dataRows > 100) {
      regularSheet.deleteRows(8, dataRows);
    }
    
    clearedSheets.push('Field Uploader');
    totalRowsCleared += dataRows;
  }
  
  // Clear history uploader
  if (historySheet && historySheet.getLastRow() > 13) {
    const dataRows = historySheet.getLastRow() - 13;
    const maxCols = historySheet.getMaxColumns();
    
    historySheet.getRange(14, 1, dataRows, maxCols)
      .clearContent()
      .clearFormat()
      .clearNote()
      .clearDataValidations();
    
    if (dataRows > 100) {
      historySheet.deleteRows(14, dataRows);
    }
    
    clearedSheets.push('History Uploader');
    totalRowsCleared += dataRows;
  }
  
  if (clearedSheets.length === 0) {
    SpreadsheetApp.getUi().alert(
      ' Already Clean',
      'No upload data to clear.\n\nBoth uploader sheets are empty and ready to use.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert(
      ' Data Cleared',
      `Cleared ${totalRowsCleared} rows from:\n ${clearedSheets.join('\n ')}\n\n` +
      `Removed all notes, formatting, and validations.\n` +
      `Field/table selections preserved.\n` +
      `Ready for new uploads!`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ============================================================================
// HISTORY UPLOADER FUNCTIONS
// ============================================================================

function setupHistoryUploader() {
  const sh = getOrCreateSheet_('History Uploader');
  sh.clear();
  
  // Title
  sh.getRange('A1').setValue('History Table Uploader')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT);
  
  sh.getRange('A2').setValue('Update salary, work history, or variable pay tables')
    .setFontStyle('italic')
    .setFontColor('#666666');
  
  sh.getRange('A3').setValue('Table Type (select one) *').setFontWeight('bold');
  
  formatRequiredInput_(sh.getRange('B4'), 'Select table type: Salary/Payroll, Work History, or Variable Pay');
  sh.getRange('B4').setValue('')
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Salary / Payroll', 'Work History', 'Variable Pay'], true)
        .setAllowInvalid(false)
        .setHelpText('⚠️ REQUIRED: Select which history table to update')
        .build()
    );
  
  sh.getRange('A6').setValue('Instructions:').setFontWeight('bold')
    .setBackground(CONFIG.COLORS.SECTION_HEADER);
  sh.getRange('A7').setValue('1. Select table type in B4 above');
  sh.getRange('A8').setValue('2. Click "Generate Columns for Table"');
  sh.getRange('A9').setValue('3. Paste your data starting at row 14');
  sh.getRange('A10').setValue('4. Click "Validate History Upload Data"');
  sh.getRange('A11').setValue('5. Click "Quick" or "Batch" upload');
  
  sh.setColumnWidth(1, 300);
  sh.setColumnWidth(2, 250);
  
  toast_('✅ History Uploader sheet created.\n\nSelect table type in B4, then generate columns.');
}

function generateHistoryColumns() {
  const sh = getOrCreateSheet_('History Uploader');
  const tableType = String(sh.getRange('B4').getValue() || '').trim();
  
  if (!tableType) {
    SpreadsheetApp.getUi().alert(
      '⚠️ No Table Selected',
      'Select a table type in B4 first.\n\nChoices: Salary/Payroll, Work History, or Variable Pay',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  let columns = [];
  
  if (tableType === 'Salary / Payroll') {
    columns = [
      { name: 'CIQ ID', required: true, listName: null },
      { name: 'Effective Date *', required: true, listName: null },
      { name: 'Base Salary *', required: true, listName: null },
      { name: 'Currency *', required: true, listName: 'Currency' },
      { name: 'Pay Period *', required: true, listName: null },
      { name: 'Pay Frequency', required: false, listName: null },
      { name: 'Change Type', required: false, listName: 'Change Type' },
      { name: 'Reason', required: false, listName: null }
    ];
  } else if (tableType === 'Work History') {
    columns = [
      { name: 'CIQ ID', required: true, listName: null },
      { name: 'Effective Date *', required: true, listName: null },
      { name: 'Job Title', required: false, listName: 'Job Title' },
      { name: 'Department', required: false, listName: 'Department' },
      { name: 'Site', required: false, listName: 'Site' },
      { name: 'Reports To', required: false, listName: null },
      { name: 'Change Type', required: false, listName: 'Change Type' },
      { name: 'Reason', required: false, listName: null }
    ];
  } else if (tableType === 'Variable Pay') {
    columns = [
      { name: 'CIQ ID', required: true, listName: null },
      { name: 'Effective Date *', required: true, listName: null },
      { name: 'Variable Type', required: false, listName: 'Variable Type' },
      { name: 'Amount', required: false, listName: null },
      { name: 'Pay Period', required: false, listName: null },
      { name: 'Pay Frequency', required: false, listName: null },
      { name: 'Reason', required: false, listName: null }
    ];
  }
  
  // Store column count in metadata for later use
  const metaRange = sh.getRange('D2');
  metaRange.setValue(columns.length);
  
  // Clear rows 12-1000 
  sh.getRange(12, 1, 989, 50).clearContent().clearFormat();
  
  // Write label in row 12
  formatSectionHeader_(sh.getRange(12, 1), 'Data Columns - Paste data starting at row 14');
  
  // Build full headers: data columns + status columns
  const headerRow = columns.map(c => c.name);
  const statusColumns = ['GET Status', 'POST Status', 'HTTP', 'Error', 'Entry ID'];
  const fullHeader = headerRow.concat(statusColumns);
  
  // Write headers in row 13 (data entry starts at row 14)
  sh.getRange(13, 1, 1, fullHeader.length).setValues([fullHeader]);
  formatHeaderRow_(sh, 13, fullHeader.length);
  
  // Get all list values from Bob Lists
  const listMap = buildListNameMap_();
  
  // Apply data validation for each column
  for (let colIndex = 0; colIndex < columns.length; colIndex++) {
    const colNum = colIndex + 1;
    const column = columns[colIndex];
    const range = sh.getRange(14, colNum, 1000, 1);
    
    // Check if this column has a list name
    if (column.listName && listMap[column.listName]) {
      const listValues = Object.values(listMap[column.listName]).map(v => v.label);
      
      if (listValues.length > 0) {
        range.setDataValidation(
          SpreadsheetApp.newDataValidation()
            .requireValueInList(listValues, true)
            .setHelpText(`🔽 Select from ${column.listName} list`)
            .build()
        );
      }
    } else if (column.name === 'Pay Period *' || column.name === 'Pay Period') {
      const payPeriodVals = ['Annual', 'Hourly', 'Daily', 'Weekly', 'Monthly'];
      range.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(payPeriodVals, true)
          .setHelpText('🔽 Select pay period')
          .build()
      );
    } else if (column.name === 'Pay Frequency') {
      const payFreqVals = ['Monthly', 'Semi Monthly', 'Weekly', 'Bi-Weekly'];
      range.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(payFreqVals, true)
          .setHelpText('🔽 Select pay frequency')
          .build()
      );
    }
  }
  
  // Freeze header rows
  sh.setFrozenRows(13);
  
  // Auto-fit all columns
  autoFitAllColumns_(sh);
  
  toast_(`✅ Columns generated for "${tableType}"\n\nAll list fields have dropdowns.\nPaste data starting at row 14.`);
}

function validateHistoryUpload() {
  const sh = getOrCreateSheet_('History Uploader');
  const tableType = String(sh.getRange('B4').getValue() || '').trim();
  
  if (!tableType) {
    SpreadsheetApp.getUi().alert(
      '⚠️ No Table Selected',
      'Select a table type in B4 first.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const lastRow = sh.getLastRow();
  if (lastRow <= 13) {
    SpreadsheetApp.getUi().alert(
      '⚠️ No Data',
      'No data to validate.\n\nPaste data starting at row 14.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const dataRange = sh.getRange(14, 1, lastRow - 13, 10).getValues();
  const mapCIQtoBob = buildCiqToBobMap_();
  
  let valid = 0, invalid = 0, empty = 0;
  const issues = [];
  
  for (let i = 0; i < dataRange.length; i++) {
    const row = 14 + i;
    const ciq = String(dataRange[i][0] || '').trim();
    const effectiveDate = String(dataRange[i][1] || '').trim();
    const value2 = String(dataRange[i][2] || '').trim();
    
    if (!ciq && !effectiveDate && !value2) {
      empty++;
      continue;
    }
    
    if (!ciq) {
      invalid++;
      if (issues.length < 10) issues.push(`Row ${row}: Missing CIQ ID`);
      continue;
    }
    
    if (!effectiveDate) {
      invalid++;
      if (issues.length < 10) issues.push(`Row ${row}: Missing Effective Date`);
      continue;
    }
    
    if (!mapCIQtoBob[ciq]) {
      invalid++;
      if (issues.length < 10) issues.push(`Row ${row}: CIQ ${ciq} not found`);
      continue;
    }
    
    valid++;
  }
  
  const totalRows = dataRange.length - empty;
  const estimatedMinutes = Math.ceil(totalRows / BATCH_SIZE) * TRIGGER_INTERVAL;
  
  let msg = `📊 History Upload Validation\n\n`;
  msg += `Table: ${tableType}\n`;
  msg += `Total Rows: ${totalRows}\n\n`;
  msg += `✅ Valid: ${valid}\n`;
  
  if (empty > 0) msg += `⏭️ Empty: ${empty}\n`;
  if (invalid > 0) msg += `❌ Invalid: ${invalid}\n`;
  
  if (issues.length > 0) {
    msg += `\n⚠️ Issues (first 10):\n`;
    issues.forEach(issue => msg += `• ${issue}\n`);
    
    if (invalid > 10) msg += `...and ${invalid - 10} more issues\n`;
    
    msg += `\n💡 SOLUTION: Change B3 to "All" in Employees sheet,\n`;
    msg += `   then re-run "3. Pull Employees"\n`;
  }
  
  msg += `\n`;
  
  if (valid > 0) {
    msg += `✅ READY TO UPLOAD!\n\n`;
    msg += `📊 Estimated Time: ~${estimatedMinutes} minutes\n`;
    msg += totalRows <= 40 ? `💡 Recommendation: Use Quick Upload` : `💡 Recommendation: Use Batch Upload`;
  } else {
    msg += `⚠️ FIX ERRORS BEFORE UPLOADING`;
  }
  
  SpreadsheetApp.getUi().alert('Validation Report', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================================
// DOCUMENTATION FUNCTION
// ============================================================================

/**
 * Create comprehensive documentation sheet
 * Version: 2.0
 * Date: 2025-10-29
 */
function createDocumentationSheet() {
  const sh = getOrCreateSheet_(CONFIG.DOCS_SHEET);
  sh.clear();
  
  // Title
  sh.getRange('A1').setValue('HiBob Data Updater - Complete Guide')
    .setFontSize(18)
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT);
  
  sh.getRange('A2').setValue('Version 2.0 | Last Updated: 2025-10-29')
    .setFontStyle('italic')
    .setFontColor('#666666');
  
  // Table of Contents
  let row = 4;
  formatSectionHeader_(sh.getRange(`A${row}`), '📋 TABLE OF CONTENTS');
  row++;
  
  const toc = [
    ['1. Overview', 'What this tool does and how it works'],
    ['2. Initial Setup', 'One-time configuration steps'],
    ['3. Workflow Steps', 'Step-by-step process for updates'],
    ['4. Function Reference', 'Detailed explanation of each menu item'],
    ['5. Troubleshooting', 'Common issues and solutions'],
    ['6. Best Practices', 'Tips for efficient usage'],
    ['7. API Rate Limits', 'Understanding HiBob API constraints']
  ];
  
  toc.forEach(item => {
    sh.getRange(`A${row}`).setValue(item[0]).setFontWeight('bold');
    sh.getRange(`B${row}`).setValue(item[1]);
    row++;
  });
  
  row += 2;
  
  // 1. OVERVIEW
  formatSectionHeader_(sh.getRange(`A${row}`), '1. OVERVIEW');
  row++;
  
  sh.getRange(`A${row}`).setValue('The HiBob Data Updater is a Google Apps Script tool that:');
  row++;
  const overview = [
    '• Pulls employee data, field metadata, and list values from HiBob API',
    '• Allows bulk updates to employee fields via an intuitive spreadsheet interface',
    '• Supports both regular fields and history tables (Salary, Work History, Variable Pay)',
    '• Handles rate limiting and batch processing automatically',
    '• Validates data before upload to prevent errors',
    '• Provides detailed status reporting and error messages'
  ];
  overview.forEach(item => {
    sh.getRange(`A${row}`).setValue(item);
    row++;
  });
  
  row += 2;
  
  // 2. INITIAL SETUP
  formatSectionHeader_(sh.getRange(`A${row}`), '2. INITIAL SETUP (One-Time)');
  row++;
  
  sh.getRange(`A${row}`).setValue('Before using the tool, complete these steps once:');
  row++;
  
  const setup = [
    ['Step', 'Action', 'How To'],
    ['1', 'Get HiBob API Credentials', 'Log in to HiBob → Settings → API → Service Users → Create/Copy credentials'],
    ['2', 'Store Credentials', 'Extensions → Apps Script → Select "setBobServiceUser" → Run → Enter ID and Token'],
    ['3', 'Test Connection', 'Extensions → Apps Script → Select "testBobConnection" → Run → Check logs'],
    ['4', 'Pull Field Metadata', 'Bob menu → Setup → 1. Pull Fields'],
    ['5', 'Pull List Values', 'Bob menu → Setup → 2. Pull Lists'],
    ['6', 'Pull Employee Data', 'Bob menu → Setup → 3. Pull Employees']
  ];
  
  sh.getRange(row, 1, setup.length, setup[0].length).setValues(setup);
  formatHeaderRow_(sh, row, 3);
  row += setup.length + 1;
  
  // 3. WORKFLOW
  formatSectionHeader_(sh.getRange(`A${row}`), '3. STANDARD WORKFLOW');
  row++;
  
  const workflow = [
    ['Phase', 'Steps', 'Description'],
    ['SETUP', '1. Select Update Type', 'Choose between Regular Field Update or History Table Update'],
    ['', '2. Setup Uploader Sheet', 'Creates template with appropriate columns and validations'],
    ['', '3. Select Field/Table', 'Choose which field or history table to update'],
    ['VALIDATE', '4. Prepare Data', 'Paste CIQ IDs and new values into the sheet'],
    ['', '5. Run Validation', 'Checks for missing CIQs, blank values, and data quality'],
    ['UPLOAD', '6. Choose Upload Method', 'Quick (<40 rows) or Batch (40-1000+ rows)'],
    ['', '7. Monitor Progress', 'Check status column for real-time updates'],
    ['CLEANUP', '8. Review Results', 'Check completed/skipped/failed counts'],
    ['', '9. Clear Data', 'Clean sheet for next upload (keeps field selection)']
  ];
  
  sh.getRange(row, 1, workflow.length, workflow[0].length).setValues(workflow);
  formatHeaderRow_(sh, row, 3);
  row += workflow.length + 2;
  
  // Auto-fit all columns
  autoFitAllColumns_(sh);
  
  // Set column widths for better readability
  sh.setColumnWidth(1, 250);
  sh.setColumnWidth(2, 400);
  sh.setColumnWidth(3, 350);
  
  // Freeze header
  sh.setFrozenRows(1);
  
  toast_('✅ Documentation sheet created!');
}

// Note: Additional functions for salary data and performance reports can be added from bob-performance-module.gs
// This file focuses on the Field Updater functionality as requested.
