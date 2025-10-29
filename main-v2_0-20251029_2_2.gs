/*************************************************
 * HiBob Data Updater - MAIN FILE
 * Version: 2.0
 * Date: 2025-10-29 (Update 2)
 * 
 * Changes in this update:
 * - Fixed filters to apply from row 7 (data starts at row 8)
 * - Added "Date of Hire" and "Date of Termination" columns
 * - Fixed employment status to use status field
 * - Fixed employment type to use payrollEmploymentType valueLabel
 * - Added Site dropdown with proper valueLabels
 * - Added Location dropdown from root.field_1696948109629
 * - Enhanced cleanup to clear notes, formatting, and validation
 *************************************************/

/* ===== Aliases from CONFIG ===== */
const BASE            = CONFIG.HIBOB_BASE_URL;
const SHEET_FIELDS    = CONFIG.META_SHEET;
const SHEET_LISTS     = CONFIG.LISTS_SHEET;
const SHEET_EMPLOYEES = CONFIG.EMPLOYEES_SHEET;
const SHEET_UPLOADER  = CONFIG.UPDATES_SHEET;

/* ===== Rate limiting for PUT /v1/people/{id} ===== */
const PUTS_PER_MIN = CONFIG.PUTS_PER_MINUTE || 10;
const PUT_DELAY_MS = Math.ceil(60000 / PUTS_PER_MIN);

/* ===== Batch upload constants ===== */
const BATCH_SIZE = 45;
const TRIGGER_INTERVAL = 5;
const MAX_EXECUTION_TIME = 330000;

/* ===== NEW MENU STRUCTURE ===== */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Bob')
    // Documentation
    .addItem('📖 View Documentation', 'createDocumentationSheet')
    .addSeparator()
    
    // Test Connection
    .addItem('🔧 Test API Connection', 'testApiConnection')
    .addSeparator()
    
    // SETUP Section
    .addSubMenu(ui.createMenu('📥 SETUP')
      .addItem('1. Pull Fields', 'pullFields')
      .addItem('2. Pull Lists', 'pullLists')
      .addItem('3. Pull Employees', 'setupAndPullEmployees')
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
    .addSubMenu(ui.createMenu('✅ VALIDATE')
      .addItem('Validate Field Upload Data', 'validateUploadData')
      .addItem('Validate History Upload Data', 'validateHistoryUpload')
    )
    .addSeparator()
    
    // UPLOAD Section
    .addSubMenu(ui.createMenu('🚀 UPLOAD')
      .addItem('Quick Upload (<40 rows)', 'runQuickUpload')
      .addItem('Batch Upload (40-1000+ rows)', 'runBatchUpload')
      .addItem('Retry Failed Rows Only', 'retryFailedRows')
    )
    .addSeparator()
    
    // MONITORING Section
    .addSubMenu(ui.createMenu('📊 MONITORING')
      .addItem('Check Batch Status', 'checkBatchStatus')
    )
    .addSeparator()
    
    // CONTROL Section
    .addSubMenu(ui.createMenu('🛡️ CONTROL')
      .addItem('Stop Batch Upload', 'clearBatchUpload')
    )
    .addSeparator()
    
    // CLEANUP Section
    .addSubMenu(ui.createMenu('🧹 CLEANUP')
      .addItem('Clear All Upload Data', 'clearAllUploadData')
    )
    
    .addToUi();
}

/* =====================================================
 * Sheet 1 — Fields metadata
 * ===================================================== */
function pullFields() {
  const { auth } = getCreds_();
  const resp = UrlFetchApp.fetch(`${BASE}/v1/company/people/fields`, {
    method: 'get',
    muteHttpExceptions: true,
    headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' }
  });
  
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  
  if (code < 200 || code >= 300) {
    throw new Error(`Fields fetch failed (${code}): ${text.slice(0,300)}`);
  }
  
  // Check if response is HTML
  if (text.trim().toLowerCase().startsWith('<!doctype') || text.trim().toLowerCase().startsWith('<html')) {
    throw new Error(`API returned HTML instead of JSON. Check credentials and API endpoint.\n\nHTTP ${code}\n\nResponse: ${text.slice(0,300)}`);
  }
  
  let data;
  try {
    data = JSON.parse(text);
  } catch (parseError) {
    throw new Error(`Failed to parse response as JSON.\n\nHTTP ${code}\n\nResponse: ${text.slice(0,500)}\n\nError: ${parseError.message}`);
  }
  
  if (!Array.isArray(data)) throw new Error('Unexpected response for fields (expected array).');

  const sh = getOrCreateSheet_(SHEET_FIELDS);
  sh.clear();

  const header = [
    'id','name','category','description','jsonPath','type',
    'historical','calculated',
    'typeData.listName','typeData.multiple','typeData.subType','typeData.example',
    'typeData (raw JSON)','raw (full JSON)'
  ];
  
  const rows = data.map(f => {
    const td = f.typeData || {};
    return [
      safe(f.id), safe(f.name), safe(f.category), safe(f.description), safe(f.jsonPath), safe(f.type),
      boolOrBlank(f.historical), boolOrBlank(f.calculated),
      safe(td.listName), boolOrBlank(td.multiple), safe(td.subType), safe(td.example),
      jsonOrBlank(td), jsonOrBlank(f)
    ];
  });

  sh.getRange(1,1,1,header.length).setValues([header]);
  formatHeaderRow_(sh, 1, header.length);
  
  if (rows.length) sh.getRange(2,1,rows.length,header.length).setValues(rows);
  
  sh.setFrozenRows(1);
  safeCreateFilter_(sh);
  autoFitAllColumns_(sh);
  
  toast_(`✅ Fields: ${rows.length} rows pulled from HiBob`);
}

/* =====================================================
 * Sheet 2 — Company named lists
 * ===================================================== */
function pullLists() {
  const { auth } = getCreds_();
  const sh = getOrCreateSheet_(SHEET_LISTS);
  sh.clear();

  const header = ['listId','listName','valueId','valueLabel','extraJson'];
  sh.getRange(1,1,1,header.length).setValues([header]);
  formatHeaderRow_(sh, 1, header.length);

  const baseHeaders = { Authorization: `Basic ${auth}`, Accept: 'application/json' };
  const rows = [];

  let bulkWorked = false, diagA = null;
  try {
    const url = `${BASE}/v1/company/named-lists?includeArchived=false`;
    const r = fetchJsonLoose_(url, { method:'get', muteHttpExceptions:true, headers: baseHeaders });
    if (r.json) {
      const lists = normalizeNamedLists_(r.json);
      lists.forEach(l => {
        const items = l.items || [];
        if (!items.length) rows.push(['', l.name || '', '', '', JSON.stringify(l || {})]);
        else items.forEach(it => rows.push(['', l.name || '', it.id ?? '', (it.value || it.name || ''), JSON.stringify(it || {})]));
      });
      if (lists.length) bulkWorked = true;
    } else {
      diagA = r;
    }
  } catch (e) {
    diagA = { error: String(e) };
  }

  if (!bulkWorked) {
    const fieldsSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_FIELDS);
    if (!fieldsSheet) throw new Error(`No "${SHEET_FIELDS}" found. Run "1. Pull Fields" first.`);
    const data = fieldsSheet.getDataRange().getValues();
    if (data.length < 2) throw new Error(`"${SHEET_FIELDS}" has no rows. Pull fields first.`);

    const head = data[0].map(x => String(x||'').trim());
    const iList = head.indexOf('typeData.listName');
    if (iList < 0) throw new Error(`Column "typeData.listName" not found in "${SHEET_FIELDS}".`);

    const names = [...new Set(data.slice(1).map(r => String(r[iList]||'').trim()).filter(Boolean))];

    names.forEach(name => {
      const url = `${BASE}/v1/company/named-lists/${encodeURIComponent(name)}?includeArchived=false`;
      try {
        const r = fetchJsonLoose_(url, { method:'get', muteHttpExceptions:true, headers: baseHeaders });
        if (r.json) {
          const list = normalizeNamedLists_(r.json)[0] || { name, items: [] };
          const items = list.items || [];
          if (!items.length) rows.push(['', list.name || name, '', '', JSON.stringify(list || {})]);
          else items.forEach(it => rows.push(['', list.name || name, it.id ?? '', (it.value || it.name || ''), JSON.stringify(it || {})]));
        } else {
          rows.push(['', name, '', '', JSON.stringify({ note:'non-JSON', details: r }).slice(0,500)]);
        }
      } catch (perErr) {
        rows.push(['', name, '', '', JSON.stringify({ error: String(perErr).slice(0,300) })]);
      }
    });
  }

  if (rows.length === 0) {
    const diag = diagA ? JSON.stringify({ note:'No lists returned', bulkNamedLists: diagA }).slice(0,500) : '{"note":"No lists returned"}';
    rows.push(['','','','', diag]);
  }

  sh.getRange(2,1,rows.length,header.length).setValues(rows);
  sh.setFrozenRows(1);
  safeCreateFilter_(sh);
  autoFitAllColumns_(sh);
  
  toast_(`✅ Lists: ${rows.length} rows pulled from HiBob`);
}

function normalizeNamedLists_(payload) {
  if (Array.isArray(payload)) {
    return payload.map(x => ({ name: x?.name || '', items: x?.items || x?.values || [] }));
  }
  if (payload && typeof payload === 'object') {
    return Object.keys(payload).map(k => {
      const node = payload[k] || {};
      return { name: node.name || k, items: node.items || node.values || [] };
    });
  }
  return [];
}

function fetchJsonLoose_(url, options) {
  const resp   = UrlFetchApp.fetch(url, options || {method:'get', muteHttpExceptions:true});
  const code   = resp.getResponseCode();
  const text   = resp.getContentText();
  const header = resp.getHeaders();
  const ctype  = String(header['Content-Type'] || header['content-type'] || '');
  if (/^application\/json/i.test(ctype)) {
    try { return { json: JSON.parse(text), code, text }; }
    catch { return { json: null, code, text }; }
  }
  return { json: null, code, text };
}

/* =====================================================
 * DIAGNOSTIC FUNCTION - Test API Connection
 * ===================================================== */
function testApiConnection() {
  try {
    const { auth, id } = getCreds_();
    Logger.log('Testing HiBob API connection...');
    Logger.log('Service User: ' + id);
    Logger.log('Base URL: ' + BASE);
    
    const url = `${BASE}/v1/company/people/fields`;
    Logger.log('Test URL: ' + url);
    
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
    const contentType = resp.getHeaders()['Content-Type'] || 'unknown';
    
    Logger.log('Response Code: ' + code);
    Logger.log('Content-Type: ' + contentType);
    Logger.log('Response Preview: ' + text.slice(0, 300));
    
    if (text.trim().toLowerCase().startsWith('<!doctype') || text.trim().toLowerCase().startsWith('<html')) {
      SpreadsheetApp.getUi().alert(
        '❌ API Error',
        `API returned HTML instead of JSON!\n\n` +
        `This usually means:\n` +
        `• Wrong API endpoint (check BASE URL in config)\n` +
        `• Authentication failed (check credentials)\n` +
        `• Service user lacks permissions\n\n` +
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
          '✅ Connection Successful!',
          `Successfully connected to HiBob API.\n\n` +
          `Service User: ${id}\n` +
          `Base URL: ${BASE}\n` +
          `Fields Retrieved: ${Array.isArray(data) ? data.length : 'unknown'}\n\n` +
          `You can now use "Pull Fields", "Pull Lists", etc.`,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      } catch (e) {
        SpreadsheetApp.getUi().alert(
          '⚠️ Parse Error',
          `API responded with 200 but JSON parse failed.\n\n${e.message}\n\nCheck logs for details.`,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
    } else if (code === 401) {
      SpreadsheetApp.getUi().alert(
        '❌ Authentication Failed',
        `HTTP 401 Unauthorized\n\n` +
        `Your service user credentials are incorrect.\n\n` +
        `Run: setBobServiceUser("id", "token") with correct credentials.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else if (code === 403) {
      SpreadsheetApp.getUi().alert(
        '❌ Access Forbidden',
        `HTTP 403 Forbidden\n\n` +
        `Service user lacks required permissions.\n\n` +
        `Check HiBob admin settings for service user permissions.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        '❌ Unexpected Response',
        `HTTP ${code}\n\n${text.slice(0, 500)}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      '❌ Connection Failed',
      `Error: ${e.message}\n\nCheck Apps Script logs for details.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    Logger.log('Connection test error: ' + e);
  }
}

/* =====================================================
 * Sheet 3 — Pull Employees with ENHANCED SETUP
 * ===================================================== */
function setupAndPullEmployees() {
  const { auth } = getCreds_();
  const sh = getOrCreateSheet_(SHEET_EMPLOYEES);
  
  // Clear everything including old validations AND filters first
  sh.clear();
  sh.clearNotes();
  
  // Remove any existing filters first
  try {
    if (sh.getFilter()) {
      sh.getFilter().remove();
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
  const statusOptions = listMap['status'] ? Object.values(listMap['status']).map(v => v.label).filter(Boolean) : ['Active', 'Inactive'];
  
  // Add "All" option to status
  const statusOptionsWithAll = ['All', ...statusOptions];
  
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
  formatRequiredInput_(statusCell, 'Select employment status to filter employees');
  if (statusOptionsWithAll.length > 0) {
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(statusOptionsWithAll, true)
      .setAllowInvalid(false)
      .build();
    statusCell.setDataValidation(statusRule);
  }

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
  sh.getRange('A6').setValue('💡 TIP: Use "All" for status, leave Site/Location blank to include everyone')
    .setFontStyle('italic')
    .setFontColor('#E67C73')
    .setBackground('#FFF3CD');
  sh.getRange('A6:H6').mergeAcross();

  // Row 7: Data headers
  const headers = ['Bob ID', 'CIQ ID', 'Employee Name', 'Site', 'Location', 'Employment Status', 'Employment Type', 'Date of Hire', 'Date of Termination'];
  sh.getRange(7, 1, 1, headers.length).setValues([headers]);
  formatHeaderRow_(sh, 7, headers.length);
  
  // Freeze rows before pulling data
  sh.setFrozenRows(7);
  
  // Now pull the employee data (this will add rows starting at row 8)
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

  // Build list maps
  const listMap = buildListNameMap_();
  const siteIdToLabel = listMap['site'] || {};
  const locationIdToLabel = listMap['root.field_1696948109629'] || {};
  const statusIdToLabel = listMap['status'] || {};
  const employmentTypeIdToLabel = listMap['payrollEmploymentType'] || {};

  // Fetch employees
  const url = `${BASE}/v1/people?humanReadable=${CONFIG.SEARCH_HUMANREADABLE}&includeHumanReadable=true`;
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' }
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

  const rows = [];
  employees.forEach(emp => {
    const bobId = safe(emp.id);
    const ciqId = safe(getVal_(emp, 'internal.customId') || getVal_(emp, 'internal.customID') || '');
    const displayName = safe(emp.displayName || getVal_(emp, 'firstName') + ' ' + getVal_(emp, 'surname'));
    
    // Get values with proper label mapping
    const siteId = getVal_(emp, 'site');
    const siteLabel = siteId && siteIdToLabel[siteId] ? siteIdToLabel[siteId].label : siteId;
    
    const locationId = getVal_(emp, 'root.field_1696948109629');
    const locationLabel = locationId && locationIdToLabel[locationId] ? locationIdToLabel[locationId].label : locationId;
    
    const statusId = getVal_(emp, 'work.status');
    const statusLabel = statusId && statusIdToLabel[statusId] ? statusIdToLabel[statusId].label : statusId;
    
    const employmentTypeId = getVal_(emp, 'work.payrollEmploymentType');
    const employmentTypeLabel = employmentTypeId && employmentTypeIdToLabel[employmentTypeId] ? employmentTypeIdToLabel[employmentTypeId].label : employmentTypeId;
    
    const hireDate = safe(getVal_(emp, 'work.startDate') || '');
    const terminationDate = safe(getVal_(emp, 'work.terminationDate') || '');

    // Apply filters
    if (filterStatus.toLowerCase() !== 'all') {
      if (String(statusLabel || '').toLowerCase() !== filterStatus.toLowerCase()) {
        return;
      }
    }

    // Empty filter = include all, otherwise must match
    if (filterSite && String(siteLabel || '').toLowerCase() !== filterSite.toLowerCase()) {
      return;
    }

    // Empty filter = include all, otherwise must match
    if (filterLocation && String(locationLabel || '').toLowerCase() !== filterLocation.toLowerCase()) {
      return;
    }

    rows.push([bobId, ciqId, displayName, siteLabel, locationLabel, statusLabel, employmentTypeLabel, hireDate, terminationDate]);
  });

  if (rows.length === 0) {
    sh.getRange(8, 1, 1, 9).setValues([['No employees found matching filters', '', '', '', '', '', '', '', '']]);
  } else {
    sh.getRange(8, 1, rows.length, 9).setValues(rows);
  }

  // Auto-fit columns first
  autoFitAllColumns_(sh);
  
  // Create filter starting from row 7 (header row) - do this LAST
  try {
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow >= 7 && lastCol > 0) {
      sh.getRange(7, 1, lastRow - 6, lastCol).createFilter();
    }
  } catch (e) {
    Logger.log('Note: Could not create filter - ' + e.message);
  }
  
  toast_(`✅ Employees: ${rows.length} rows pulled from HiBob`);
}

/* =====================================================
 * Sheet 4 — Field Uploader Setup
 * ===================================================== */
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

  sh.getRange('A2').setValue('📋 Instructions:')
    .setFontWeight('bold')
    .setFontSize(11);
  sh.getRange('A2:H2').mergeAcross();

  const instructions = [
    '1. Click "📥 SETUP → 5. Select Field to Update" to choose which field to update',
    '2. Paste employee CIQ IDs in column A below (starting row 8)',
    '3. Enter new values in column B',
    '4. Validate with "✅ VALIDATE → Validate Field Upload Data"',
    '5. Upload with "🚀 UPLOAD → Quick Upload" or "Batch Upload"'
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
  
  toast_('✅ Field Uploader sheet ready. Use menu to select a field.');
}

function showFieldSelector() {
  const html = HtmlService.createHtmlOutputFromFile('FieldSelector')
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

function selectFieldForUpload(fieldId) {
  const fields = readFields_();
  const field = fields.find(f => f.id === fieldId);
  
  if (!field) {
    throw new Error(`Field not found: ${fieldId}`);
  }

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
  
  return `Selected field: ${field.name} (${field.jsonPath})`;
}

/* =====================================================
 * Validation
 * ===================================================== */
function validateUploadData() {
  const ul = getOrCreateSheet_(SHEET_UPLOADER);
  
  const fieldName = normalizeBlank_(ul.getRange('D1').getValue());
  const fieldPath = normalizeBlank_(ul.getRange('E1').getValue());
  
  if (!fieldName || !fieldPath) {
    throw new Error('⚠️ No field selected!\n\nUse "📥 SETUP → 5. Select Field to Update" first.');
  }
  
  const lastRow = ul.getLastRow();
  if (lastRow < 8) {
    throw new Error('⚠️ No data to validate!\n\nAdd employee CIQ IDs and new values starting at row 8.');
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
  
  let msg = `📊 Validation Results:\n\n`;
  msg += `✅ Valid rows: ${validRows}\n`;
  msg += `⚠️ Empty rows: ${emptyRows}\n`;
  
  if (issues.length > 0) {
    msg += `\n❌ Issues found (${issues.length}):\n`;
    msg += issues.slice(0, 10).join('\n');
    if (issues.length > 10) {
      msg += `\n... and ${issues.length - 10} more issues`;
    }
    
    SpreadsheetApp.getUi().alert('⚠️ Validation Issues', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    msg += `\n✅ All data looks good!\n\nReady to upload ${validRows} updates.`;
    SpreadsheetApp.getUi().alert('✅ Validation Passed', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/* =====================================================
 * Upload Functions
 * ===================================================== */
function runQuickUpload() {
  const ul = getOrCreateSheet_(SHEET_UPLOADER);
  const { auth } = getCreds_();
  
  const fieldName = normalizeBlank_(ul.getRange('D1').getValue());
  const fieldPath = normalizeBlank_(ul.getRange('E1').getValue());
  const fieldType = normalizeBlank_(ul.getRange('G1').getValue());
  const listName = normalizeBlank_(ul.getRange('H1').getValue());
  
  if (!fieldName || !fieldPath) {
    throw new Error('⚠️ No field selected. Use "Select Field to Update" first.');
  }
  
  const lastRow = ul.getLastRow();
  if (lastRow < 8) {
    throw new Error('⚠️ No data to upload. Add CIQ IDs and values starting at row 8.');
  }
  
  const dataRows = lastRow - 7;
  if (dataRows > 40) {
    const response = SpreadsheetApp.getUi().alert(
      '⚠️ Large Dataset Detected',
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
    
    ul.getRange(rowNum, 5).setValue('⏳ Processing...');
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
      ul.getRange(rowNum, 3).setNote('🔍 Found in terminated employees');
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
  
  let msg = `✅ Upload Complete!\n\n`;
  msg += `✅ Completed: ${ok}\n`;
  msg += `⏭️ Skipped: ${skip}\n`;
  msg += `❌ Failed: ${fail}`;
  
  if (foundTerminated > 0) {
    msg += `\n\n🔍 Found ${foundTerminated} terminated employees!`;
  }
  
  SpreadsheetApp.getUi().alert('Upload Results', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function runBatchUpload() {
  const ul = getOrCreateSheet_(SHEET_UPLOADER);
  
  const fieldName = normalizeBlank_(ul.getRange('D1').getValue());
  const fieldPath = normalizeBlank_(ul.getRange('E1').getValue());
  
  if (!fieldName || !fieldPath) {
    throw new Error('⚠️ No field selected. Use "Select Field to Update" first.');
  }
  
  const lastRow = ul.getLastRow();
  if (lastRow < 8) {
    throw new Error('⚠️ No data to upload.');
  }
  
  const dataRows = lastRow - 7;
  const estimatedMinutes = Math.ceil(dataRows / BATCH_SIZE) * TRIGGER_INTERVAL;
  
  const response = SpreadsheetApp.getUi().alert(
    '🚀 Start Batch Upload?',
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
  
  toast_(`🚀 Batch upload started! Check progress with "📊 MONITORING → Check Batch Status"`);
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
    
    ul.getRange(rowNum, 5).setValue('⏳ Processing...');
    
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
    throw new Error('⚠️ No field selected.');
  }
  
  const lastRow = ul.getLastRow();
  if (lastRow < 8) {
    throw new Error('⚠️ No data to retry.');
  }
  
  const data = ul.getRange(8, 1, lastRow - 7, 5).getValues();
  const ciqToBobMap = buildCiqToBobMap_();
  
  let listMap = null;
  if (listName) {
    listMap = buildListLabelToId_(listName);
  }
  
  const field = { jsonPath: fieldPath };
  
  let ok = 0, skip = 0, fail = 0;
  let foundTerminated = 0;
  
  for (let i = 0; i < data.length; i++) {
    const rowNum = i + 8;
    const status = normalizeBlank_(data[i][4]);
    
    if (status !== 'FAILED') continue;
    
    const ciq = normalizeBlank_(data[i][0]);
    const rawNew = normalizeBlank_(data[i][1]);
    
    if (!ciq) continue;
    
    ul.getRange(rowNum, 5).setValue('⏳ Retrying...');
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
    
    const wasInOriginal = buildCiqToBobMap_()[ciq];
    if (!wasInOriginal && bobId) {
      foundTerminated++;
      ul.getRange(rowNum, 3).setNote('🔍 Found in terminated employees');
    }
    
    ul.getRange(rowNum, 3).setValue(bobId);
    ul.getRange(rowNum, 4).setValue(field.jsonPath);
    
    let newVal = rawNew;
    if (listMap) {
      const hit = listMap[rawNew] || listMap[String(rawNew).toLowerCase()];
      if (hit) newVal = hit;
    }
    
    const body = buildPutBody_(field.jsonPath, newVal);
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
      const verified = readBackField_(auth, bobId, field.jsonPath);
      writeUploaderResult_(ul, rowNum, 'COMPLETED', code, '', verified);
      ok++;
    } else if (code === 304) {
      const verified = readBackField_(auth, bobId, field.jsonPath);
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
  
  let msg = `🔄 Retry Complete!\n\n✅ Completed: ${ok}\n⏭️ Skipped: ${skip}\n❌ Failed: ${fail}`;
  if (foundTerminated > 0) {
    msg += `\n\n🔍 Found ${foundTerminated} terminated employees!`;
  }
  
  SpreadsheetApp.getUi().alert('Retry Results', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function checkBatchStatus() {
  const props = PropertiesService.getScriptProperties();
  const stateJson = props.getProperty('BATCH_UPLOAD_STATE');
  
  if (!stateJson) {
    SpreadsheetApp.getUi().alert(
      'ℹ️ No Batch Running',
      'No batch upload is currently running.\n\nStart one with "🚀 UPLOAD → Batch Upload"',
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
    '📊 Batch Upload Status',
    `Field: ${state.fieldName}\n\n` +
    `Progress: ${completed}/${totalRows} rows (${progress}%)\n` +
    `Remaining: ${remaining} rows\n\n` +
    `Results So Far:\n` +
    `✅ Completed: ${stats.completed}\n` +
    `⏭️ Skipped: ${stats.skipped}\n` +
    `❌ Failed: ${stats.failed}\n\n` +
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
    toast_('🛡️ Batch upload stopped and cleared.');
  } else {
    toast_('ℹ️ No batch upload was running.');
  }
}

function clearUploadDataAfterBatch_(stats) {
  const ul = getOrCreateSheet_(SHEET_UPLOADER);
  const lastRow = ul.getLastRow();
  
  if (lastRow <= 7) return;
  
  const dataRows = lastRow - 7;
  
  const response = SpreadsheetApp.getUi().alert(
    '✅ Batch Upload Complete!',
    `Final Results:\n\n` +
    `✅ Completed: ${stats.completed}\n` +
    `⏭️ Skipped: ${stats.skipped}\n` +
    `❌ Failed: ${stats.failed}\n\n` +
    `Clear ${dataRows} rows of upload data?\n` +
    `(Field selection will be kept for next upload)`,
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  
  if (response === SpreadsheetApp.getUi().Button.YES) {
    ul.getRange(8, 1, dataRows, ul.getMaxColumns()).clearContent().clearFormat();
    
    if (dataRows > 100) {
      ul.deleteRows(8, dataRows);
    }
    
    toast_('✅ Upload data cleared. Ready for next upload!');
  }
}

/* =====================================================
 * UNIFIED CLEAR FUNCTION (ENHANCED)
 * ===================================================== */
function clearAllUploadData() {
  const ss = SpreadsheetApp.getActive();
  const regularSheet = ss.getSheetByName(SHEET_UPLOADER);
  const historySheet = ss.getSheetByName(CONFIG.HISTORY_SHEET);
  
  let clearedSheets = [];
  let totalRowsCleared = 0;
  
  // Clear regular uploader - enhanced cleanup
  if (regularSheet && regularSheet.getLastRow() > 7) {
    const dataRows = regularSheet.getLastRow() - 7;
    const maxCols = regularSheet.getMaxColumns();
    
    // Clear content, formatting, notes, and validations
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
  
  // Clear history uploader - enhanced cleanup
  if (historySheet && historySheet.getLastRow() > 13) {
    const dataRows = historySheet.getLastRow() - 13;
    const maxCols = historySheet.getMaxColumns();
    
    // Clear content, formatting, notes, and validations
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
      '✅ Already Clean',
      'No upload data to clear.\n\nBoth uploader sheets are empty and ready to use.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert(
      '✅ Data Cleared',
      `Cleared ${totalRowsCleared} rows from:\n• ${clearedSheets.join('\n• ')}\n\n` +
      `Removed all notes, formatting, and validations.\n` +
      `Field/table selections preserved.\n` +
      `Ready for new uploads!`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/* END OF MAIN FILE */
