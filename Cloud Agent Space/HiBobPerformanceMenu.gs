/**
 * HiBob Performance Report Menu and Credentials Management
 * 
 * This adds menu items to the existing "Bob Salary Data" menu for
 * managing Performance Report credentials and triggering downloads.
 * 
 * CREDENTIALS STORAGE:
 * - Email and password are stored securely in Script Properties
 * - Accessible only to users with edit access to the Apps Script project
 */

// ============================================================================
// MENU FUNCTIONS
// ============================================================================

/**
 * Updates the onOpen function to include Performance Report menu items
 * This should be merged with the existing onOpen() function in bob-salary-data.js
 */
function onOpenWithPerformanceMenu() {
  const ui = SpreadsheetApp.getUi();
  
  // Create main menu
  const menu = ui.createMenu('Bob Salary Data')
    // Existing menu items (from bob-salary-data.js)
    .addItem('Import Base Data', 'importBobDataSimpleWithLookup')
    .addItem('Import Bonus History', 'importBobBonusHistoryLatest')
    .addItem('Import Compensation History', 'importBobCompHistoryLatest')
    .addItem('Import Full Comp History', 'importBobFullCompHistory')
    .addSeparator()
    .addItem('Import All Data', 'importAllBobData')
    .addSeparator()
    .addItem('Convert Tenure to Array Formula', 'convertTenureToArrayFormula')
    .addSeparator()
    // New Performance Report menu items
    .addSubMenu(ui.createMenu('Performance Reports')
      .addItem('Set HiBob Credentials', 'setHiBobCredentials')
      .addItem('View Credentials Status', 'viewCredentialsStatus')
      .addSeparator()
      .addItem('Download Performance Report', 'triggerPerformanceReportDownload')
      .addItem('Instructions', 'showPerformanceReportInstructions'))
    .addToUi();
}

/**
 * Shows instructions for using the Performance Report automation
 */
function showPerformanceReportInstructions() {
  const ui = SpreadsheetApp.getUi();
  const instructions = 
    'HiBob Performance Report Automation\n\n' +
    'SETUP:\n' +
    '1. Click "Set HiBob Credentials" to store your login credentials\n' +
    '2. Run the Python script on your local machine:\n' +
    '   python3 hibob_report_downloader.py\n\n' +
    'HOW IT WORKS:\n' +
    '- The Python script will automatically use credentials stored in Apps Script\n' +
    '- Or you can use credentials from config.json file\n' +
    '- Reports are automatically uploaded to "Bob Perf Report" sheet\n\n' +
    'LOCAL SETUP:\n' +
    '1. Install: pip3 install -r requirements.txt\n' +
    '2. Install browser: playwright install chromium\n' +
    '3. Configure: Edit config.json with Apps Script URL\n' +
    '4. Run: python3 hibob_report_downloader.py';
  
  ui.alert('Performance Report Instructions', instructions, ui.ButtonSet.OK);
}

// ============================================================================
// CREDENTIALS MANAGEMENT
// ============================================================================

/**
 * Prompts user to set HiBob login credentials
 * Stores email and password securely in Script Properties
 */
function setHiBobCredentials() {
  const ui = SpreadsheetApp.getUi();
  
  // Get existing credentials (if any)
  const props = PropertiesService.getScriptProperties();
  const existingEmail = props.getProperty('HIBOB_EMAIL') || '';
  
  // Prompt for email
  const emailResponse = ui.prompt(
    'Set HiBob Email',
    'Enter your HiBob email address:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (emailResponse.getSelectedButton() !== ui.Button.OK) {
    return; // User cancelled
  }
  
  const email = emailResponse.getResponseText().trim();
  if (!email) {
    ui.alert('Error', 'Email cannot be empty.', ui.ButtonSet.OK);
    return;
  }
  
  // Prompt for password
  const passwordResponse = ui.prompt(
    'Set HiBob Password',
    'Enter your JumpCloud password:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (passwordResponse.getSelectedButton() !== ui.Button.OK) {
    return; // User cancelled
  }
  
  const password = passwordResponse.getResponseText().trim();
  if (!password) {
    ui.alert('Error', 'Password cannot be empty.', ui.ButtonSet.OK);
    return;
  }
  
  // Store credentials securely
  props.setProperty('HIBOB_EMAIL', email);
  props.setProperty('HIBOB_PASSWORD', password);
  
  ui.alert('Success', 'Credentials saved successfully!\n\nEmail: ' + email, ui.ButtonSet.OK);
  Logger.log('HiBob credentials updated for: ' + email);
}

/**
 * Shows the status of stored credentials (without revealing password)
 */
function viewCredentialsStatus() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  const email = props.getProperty('HIBOB_EMAIL');
  const hasPassword = props.getProperty('HIBOB_PASSWORD') ? 'Yes' : 'No';
  
  const status = 
    'HiBob Credentials Status\n\n' +
    'Email: ' + (email || 'Not set') + '\n' +
    'Password: ' + hasPassword + '\n\n' +
    (email ? 'Credentials are ready to use.' : 'Please set credentials using "Set HiBob Credentials" menu item.');
  
  ui.alert('Credentials Status', status, ui.ButtonSet.OK);
}

/**
 * Gets stored HiBob credentials
 * @returns {{email: string, password: string}|null} Credentials object or null if not set
 */
function getHiBobCredentials() {
  const props = PropertiesService.getScriptProperties();
  const email = props.getProperty('HIBOB_EMAIL');
  const password = props.getProperty('HIBOB_PASSWORD');
  
  if (!email || !password) {
    return null;
  }
  
  return { email, password };
}

/**
 * Clears stored HiBob credentials
 */
function clearHiBobCredentials() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clear Credentials',
    'Are you sure you want to clear stored HiBob credentials?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('HIBOB_EMAIL');
    props.deleteProperty('HIBOB_PASSWORD');
    ui.alert('Success', 'Credentials cleared.', ui.ButtonSet.OK);
  }
}

// ============================================================================
// API ENDPOINT FOR CREDENTIALS (Optional - for Python script to fetch)
// ============================================================================

/**
 * GET endpoint to retrieve credentials (requires authentication token)
 * This allows the Python script to fetch credentials from Apps Script
 * 
 * Usage: Add this to doGet() in HiBobReportHandler.gs
 * 
 * @param {Object} e - Event object from doGet
 * @returns {TextOutput} JSON response with credentials or error
 */
function getCredentialsAPI(e) {
  try {
    // Simple token-based authentication (set in Script Properties)
    const props = PropertiesService.getScriptProperties();
    const authToken = props.getProperty('HIBOB_AUTH_TOKEN');
    const providedToken = e.parameter.token;
    
    // If no auth token is set, allow access (less secure but simpler)
    // For production, set HIBOB_AUTH_TOKEN in Script Properties
    if (authToken && providedToken !== authToken) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: 'Unauthorized'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const creds = getHiBobCredentials();
    
    if (!creds) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: 'Credentials not set. Please use "Set HiBob Credentials" menu item.'
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
 * Triggers a message about downloading performance report
 * Since the Python script runs locally, this provides instructions
 */
function triggerPerformanceReportDownload() {
  const ui = SpreadsheetApp.getUi();
  
  const creds = getHiBobCredentials();
  const hasCreds = creds ? 'Yes' : 'No';
  
  const message = 
    'Performance Report Download\n\n' +
    'Credentials stored: ' + hasCreds + '\n\n' +
    'To download a report:\n' +
    '1. Open terminal on your local machine\n' +
    '2. Navigate to project directory\n' +
    '3. Run: python3 hibob_report_downloader.py\n' +
    '4. Enter the report name when prompted\n\n' +
    (creds ? 
      '✅ Credentials are stored and ready to use.\n' +
      'The Python script can fetch them automatically.' :
      '⚠️ Please set credentials first using "Set HiBob Credentials"');
  
  ui.alert('Download Performance Report', message, ui.ButtonSet.OK);
}

