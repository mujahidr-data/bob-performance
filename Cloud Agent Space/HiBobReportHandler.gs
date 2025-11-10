/**
 * HiBob Report Handler for Google Apps Script
 * 
 * This script receives uploaded files from the local Python automation
 * and imports them into the Google Sheet.
 * 
 * SETUP INSTRUCTIONS:
 * 1. Copy this code to your Google Apps Script project (1Pi7HmLV8K2v8q9ciiuZ1hbvFp1_NCyuM6svz2LsNiAK5oZ1zU-0qVF47)
 * 2. Deploy as Web App:
 *    - Click "Deploy" > "New deployment"
 *    - Select type: "Web app"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone"
 *    - Click "Deploy"
 * 3. Copy the deployment URL and paste it in config.json as "apps_script_url"
 * 
 * SHEET CONFIGURATION:
 * - Target Google Sheet ID: 1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA
 * - Target Sheet Name: "Bob Perf Report"
 */

// Configuration
const SHEET_ID = '1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA';
const TARGET_SHEET_NAME = 'Bob Perf Report';

/**
 * Handle GET requests (for credentials API)
 */
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
    
    // Check if file was uploaded
    if (!e || !e.postData) {
      return createResponse(400, 'No data received');
    }
    
    // Parse the multipart form data
    const boundary = extractBoundary(e.postData.type);
    if (!boundary) {
      return createResponse(400, 'Invalid content type. Expected multipart/form-data');
    }
    
    // Extract file from multipart data
    const fileData = extractFileFromMultipart(e.postData.contents, boundary);
    if (!fileData) {
      return createResponse(400, 'No file found in request');
    }
    
    Logger.log('File received: ' + fileData.filename);
    
    // Import file to Google Sheet
    const result = importFileToSheet(fileData);
    
    if (result.success) {
      return createResponse(200, result.message, {
        rowsImported: result.rowsImported,
        sheetName: TARGET_SHEET_NAME,
        sheetId: SHEET_ID
      });
    } else {
      return createResponse(500, result.message);
    }
    
  } catch (error) {
    Logger.log('Error in doPost: ' + error.toString());
    return createResponse(500, 'Internal error: ' + error.toString());
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
    // Split by boundary
    const parts = contents.split('--' + boundary);
    
    for (let i = 0; i < parts.length; i++) {
      const part = parts[i];
      
      // Look for file upload part
      if (part.indexOf('filename=') !== -1) {
        // Extract filename
        const filenameMatch = part.match(/filename="([^"]+)"/);
        const filename = filenameMatch ? filenameMatch[1] : 'unknown';
        
        // Extract content type
        const contentTypeMatch = part.match(/Content-Type: ([^\r\n]+)/);
        const contentType = contentTypeMatch ? contentTypeMatch[1].trim() : 'application/octet-stream';
        
        // Extract file content (after double CRLF)
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
 * Import file to Google Sheet
 */
function importFileToSheet(fileData) {
  try {
    // Open the target spreadsheet
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    
    // Get or create the target sheet
    let sheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);
    
    if (sheet) {
      // Clear existing data
      sheet.clear();
      Logger.log('Cleared existing sheet: ' + TARGET_SHEET_NAME);
    } else {
      // Create new sheet
      sheet = spreadsheet.insertSheet(TARGET_SHEET_NAME);
      Logger.log('Created new sheet: ' + TARGET_SHEET_NAME);
    }
    
    // Parse the file content based on type
    let data = [];
    const filename = fileData.filename.toLowerCase();
    
    if (filename.endsWith('.csv')) {
      data = parseCSV(fileData.content);
    } else if (filename.endsWith('.xlsx') || filename.endsWith('.xls')) {
      // For Excel files, we need to convert to blob and import
      // Note: This is more complex and may require additional handling
      return {
        success: false,
        message: 'Excel file format not yet supported. Please download as CSV format.'
      };
    } else if (filename.endsWith('.txt')) {
      data = parseCSV(fileData.content); // Try parsing as CSV
    } else {
      // Try to parse as CSV by default
      data = parseCSV(fileData.content);
    }
    
    if (!data || data.length === 0) {
      return {
        success: false,
        message: 'No data found in file or unable to parse file format'
      };
    }
    
    // Write data to sheet
    const numRows = data.length;
    const numCols = data[0].length;
    
    sheet.getRange(1, 1, numRows, numCols).setValues(data);
    
    // Format header row (first row)
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
    
    Logger.log('Successfully imported ' + numRows + ' rows to ' + TARGET_SHEET_NAME);
    
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
 * Parse CSV content into 2D array
 */
function parseCSV(content) {
  try {
    const lines = content.split(/\r?\n/);
    const data = [];
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (line) {
        // Simple CSV parsing (handles basic cases)
        // For more complex CSV (with quotes, commas in fields), this may need enhancement
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
        // Escaped quote
        cell += '"';
        i++; // Skip next quote
      } else {
        // Toggle quote mode
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      // End of cell
      result.push(cell);
      cell = '';
    } else {
      cell += char;
    }
  }
  
  // Add last cell
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

/**
 * API endpoint to get credentials (called by Python script)
 * Requires credentials to be set via menu: Performance Reports > Set HiBob Credentials
 */
function getCredentialsAPI(e) {
  try {
    const props = PropertiesService.getScriptProperties();
    const email = props.getProperty('HIBOB_EMAIL');
    const password = props.getProperty('HIBOB_PASSWORD');
    
    if (!email || !password) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: 'Credentials not set. Please use "Set HiBob Credentials" menu item in Google Sheets.'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      email: email,
      password: password
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
 * Test function - can be run manually to verify setup
 */
function testSheetAccess() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    Logger.log('Successfully accessed spreadsheet: ' + spreadsheet.getName());
    
    let sheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(TARGET_SHEET_NAME);
      Logger.log('Created new sheet: ' + TARGET_SHEET_NAME);
    } else {
      Logger.log('Found existing sheet: ' + TARGET_SHEET_NAME);
    }
    
    // Write test data
    sheet.getRange('A1').setValue('Test successful at ' + new Date());
    
    return 'Test passed!';
  } catch (error) {
    Logger.log('Test failed: ' + error.toString());
    return 'Test failed: ' + error.toString();
  }
}

