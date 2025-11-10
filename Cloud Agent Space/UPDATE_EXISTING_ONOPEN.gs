/**
 * UPDATED onOpen() Function
 * 
 * Replace the existing onOpen() function in bob-salary-data.js with this version
 * to add Performance Report menu items.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Bob Salary Data')
    // Existing menu items
    .addItem('Import Base Data', 'importBobDataSimpleWithLookup')
    .addItem('Import Bonus History', 'importBobBonusHistoryLatest')
    .addItem('Import Compensation History', 'importBobCompHistoryLatest')
    .addItem('Import Full Comp History', 'importBobFullCompHistory')
    .addSeparator()
    .addItem('Import All Data', 'importAllBobData')
    .addSeparator()
    .addItem('Convert Tenure to Array Formula', 'convertTenureToArrayFormula')
    .addSeparator()
    // NEW: Performance Report submenu
    .addSubMenu(ui.createMenu('Performance Reports')
      .addItem('Set HiBob Credentials', 'setHiBobCredentials')
      .addItem('View Credentials Status', 'viewCredentialsStatus')
      .addSeparator()
      .addItem('Download Performance Report', 'triggerPerformanceReportDownload')
      .addItem('Instructions', 'showPerformanceReportInstructions'))
    .addToUi();
}

