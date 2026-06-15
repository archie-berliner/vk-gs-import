/**
 * Creates a custom menu in Google Sheets when the spreadsheet opens
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('TakeFive')
      .addItem('📥 Import Wall Posts', 'importVKWall')
      .addItem('📥 Import Followers', 'importVKFollowers')
      .addItem('🧹 Clean CSV Import', 'cleanImportedCSV')
      .addToUi();
}

/**
 * Helper function to run both imports at once
 */
function importBoth() {
  importVKWall();

  SpreadsheetApp.getUi().alert('Import complete!');
}
