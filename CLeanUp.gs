function cleanImportedCSV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("MockDataVKgroup");

  if (!sheet) {
    throw new Error("Sheet 'MockDataVKgroup' not found!");
  }

  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    Logger.log("No data to process");
    return;
  }

  Logger.log("Processing " + lastRow + " rows...");

  // Get only columns A-I (indices 0-9) to avoid touching formulas in other columns
  var data = sheet.getRange(1, 1, lastRow, 9).getValues();
  var headerRow = data[0]; // Store the first row as header (columns A-I)

  // Find the first duplicate header row (compare only columns A-I, indices 0-8)
  var firstDuplicateIndex = -1;
  for (var i = 1; i < data.length; i++) {
    var isHeaderRow = true;
    // Only compare first 9 columns (A-I, indices 0-8)
    for (var j = 0; j < 9 && j < headerRow.length; j++) {
      if (data[i][j] !== headerRow[j]) {
        isHeaderRow = false;
        break;
      }
    }
    if (isHeaderRow) {
      firstDuplicateIndex = i;
      Logger.log("Found first duplicate header at row " + (i + 1));
      break;
    }
  }

  // If no duplicate header found, nothing new was appended - exit
  if (firstDuplicateIndex === -1) {
    Logger.log("No duplicate header found - no new data to process");
    return;
  }

  // Keep all rows before the duplicate header (they're already cleaned)
  var cleanedRows = data.slice(0, firstDuplicateIndex);
  Logger.log("Keeping " + (cleanedRows.length - 1) + " already-cleaned rows");

  // Process only the new appended data (from duplicate header onwards)
  var newData = data.slice(firstDuplicateIndex);
  Logger.log("Processing " + newData.length + " newly appended rows...");

  // Debug: Log a sample row from new data
  if (newData.length > 1) {
    Logger.log("Sample new data row: " + JSON.stringify(newData[1]));
    Logger.log("Column F value: '" + newData[1][5] + "'");
    Logger.log("Column G value: '" + newData[1][6] + "'");
    Logger.log("Column H value: '" + newData[1][7] + "'");
  }

  var deletedCount = 0;
  var deleteReasons = {
    duplicateHeader: 0,
    colFNotMatch: 0,
    colGMatch: 0,
    colHMatch: 0
  };

  // Process new data
  for (var i = 0; i < newData.length; i++) {
    var row = newData[i];

    // Trim values to avoid whitespace issues
    var colF = String(row[5]).trim();  // Column F (index 5) - "Сортировка: гранулярность"
    var colG = String(row[6]).trim();  // Column G (index 6) - "Сортировка: вид разреза"
    var colH = String(row[7]).trim();  // Column H (index 7) - "Параметр легенды"

    var shouldKeep = true;
    var deleteReason = "";

    // Check if this is a duplicate header row (compare only columns A-I)
    var isHeaderRow = true;
    for (var j = 0; j < 9 && j < headerRow.length; j++) {
      if (row[j] !== headerRow[j]) {
        isHeaderRow = false;
        break;
      }
    }

    if (isHeaderRow) {
      shouldKeep = false;
      deleteReason = "duplicateHeader";
      deleteReasons.duplicateHeader++;
    }
    // Check filtering conditions
    else if (colF !== "По дням") {
      shouldKeep = false;
      deleteReason = "colFNotMatch";
      deleteReasons.colFNotMatch++;
      if (i <= 5) { // Log first few for debugging
        Logger.log("Row " + (i+1) + " deleted: Column F is '" + colF + "' (not 'По дням')");
      }
    }
    else if (colG === "Вся аудитория") {
      shouldKeep = false;
      deleteReason = "colGMatch";
      deleteReasons.colGMatch++;
      if (i <= 5) {
        Logger.log("Row " + (i+1) + " deleted: Column G is '" + colG + "'");
      }
    }
    else if (colH === "All content") {
      shouldKeep = false;
      deleteReason = "colHMatch";
      deleteReasons.colHMatch++;
      if (i <= 5) {
        Logger.log("Row " + (i+1) + " deleted: Column H is '" + colH + "'");
      }
    }

    if (shouldKeep) {
      cleanedRows.push(row);
    } else {
      deletedCount++;
    }
  }

  Logger.log("Delete reasons breakdown:");
  Logger.log("  Duplicate headers: " + deleteReasons.duplicateHeader);
  Logger.log("  Column F not 'По дням': " + deleteReasons.colFNotMatch);
  Logger.log("  Column G is 'Вся аудитория': " + deleteReasons.colGMatch);
  Logger.log("  Column H is 'All content': " + deleteReasons.colHMatch);

  var newCleanedRowsCount = cleanedRows.length - firstDuplicateIndex;
  Logger.log("From new data: kept " + newCleanedRowsCount + " rows, deleted " + deletedCount + " rows");
  Logger.log("Total rows after cleanup: " + (cleanedRows.length - 1) + " data rows + 1 header");

  // Delete excess rows beyond what we need
  var rowsNeeded = cleanedRows.length;
  if (lastRow > rowsNeeded) {
    sheet.deleteRows(rowsNeeded + 1, lastRow - rowsNeeded);
  }

  // Write cleaned data back to columns A-I only (preserves formulas in columns N onwards)
  if (cleanedRows.length > 0) {
    sheet.getRange(1, 1, cleanedRows.length, 9).setValues(cleanedRows);
  }

  Logger.log("Cleanup complete!");
  Logger.log("Final row count: " + (cleanedRows.length - 1) + " data rows + 1 header");
}
