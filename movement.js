/**
 * Creates a custom menu in Google Sheets when the spreadsheet opens.
 */
/**
 * Creates a custom menu in Google Sheets when the spreadsheet opens.
 */


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(' Inventory Movement Tools')
    .addItem('1. Log Quantity In/Out (Single Cell)', 'logMovement')
    .addItem('2. Log All Rows Q:AP', 'logAllMovements')
    .addToUi();
}

function logAllMovements() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getActiveSheet();
  var logSheet = ss.getSheetByName("Movements");
  if (!logSheet) {
    SpreadsheetApp.getUi().alert("'Movements' tab not found! Please create it first.");
    return;
  }
  var startRow = 5;
  var startCol = 17; // Q
  var endCol = 42;   // AP
  var lastRow = mainSheet.getLastRow();
  var numRows = lastRow - startRow + 1;
  var numCols = endCol - startCol + 1;
  if (numRows <= 0) {
    SpreadsheetApp.getUi().alert("No data rows found.");
    return;
  }
  var headers = mainSheet.getRange(4, startCol, 1, numCols).getValues()[0];
  var itemIds = mainSheet.getRange(startRow, 1, numRows, 1).getValues();
  var colFValues = mainSheet.getRange(startRow, 6, numRows, 1).getValues();
  var stockValues = mainSheet.getRange(startRow, 12, numRows, 3).getValues();
  var timelineValues = mainSheet.getRange(startRow, startCol, numRows, numCols).getValues();
  var output = [];
  for (var r = 0; r < numRows; r++) {
    var itemId = itemIds[r][0];
    if (!itemId) continue;
    var staticStockSum =
      (Number(stockValues[r][0]) || 0) +
      (Number(stockValues[r][1]) || 0) +
      (Number(stockValues[r][2]) || 0);
    var previousCellVal = 0;
    var foundPrevious = false;
    for (var c = 0; c < numCols; c++) {
      var dateVal = headers[c];
      if (!dateVal) continue;
      var currentCellVal = Number(timelineValues[r][c]) || 0;
      var additionalQty = getQtyFromColumnFForTargetColumn(
        colFValues[r][0],
        dateVal,
        ss.getSpreadsheetTimeZone()
      );
      var adjustedCurrentCellVal = currentCellVal + additionalQty;
      var xValue = 0;
      if (c === 0) {
        if (adjustedCurrentCellVal === 0) {
          xValue = 0;
        } else {
          xValue = staticStockSum + adjustedCurrentCellVal;
          previousCellVal = adjustedCurrentCellVal;
          foundPrevious = true;
        }
      } else {
        if (adjustedCurrentCellVal === 0) {
          xValue = 0;
        } else {
          if (foundPrevious) {
            xValue = adjustedCurrentCellVal - previousCellVal;
          } else {
            xValue = adjustedCurrentCellVal - staticStockSum;
          }
          previousCellVal = adjustedCurrentCellVal;
          foundPrevious = true;
        }
      }
      output.push([dateVal, itemId, xValue]);
      /** *if (xValue !== 0) { output.push([dateVal, itemId, xValue]); }*/ 
    }
  }
  if (output.length > 0) {
    logSheet
      .getRange(logSheet.getLastRow() + 1, 1, output.length, 3)
      .setValues(output);
  }

      SpreadsheetApp.getUi().alert(
    "Batch completed. Logged " + output.length + " movement rows."
  );
}
  


/**
 * Automatically calculates and logs the inventory adjustment (x) to the Movements tab
 * based on the updated cumulative addition/subtraction formula.
 */
function logMovement() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getActiveSheet();
  var activeCell = mainSheet.getActiveCell();
  var col = activeCell.getColumn();
  var row = activeCell.getRow();
  var colFValue = mainSheet.getRange(row, 6).getValue();
  

  
  // Validation: Ensure user selected a cell within the timeline grid (Column Q onwards, Row 5 onwards)
  if (col < 17 || row < 5) {
    SpreadsheetApp.getUi().alert("Please select a cell inside your weekly timeline grid (Column Q onwards, Row 5 onwards).");
    return;
  }
  
  var dateVal = mainSheet.getRange(4, col).getValue();
  var itemId = mainSheet.getRange(row, 1).getValue(); // Assumes Item ID is in Column A
  
  if (!dateVal) {
    SpreadsheetApp.getUi().alert("Could not find a date in Row 4 for this column.");
    return;
  }
  
  // 1. DYNAMICALLY FETCH & SUM COLUMNS L, M, AND N FOR THE ACTIVE ROW
  // Column L is the 12th column. We read 1 row and 3 columns wide (L, M, N).
  var stockValues = mainSheet.getRange(row, 12, 1, 3).getValues()[0];
  var staticStockSum = (Number(stockValues[0]) || 0) + 
                       (Number(stockValues[1]) || 0) + 
                       (Number(stockValues[2]) || 0);
  
  // 2. GET CURRENT CELL VALUE (Defaults to 0 if empty)
  var currentCellVal = Number(mainSheet.getRange(row, col).getValue()) || 0;
  
  // 3 & 4. CALCULATE CUMULATIVE VALUE (xValue)
  var xValue = 0;
  
  if (col === 17) {
    // ---- COLUMN Q (17) LOGIC ----
    if (currentCellVal === 0) {
      xValue = 0;
    } else {
      xValue = staticStockSum + currentCellVal;
    }
    
  } else {
    // ---- COLUMN R (18) AND BEYOND LOGIC ----
    if (currentCellVal === 0) {
      xValue = 0;
    } else {
      var previousCellVal = 0;
      var foundPrevious = false;
      
      // Fetch all previous timeline cells in this row up to the current column
      // This is much faster than checking cells one-by-one
      var historyRange = mainSheet.getRange(row, 17, 1, col - 17).getValues()[0];
      
      // Read backwards to find the most recent non-zero entry
      for (var i = historyRange.length - 1; i >= 0; i--) {
        var checkVal = Number(historyRange[i]) || 0;
        if (checkVal !== 0) {
          previousCellVal = checkVal;
          foundPrevious = true;
          break;
        }
      }
      
      if (foundPrevious) {
        // x2 logic: There is an older cell (Current Cell1). Current is Current Cell2.
        xValue = currentCellVal - previousCellVal;
      } else {
        // x1 logic: No previous cells found. This is the first entry (Current Cell1).
        xValue = currentCellVal - staticStockSum;
      }
    }
  }
  // New logic: add quantities from column F if their dates align with this column
  var colFValue = mainSheet.getRange(row, 6).getValue(); // Column F
  var additionalQty = getQtyFromColumnFForTargetColumn(colFValue, dateVal, ss.getSpreadsheetTimeZone());
  xValue += additionalQty;
  // 5. APPEND TO MOVEMENTS TAB
  var logSheet = ss.getSheetByName("Movements");
  if (!logSheet) {
    SpreadsheetApp.getUi().alert("'Movements' tab not found! Please create it first.");
    return;
  }
  
  logSheet.appendRow([dateVal, itemId, xValue]);
  SpreadsheetApp.getUi().alert(
    "Successfully calculated and logged " +
    xValue +
    " for Item \"" +
    itemId +
    "\". Added Column F quantity: " +
    additionalQty
  );

  // Confirmation alert to show what was calculated
  var formattedDate = Utilities.formatDate(new Date(dateVal), ss.getSpreadsheetTimeZone(), "MM/dd");
  SpreadsheetApp.getUi().alert("Successfully calculated and logged " + xValue + " for Item \"" + itemId + "\" on " + formattedDate + "!");
}

/**
 * Reads Column F format like:
 * 2026/07/02*117;2026/07/09*140;2026/07/16*234
 *
 * Then checks whether each date belongs to the selected target column header.
 */
function getQtyFromColumnFForTargetColumn(colFValue, targetHeader, timezone) {
  if (!colFValue) return 0;
  var total = 0;
  var entries = String(colFValue).split(";");

  for (var i = 0; i < entries.length; i++) {
    var parts = entries[i].trim().split("*");
    if (parts.length !== 2) continue;

    var entryDate = parseDateText(parts[0].trim());
    var qty = Number(parts[1].trim()) || 0;

    if (!entryDate || qty === 0) continue;

    if (dateBelongsToHeader(entryDate, targetHeader, timezone)) {
      total += qty;
    }
  }
  return total;
}

/**
 * Supports:
 * - real date headers, e.g. 2026/07/02
 * - text weekly headers, e.g. 07/08~07/14
 * - Over09/22
 */
function dateBelongsToHeader(entryDate, header, timezone) {
  var entry = clearTime(entryDate);
  var headerText = String(header).trim();
  // Case 1: actual Google Sheets date value
  if (Object.prototype.toString.call(header) === "[object Date]" && !isNaN(header.getTime())) {
    return entry.getTime() === clearTime(header).getTime();
  }
  // Case 2: weekly range like 07/08~07/14
  if (headerText.indexOf("~") !== -1) {
    var parts = headerText.split("~");
    var start = parseMonthDay(parts[0]);
    var end = parseMonthDay(parts[1]);
    if (!start || !end) return false;
    return entry.getTime() >= start.getTime() && entry.getTime() <= end.getTime();
  }
  // Case 3: Over09/22
  if (headerText.indexOf("Over") === 0) {
    var cutoffText = headerText.replace("Over", "").trim();
    var cutoff = parseMonthDay(cutoffText);
    if (!cutoff) return false;
    return entry.getTime() > cutoff.getTime();
  }
  // Case 4: text date header like 2026/07/02
  var fullDateHeader = parseDateText(headerText);
  if (fullDateHeader) {
    return entry.getTime() === clearTime(fullDateHeader).getTime();
  }
  return false;
}
function parseDateText(text) {
  text = String(text).trim();
  // Must be exactly yyyy/mm/dd
  if (!/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(text)) {
    return null;
  }
  var parts = text.split("/");
  var year = Number(parts[0]);
  var month = Number(parts[1]);
  var day = Number(parts[2]);
  var d = new Date(year, month - 1, day);
  if (isNaN(d.getTime())) return null;
  return d;
}


/**
 * Parses MM/DD using year 2026.
 * Adjust year if your planning file changes year.
 */
function parseMonthDay(text) {
  var parts = text.trim().split("/");
  if (parts.length !== 2) return null;
  
  var month = Number(parts[0]) - 1;
  var day = Number(parts[1]);
  
  return new Date(2026, month, day);
}

function clearTime(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}



