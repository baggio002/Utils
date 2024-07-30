/**
 * input A1, B2
 * return A1:B2
 */
function makeRangeStr(start, end) {
  return start + ":" + end;
}

/**
 * input A1, B, 2
 * return A1:B2
 */
function makeRangeWithLastColAndLastRow(start, col, row) {
  return start + ":" + col + row;
}

/**
 * input sheet, A1
 * return A1:B2 (B2 is the last col and last row)
 */
function makeRangeWithoutLastCol(sheet, start) {
  return makeRangeStr(start, getLastCol(sheet) + getLastRow(sheet))
}

/**
 * input sheet, A1:B
 * return A1:B2 (2 is the last row)
 */
function makeRangeWithoutLastRow(sheet, start) {
  return start + getLastRow(sheet);
}

/**
 * input spreadsheet id, sheet name, A1:B
 * return A1:B2 (2 is the last row)
 */
function makeRemoteRangeWithoutLastRow(ss_id, sheet, start) {
  // Logger.log(ss_id + " sheet = " + sheet + " " + start);
  return start + getLastRowRemote(ss_id, sheet);
}

/**
 * input spreadsheet id, sheet name
 * return [] (the last row raw data)
 */
function getLastRowRemote(ss_id, sheet) {
  let row = 0;
  // Logger.log(ss_id + ' ' + sheet + ' ' + SpreadsheetApp.openById(ss_id).getSheetByName(sheet));
  row = SpreadsheetApp.openById(ss_id).getSheetByName(sheet).getLastRow();
  return row;
}

/**
 * return last row number
 */
function getLastRow(sheet) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getLastRow();
}

/**
 * return last column number(int)
 */
function getLastCol(sheet) {
  return convertNum2Letter(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getLastColumn());
}

/**
 * input sheet name, A1:B2, [][]
 * it will update [][] to the sheet
 */
function exportRawDataToSheet(sheet, range, raws) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getRange(range).setValues(raws);
}

/**
 * append the row to the tail(for remote spreadsheet)
 */
function appendRowToRemoteSheet(ss_id, sheet, row) {
  SpreadsheetApp.openById(ss_id).getSheetByName(sheet).appendRow(row);
}

/**
 * append the row to the tail
 */
function appendRowToSheet(sheet, row) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).appendRow(row);
}

/**
 * append the multiple rows(raws is [][]) to the tail
 */
function appendRawsToSheet(sheet, raws) {
  raws.forEach(
    row => {
      appendRowToSheet(sheet, row);
    }
  );
}

/**
 * set raws to spreadsheet
 */
function setRemoteValues(ss_id, sheet, range, raws) {
  SpreadsheetApp.openById(ss_id).getSheetByName(sheet).getRange(range).setValues(raws);
}

/**
 * clear a range for a spreadsheet
 */
function remoteClear(ss_id, sheet, range) {
  Logger.log('id = ' + ss_id + ' sheet = ' + sheet + ' range = ' + range);
  SpreadsheetApp.openById(ss_id).getSheetByName(sheet).getRange(range).clear();
}

/**
 * clear sheet of a spreadsheet, the range is like A1:B
 */
function remoteClearNonLastRow(ss_id, sheet, range) {
  let lastRow = getLastRowRemote(ss_id, sheet);
  Logger.log('id = ' + ss_id + ' sheet = ' + sheet + ' range = ' + (range + lastRow));
  if (lastRow != 0) {
    SpreadsheetApp.openById(ss_id).getSheetByName(sheet).getRange(range + lastRow).clear();
  }
}

/**
 * clear a range of the sheet
 */
function clear(sheet, range) {
  // Logger.log(sheet + " " + range);
  try {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getRange(range).clearDataValidations().clear();
  } catch(e) {
    Logger.log('Utils clear => sheet = ' + sheet + ' range = ' + range);
    Logger.log(' e = ' + e);
  }
}

/**
 * get [][] from a sheet, A1:B2
 */
function getValues(sheet, range) {
  let values = [];
  Logger.log(range + ' ==== ' + sheet);
  try {
    values = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getRange(range).getValues();
  } catch (e) {
    Logger.log(e);
  } finally {
    return values;
  }
  
  // return values;
}

/**
 * get all data [][] from a sheet, range is A1:B
 */
function getValuesWithNonLastRow(sheet, range) {
  let values = [];
  let lastRow = makeRangeWithoutLastRow(sheet, range);
  values = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getRange(lastRow).getValues();    
  return values;
}

/**
 * get data [][] from a spreadsheet
 */
function getRemoteValues(ss_id, sheet, range) {
  let raw = [];
  try {
    raw = SpreadsheetApp.openById(ss_id).getSheetByName(sheet).getRange(range).getValues();
  } catch (e) {
    Logger.log(e);
  } finally {
    return raw;
  }
}

/**
 * get all data [][] from a spreadsheet, the range is A1:B
 */
function getRemoteValueWithNonLastRowRange(ss_id, sheet, start) {
  let values = [];
  let lastRow = makeRemoteRangeWithoutLastRow(ss_id, sheet, start);
  // Logger.log('lastRow = ' + lastRow);
  values = getRemoteValues(ss_id, sheet, makeRemoteRangeWithoutLastRow(ss_id, sheet, start));
  return values;
}

/**
 * set a cell value
 */
function setCellValue(sheet, range, value) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getRange(range).getCell(1, 1).setValue(value);
}

/**
 * set a cell value to a spreadsheet
 */
function setRemoteCellValue(ss_id, sheet, range, value) {
  SpreadsheetApp.openById(ss_id).getSheetByName(sheet).getRange(range).getCell(1, 1).setValue(value);
}

/**
 * hide a column
 */
function hideCol(sheet, colIndex) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).hideColumns(colIndex);
}

/**
 * show a column
 */
function showCol(sheet, colIndex) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).showColumns(colIndex);
}



