/**
 * generate data validation(dropdown menu)
 */
function generateRules(sheet, range, values) {
  let partRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet)
          .getRange(range);
  let partRule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
  partRange.setDataValidation(partRule);
}

/**
 * set color to background, based on specified text
 */
function setBGBasedOnText(sheet, range, text, color) {
  let newRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains(text)
      .setBackground(color)
      .setRanges([SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getRange(range)])
      .build()
  var rules = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getConditionalFormatRules();
  rules.push(newRule);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).setConditionalFormatRules(rules);
}

function setBackgroundColour(sheet, range, rgb) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getRange(range).setBackground(rgb);
}

function setFormula(sheet, range, formula) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getRange(range).setFormula(formula);
}

function generateCheckbox(sheet, range, value) {
  var partRule = SpreadsheetApp.newDataValidation().requireCheckbox(value).build();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getRange(range).setDataValidation(partRule);
}

function setCellBGColour(sheet, range, rgb) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getRange(range).getCell(1, 1).setBackground(rgb);
}
