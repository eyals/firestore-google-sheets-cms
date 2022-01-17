/**  ================== STYLE ================== */

/** On every edit, makes changes to the styling of the content range, and sets a checkbox */
function formatCells() {

  /** Conerting first row to checkbox */
  var range = sheetContentRange();
  for (row = range.getRow(); row < range.getRow() + range.getNumRows(); row++) {
    setActiveCheckbox(row);
  }

  /** Applying conditional formatting */
  var rules = [];//sheet().getConditionalFormatRules();

  var markEmpty = SpreadsheetApp.newConditionalFormatRule()
    .whenCellEmpty()
    .setBackground("#DDDDDD")
    .setRanges([range])
    .build();

  var dimInactive = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2=false')
    .setBackground("#EEEEEE")
    .setFontColor("#666666")
    // .setStrikethrough(true)
    .setRanges([range])
    .build();

  var boldActive = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2=true,ISNUMBER(SEARCH("_id",A$1)))')
    .setBackground("#FFFFFF")
    .setFontColor("#000000")
    .setBold(true)
    .setRanges([range])
    .build();

  var markMissingMandatory = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISBLANK(A2), OR(ISNUMBER(SEARCH("~*", A$1)),A$1="_id"), $A2=TRUE)')
    .setBackground("#FEC4BC")
    .setFontColor("#9F1B09")
    .setRanges([range])
    .build();

  var markArrays = SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith("[")
    .whenTextEndsWith(']')
    .setFontColor("#9B4FFC")
    .setRanges([range])
    .build();

  rules.push(markMissingMandatory);
  rules.push(dimInactive);
  rules.push(markArrays);
  rules.push(boldActive);
  rules.push(markEmpty);

  sheet().setConditionalFormatRules(rules);
}
