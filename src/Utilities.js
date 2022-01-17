/** ================== UTILITIES ================== */

/** Returns the active sheet object */
var sheet = () => SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

/** Returns the name of the active sheet */
var sheetName = () => normalisedString(sheet().getName());

/** returns the range of the content in the active sheet, excluding the headers, or null if empty */
var sheetContentRange = () => sheet().getLastRow() > 1 ? sheet().getRange(2, 1, sheet().getLastRow() - 2 + 1, sheet().getLastColumn()) : null;

/**
 * Returns an array of header titles.
 * Based on the first frozen row.
 * Normalizes headers by removing non alphanumerical values.
 * Includes an indication if a field is mandatory, by checking if it ends with *.
 */
function getSheetHeaders() {
  var firstRowRange = sheet().getRange(1, 1, 1, sheet().getMaxColumns());
  var headerTitlesArray = firstRowRange.getValues()[0];//[0] since getValues returns 2d array
  var headerObjectsArray = []
  headerTitlesArray.forEach(function (headerTitle) {
    if (headerTitle == "") return;
    headerTitle = headerTitle.trim();
    headerObjectsArray.push({
      label: normalisedString(headerTitle),
      isMandatory: (lastCharacter(headerTitle) == "*"),
      isSync: (lastCharacter(headerTitle) !== "~"),
    })
  });
  return headerObjectsArray;
}

/** Returns the column number of a given header */
function headerColumnNum(headerName) {
  return getSheetHeaders().findIndex(h => {
    return h.label === normalisedString(headerName);
  }) + 1;
}

/** Returns the last character of a string. Used to detect header prefix, like * for mandatory */
var lastCharacter = (str) => str.charAt(str.length - 1);

/** Strips any characters other than alphanumeric, _ or - */
var normalisedString = (str) => str.replace(" ", "_").replace(/[^A-Za-z0-9_-]/g, "");

/** Shortcut for the UI actions, like prompt */
var sheetsUi = () => SpreadsheetApp.getUi();

/** Shortcut for Toast UI */
function toast(message, title = "", timeout = 5) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title, timeout);
}

/** Prints a value and its type, either in a sidebar or in the console */
function print(val) {
  try {
    var htmlOutput = HtmlService.createHtmlOutput("<div style='font-family:monospace'>" + JSON.stringify(val, null, 2) + "</div>").setTitle(typeof val);
    sheetsUi().showSidebar(htmlOutput);
  } catch {
    try {
      sheetsUi().alert(typeof val + "\n" + JSON.stringify(val, null, 2));
    } catch {
      console.log(val, typeof val);
    }
  }
};

/** Converts a cell into a checkbox */
function setCheckbox(cell, state = null) {
  var validationRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  cell.setDataValidation(validationRule);
  if (state != null && (state === true || state === false)) {
    cell.setValue(state);
    retuen;
  }
}

/** Adds a checkbox if a row is not empty. Removes the checkbox if the row is empty */
function setActiveCheckbox(rowNumber) {
  //Clearing the checkbox if the row is emoty
  var activeColNum = headerColumnNum("_active");
  if (isNaN(activeColNum)) return;
  var checkboxCell = sheet().getRange(rowNumber, activeColNum);
  var rowContentRange = sheet().getRange(rowNumber, 2, 1, sheet().getLastColumn());//Anything in the row other than the first column
  if (rowContentRange.isBlank()) {
    checkboxCell.clearDataValidations();
    checkboxCell.clear();
    return;
  } else {
    setCheckbox(checkboxCell);
  }
}



