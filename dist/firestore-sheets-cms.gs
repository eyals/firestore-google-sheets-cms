/** ================== FIRESTORE SYNC ================== */


/**
 * 1. Goes through the sheet rows and:
 *    - If a mandatory field is missing - skips row
 *    - If the row is active - updates in Firestore
 *    - If the row is inactive - deletes it from Firestore
 * 2. Goes through the existing docs in Firestore, looking for IDs that aren't in the sheet, and adding them to the sheet
 */

function syncWithFirestore() {

  if (sheetContentRange() == null) return;
  var headers = getSheetHeaders();
  if (headers.length == 0) return;

  if (headers.findIndex(h => h.label == "_id") < 0 || headers.findIndex(h => h.label == "_active") < 0) {
    sheetsUi().alert("Missing colomns _id or _active. Use 'Prepare sheet' from menu to fix this.");
    return;
  }



  var docsUpdated = 0;
  var docsDeleted = 0;
  var docsAdded = 0;
  var missingIds = 0;
  var missingMandatory = 0;

  var fs = firestore();
  if (fs == null) return;

  // Load sheet content 
  var sheetValues = sheetContentRange().getValues();
  var mandatoryFields = headers.filter(header => header.isMandatory).map(header => header.label);

  // Update FS documents from sheet
  sheetValues.forEach((rowValues) => {

    var rowObject = {};
    for (var i = 0; i < rowValues.length; i++) {
      rowObject[headers[i].label] = rowValues[i]
    };
    var rowId = rowObject["_id"];

    //skipping a row if ID is missing
    if (rowId == "") {
      missingIds++;
      return;
    }

    //skipping it a mandatory field is empty
    var missingMandatoryFound = false;
    mandatoryFields.forEach((fieldName) => {
      // var fieldValue = rowValues[headerColumnNum(fieldName) - 1];
      if (rowObject[fieldName] == "") {
        missingMandatoryFound = true;
      }
    });
    if (missingMandatoryFound && rowObject["_active"] === true) {
      missingMandatory++;
      return;
    }


    //Deleting FS doc if row not marked as active
    // var rowIsActive = rowValues[headerColumnNum("_active") - 1] === true;
    if (!rowObject["_active"] === true) {
      fs.deleteDocument(sheetName() + "/" + rowId);
      docsDeleted++;
      return;
    }

    //Updating/creating FS documents from valid rows
    fs.updateDocument(sheetName() + "/" + rowId, rowToFsDocObject(rowValues));
    docsUpdated++;
  });

  //Load current FS documents
  const fsDocs = fs.getDocuments(sheetName());

  //Add to sheet the FS documents that are missing
  var idColumnValues = sheet().getRange(2, headerColumnNum("_id"), sheet().getLastRow() - 2 + 1, 1).getValues();
  var existingIds = idColumnValues.map(idRangeValues => idRangeValues[0]);//since each row is an array
  fsDocs.forEach((doc) => {
    var docId = doc.name.substring(doc.name.lastIndexOf("/") + 1);
    if (existingIds.indexOf(docId) < 0) {
      addDocFromFirestore(docId, doc.fields);
      docsAdded++;
    }
  });

  formatCells();

  showResultMessage([
    docsDeleted + " deleted",
    docsAdded + " added",
    docsUpdated + " updated",
    missingIds + " missing ID",
    missingMandatory + " missing mandatory fields",
    "<a href='https://console.firebase.google.com/project/" + serviceAccount().projectId + "/firestore/data/~2F" + sheetName() + "' target='_blank'>Open in Firestore</a>",
  ]);

}

/**
 * A doc object look like this:
 * {
 *   "name":"projects/my-firebase-project/(default)/documents/products/op-1,, 
 *   "fields":{
 *          "productName": { "stringValue": "OP-1 Portable Synthesizer" },
 *          "price":  {"integerValue": "143" },
 *          "inSock": { "booleanValue": true },
 *          "categories": { "arrayValue": { "values": [ { "stringValue": "handheld" } ,{ "stringValue": "synth" } ] } }, 
 *         }
 *    "createTime": ...
 * }
 * 
 * Extracts the ID from teh end of the "name" 
 * and the fields from the "fields"
 * and makes the necessary conversions.
 */
function addDocFromFirestore(docId, docFields) {
  var rowData = [];
  var headers = getSheetHeaders();
  headers.forEach((header) => {
    //Sets the row as active
    if (header.label == "_active") {
      rowData.push(true);
      return;
    }
    //Sets the row ID
    if (header.label == "_id") {
      rowData.push(docId);
      return;
    }
    //Sets all other fields that are found in sheet headers
    if (docFields[header.label] != null) {
      rowData.push(fsObjectToCell(docFields[header.label]))
    } else {
      rowData.push("");
    }
  });
  sheet().appendRow(rowData);
}

function showResultMessage(messageRows) {
  var message = messageRows.join("<br/>");
  var htmlOutput = HtmlService.createHtmlOutput("<div style='font-family:monospace'>" + message + "</div>").setTitle('Sync complete');
  sheetsUi().showSidebar(htmlOutput);

}
/** ================== SHEET EVENTS ================== */

function onOpen() {
  createMenu();
  //serviceAccountSheet().hideSheet();
}

function onEdit(e) {
  formatCells();
}/** ================== UTILITIES ================== */

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



/** ================== SERVICE ACCOUNT ================== */

/**
 * Returns the ServiceAccount sheet.
 * If it doesn't exist - creates one.
 */
function serviceAccountSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ServiceAccount");
  if (!sheet) {
    var newSheet = createServiceAccountSheet();
    if (newSheet == null) return null;
    return newSheet;
  } else {
    sheet.hideSheet();
    return sheet;
  }
}

/**
 * Guides the user through a series of questions.
 * The answers are added to a new, hidden sheet, called ServiceAccount,
 * which will be used for Firebase authentication.
 * This event is triggered in these cases:
 * - Trying to sync when there is no Service Account sheet
 * - Clicking the action in the menu (good for replacing the key)
 */
function createServiceAccountSheet() {

  var ui = sheetsUi();

  /** Explains what the process is for */
  var introPrompt = ui.alert('Set up a service account', 'Let\'s set up the connection to your Firebase project.\nNote: This will store your credentials in a hidden tab, so don\'t share this spreadsheet with people you don\'t trust.\n\nClick OK to continue, or Cancel to do it later.', ui.ButtonSet.OK_CANCEL);
  if (introPrompt == ui.Button.CANCEL) return null;

  /** Asks for project ID */
  var projPrompt = ui.prompt('What is your Firebase project ID?', '', ui.ButtonSet.OK_CANCEL);
  if (projPrompt.getSelectedButton() == ui.Button.CANCEL) return;
  var projectId = projPrompt.getResponseText().split("\"").join("");
  if (projectId == "") return;

  /** Instructs how to get the key file */
  var getKeyAlert = ui.alert('Get a private key', '- Visit https://console.firebase.google.com/project/' + projectId + '/settings/serviceaccounts\n- Click "Generate new private key"\n- Open the downloaded file in a text editor.\n- You\'ll need to copy two fields from there in the next screens.\n\nClick OK to contine.', ui.ButtonSet.OK_CANCEL);
  if (getKeyAlert == ui.Button.CANCEL) return null;

  /** Asks for client email */
  var emailPrompt = ui.prompt('Your Service Account email', 'Appears as \'client_email\' in your private key.\nThis is NOT your personal email.', ui.ButtonSet.OK_CANCEL);
  if (emailPrompt.getSelectedButton() == ui.Button.CANCEL) return null;
  var email = emailPrompt.getResponseText().split("\"").join("");
  if (email == "") return null;

  /** Asks for secret key */
  var keyPrompt = ui.prompt('Your Service Account key', 'Appears as \'private_key\' in your private key.\nIt\'s a long blob starting with -----BEGIN PRIVATE KEY----- and ending with -----END PRIVATE KEY-----\\n.', ui.ButtonSet.OK_CANCEL);
  if (keyPrompt.getSelectedButton() == ui.Button.CANCEL) return null;
  var key = keyPrompt.getResponseText().split("\"").join("");
  if (key == "") return null;

  /** Creates the sheet, hides it, and sets the values */
  var newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("ServiceAccount");
  newSheet.hideSheet();
  newSheet.getRange("A1:B5").setValues([
    ["projectId", projectId],
    ["email", email],
    ["key", key],
    ["", ""],
    ["WARNING:", "Do not modify this sheet if you're not sure what this is. This sheet is best as hidden.\nUse the Firestore menu to generate a service account if you fail to connect to Firebase."]
  ]);

  toast("Service account set up and ready to use");
  return newSheet;

}

var serviceAccount = () => {
  var saSheet = serviceAccountSheet();
  if (!saSheet) return null;
  return {
    email: saSheet.getRange("ServiceAccount!B2").getValue(),
    projectId: saSheet.getRange("ServiceAccount!B1").getValue(),
    key: saSheet.getRange("ServiceAccount!B3").getValue().replace(/\\n/g, '\n'),
  }
}/**  ================== STYLE ================== */

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
/** ================== GOOGLE SHEETS MENU ==================  */

function createMenu(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [
    { name: "🔄 Sync sheet with Firestore", functionName: "syncWithFirestore" },
    null,
    { name: "Prepare sheet", functionName: "prepSheet" },
    { name: "Configure service account", functionName: "serviceAccountPrompt" },
    { name: "Help", functionName: "openHelp" },
  ];
  spreadsheet.addMenu("🔥 Firestore", menuEntries);

}

function openHelp() {
  var message = [
    "<a href='https://github.com/myproject' target='_blank'>Open project on Github</a>",
    "<a href='https://youtube.com/myVideo' target='_blank'>Watch video guide</a>",
  ].join("<br/>");
  var html = "<div style='font-family:sans-serif'>" + message + "</div>";
  sheetsUi().showModalDialog(HtmlService.createHtmlOutput(html).setHeight(50).setWidth(250), "Help");
}
/** ================== PREPARE SHEET ================== */

/** Adds the _active and  _id headers, freezes 1st row */
function prepSheet() {

  var headers = getSheetHeaders();
  if (headers.length == 0 || headers[0].label != "_active") {
    sheet().insertColumns(1);
    sheet().getRange("A1:A1").setValues([["_active"]]);
  }

  headers = getSheetHeaders();//assigning again since it has changed
  var idColumnIndex = headers.findIndex((h) => h.label == "_id");
  if (idColumnIndex==-1){
    /** _id field not found? Adding a new column */
    sheet().insertColumns(2)
    sheet().getRange("B1:B1").setValues([["_id"]]);
  }else{
    /** _id field found? Moving it to column 2 */
    var idColumn = idColumnIndex+1;
    var idValues = sheet().getRange(1,idColumn, sheet().getLastRow(),1).getValues();
    sheet().insertColumns(2)
    sheet().getRange(1,2, sheet().getLastRow(),1).setValues(idValues);
    sheet().deleteColumn(idColumn+1);
  }

  sheet().setFrozenRows(1);
  sheet().getRange("A1:AA1").setTextStyle(
    SpreadsheetApp.newTextStyle()
      .setBold(true).build()
  )
}/** ================== FIRESTORE UTILS ==================  */

/** Returns a Firestore object with authentication */
var firestore = () => (serviceAccount() == null)
  ? null
  : FirestoreApp.getFirestore(serviceAccount().email, serviceAccount().key, serviceAccount().projectId);


/**
 * Creates an object with column headers as keys
 * Skips id, active, empty headers, or headers marked for no sync (~)
 */
function rowToFsDocObject(rowData) {
  var headers = getSheetHeaders();
  var docObject = {};
  for (i = 0; i < headers.length; i++) {
    if (!headers[i].isSync) continue;
    var label = headers[i].label;
    //Not adding the _active/_id as doc properties
    if (label == "_id") continue;
    if (label == "_active") continue;
    if (label === "") continue;
    if (rowData[i] !== "") {
      docObject[label] = cellToFsObject(rowData[i]);
    }
  }
  return docObject;
}


/**
 * Converts a firebase object like {"valueType":"valueAsString"} to typedValue.
 * {"stringValue":"abc"} => "abc"
 * {"booleanValue":false} => false
 * {"nullValue":null} => null
 * {"integerValue":"600"} => 600
 * {"arrayValue":{"values":[{"stringValue":"a"},{"stringValue":"b"}]}} => [a,b]
 */
function fsObjectToCell(fieldObj) {
  for (var prop in fieldObj) {
    var origValue = fieldObj[prop];
    var typedValue;
    switch (prop) {
      case "integerValue":
        typedValue = parseInt(origValue);
        break;
      case "arrayValue":
        var arrayValues = origValue.values.map(value => fsObjectToCell(value));
        typedValue = "[" + arrayValues.toString() + "]";
        break;
      default:
        typedValue = origValue;
    }
    break;//Only one property in any object
  }
  return typedValue;
}

/**
 * Converts cell data to a firebase object
 * "a" => "a"
 * 5 => 5
 * "[a,b]"" => ["a","b"]
 * "[1,2]"" => [1,2]
 * TRUE => true
 */
function cellToFsObject(cellData) {

  if (typeof cellData == "string") {
    cellData = cellData.trim();
    if (cellData == "") return null;

    //Detect boolean
    if (cellData.toLowerCase() == "true") return true;
    if (cellData.toLowerCase() == "false") return false;

    //Detect number
    if (!isNaN(cellData)) return parseInt(cellData);

    //Detect arras: String surrounded by [] are converted to an array
    if (/^\[.*\]$/.test(cellData)) {
      cellData = cellData.replace("[", "").replace("]", "").split(",");
      for (m in cellData) {
        cellData[m] = cellToFsObject(cellData[m]);
      }
    }
  }
  return cellData;
}