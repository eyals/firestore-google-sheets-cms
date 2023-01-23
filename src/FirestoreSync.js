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
    fs.updateDocument(sheetName() + "/" + rowId, rowToFsDocObject(rowValues),true);
    docsUpdated++;
  });

  //Load current FS documents
  const fsDocs = fs.getDocuments(sheetName());

  //Add to sheet the FS documents that are missing
  var idColumnValues = sheet().getRange(2, headerColumnNum("_id"), sheet().getLastRow() - 2 + 1, 1).getValues();
  var existingIds = idColumnValues.map(idRangeValues => idRangeValues[0].toString());//since each row is an array
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
