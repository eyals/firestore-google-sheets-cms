/** ================== FIRESTORE UTILS ==================  */

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