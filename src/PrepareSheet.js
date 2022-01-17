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
}