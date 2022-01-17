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
