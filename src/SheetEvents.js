/** ================== SHEET EVENTS ================== */

function onOpen() {
  createMenu();
  //serviceAccountSheet().hideSheet();
}

function onEdit(e) {
  formatCells();
}