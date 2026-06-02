function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  const sheet = getSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error(`Hoja '${name}' no existe. Corre createSheets() primero.`);
  }
  return sheet;
}

function getSheetOrNull(name) {
  return getSpreadsheet().getSheetByName(name);
}
