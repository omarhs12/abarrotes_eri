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

// Devuelve el primer renglon vacio (>= 2) en una columna especifica.
// Necesario en hojas con formulas prellenadas (ej. clientes.saldo_actual)
// donde getLastRow() devuelve el ultimo renglon con formula, no con datos.
// Busca la primera celda vacia en la columna indicada.
function nextEmptyRow(sheet, columnName) {
  const colIdx = getColumnIndex(sheet.getName(), columnName);
  const maxRow = sheet.getMaxRows();
  if (maxRow < 2) return 2;
  const values = sheet.getRange(2, colIdx, maxRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const v = values[i][0];
    if (v === '' || v === null || v === undefined) return i + 2;
  }
  return maxRow + 1;
}
