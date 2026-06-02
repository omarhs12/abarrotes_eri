// Utilidades para ordenar hojas alfabeticamente sin afectar el encabezado.

function ordenarHoja(sheetName, columnName, ascending) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return; // 0 o 1 fila de datos = nada que ordenar

  const colIdx = getColumnIndex(sheetName, columnName);
  const lastCol = sheet.getLastColumn();

  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  range.sort({ column: colIdx, ascending: ascending !== false });
}

function ordenarProductosPorNombre() {
  ordenarHoja('productos', 'nombre', true);
  SpreadsheetApp.getActiveSpreadsheet().toast('Productos ordenados A-Z por nombre', 'Listo', 3);
}

function ordenarClientesPorNombre() {
  ordenarHoja('clientes', 'nombre', true);
  SpreadsheetApp.getActiveSpreadsheet().toast('Clientes ordenados A-Z por nombre', 'Listo', 3);
}

function ordenarProveedoresPorNombre() {
  ordenarHoja('proveedores', 'nombre', true);
  SpreadsheetApp.getActiveSpreadsheet().toast('Proveedores ordenados A-Z por nombre', 'Listo', 3);
}
