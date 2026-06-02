function createSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  SCHEMA.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
    }

    // Reset row 1 con headers definitivos
    sheet.getRange(1, 1, 1, s.headers.length)
      .setValues([s.headers])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, s.headers.length);
  });

  // Borrar hoja default "Hoja 1" / "Sheet1" si quedo vacia
  const defaults = ['Hoja 1', 'Sheet1', 'Hoja1'];
  defaults.forEach(name => {
    const s = ss.getSheetByName(name);
    if (s && s.getLastRow() === 0) ss.deleteSheet(s);
  });

  Logger.log(`createSheets: ${SCHEMA.length} hojas listas`);
}
