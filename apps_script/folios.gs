const FOLIO_CONFIG = {
  compras: { idCol: 'id_compra', prefix: 'C-' },
  ledger_credito: { idCol: 'id_movimiento', prefix: 'L-' }
};

function assignFolio(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const cfg = FOLIO_CONFIG[sheetName];
  if (!cfg) return;

  const row = e.range.getRow();
  if (row === 1) return; // header

  const idColIdx = getColumnIndex(sheetName, cfg.idCol);
  const idCell = sheet.getRange(row, idColIdx);
  if (idCell.getValue() !== '') return; // ya tiene folio

  // Buscar el maximo folio actual en la columna
  const lastRow = sheet.getLastRow();
  let maxN = 0;
  if (lastRow >= 2) {
    const values = sheet.getRange(2, idColIdx, lastRow - 1, 1).getValues();
    values.forEach(([v]) => {
      if (typeof v === 'string' && v.startsWith(cfg.prefix)) {
        const n = parseInt(v.substring(cfg.prefix.length), 10);
        if (!isNaN(n) && n > maxN) maxN = n;
      }
    });
  }
  idCell.setValue(`${cfg.prefix}${maxN + 1}`);
}

// Instalar trigger onEdit programaticamente (correr una sola vez)
function installFolioTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Borrar triggers previos para evitar duplicados
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'assignFolio') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('assignFolio')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  Logger.log('Trigger onEdit -> assignFolio instalado');
}
