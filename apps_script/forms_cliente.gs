// Servidor del formulario "Nuevo cliente"

function abrirFormularioCliente() {
  const html = HtmlService.createHtmlOutputFromFile('form_cliente')
    .setWidth(440)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Nuevo cliente');
}

function guardarClienteDesdeFormulario(data) {
  if (!data.id_cliente) return { ok: false, error: 'ID cliente es obligatorio' };
  if (!data.nombre) return { ok: false, error: 'Nombre es obligatorio' };

  const idCliente = String(data.id_cliente).trim();
  const nombre = String(data.nombre).trim();
  const contacto = String(data.contacto || '').trim();
  const tipo = data.tipo === 'credito' ? 'credito' : 'contado';
  const limiteCredito = Number(data.limite_credito) || 0;
  const activo = data.activo === 'NO' ? 'NO' : 'SI';
  const saldoPendiente = Number(data.saldo_pendiente) || 0;

  // Validaciones cruzadas
  if (tipo === 'credito' && limiteCredito < 0) {
    return { ok: false, error: 'Limite de credito no puede ser negativo' };
  }
  if (saldoPendiente < 0) {
    return { ok: false, error: 'Saldo pendiente no puede ser negativo' };
  }
  if (saldoPendiente > 0 && tipo !== 'credito') {
    return { ok: false, error: 'Solo clientes con tipo "credito" pueden tener saldo pendiente' };
  }

  // Dedup id_cliente — escanear TODA la columna porque el rango de datos puede
  // mezclarse con celdas que solo tienen formula prellenada en saldo_actual.
  const clientes = getSheet('clientes');
  const idColIdx = getColumnIndex('clientes', 'id_cliente');
  const maxRow = clientes.getMaxRows();
  if (maxRow >= 2) {
    const existing = clientes.getRange(2, idColIdx, maxRow - 1, 1).getValues();
    const dup = existing.find(r => String(r[0]).trim() === idCliente);
    if (dup) return { ok: false, error: `ID '${idCliente}' ya existe en clientes` };
  }

  // Como saldo_actual tiene formula prellenada (rows 2-5000), getLastRow()
  // no sirve para encontrar la siguiente fila libre. Usamos nextEmptyRow
  // que busca la primera celda vacia en la columna id_cliente.
  const headers = getSchemaFor('clientes').headers;
  const colSaldo = getColumnIndex('clientes', 'saldo_actual');
  const targetRow = nextEmptyRow(clientes, 'id_cliente');

  // Construir array completo de la fila (sin tocar la columna de saldo_actual)
  const rowValues = [
    idCliente,
    nombre,
    contacto,
    tipo,
    limiteCredito,
    null, // placeholder para saldo_actual — no lo escribimos
    activo
  ];

  // Escribir columnas individualmente, saltando saldo_actual
  for (let i = 0; i < headers.length; i++) {
    if ((i + 1) === colSaldo) continue; // skip saldo_actual
    clientes.getRange(targetRow, i + 1).setValue(rowValues[i]);
  }

  // Si tiene saldo pendiente, crear cargo en ledger_credito
  let movimientoFolio = null;
  if (saldoPendiente > 0 && tipo === 'credito') {
    movimientoFolio = asignarFolioLedger();
    const ledger = getSheet('ledger_credito');
    ledger.appendRow([movimientoFolio, new Date(), idCliente, 'cargo', 'saldo inicial', saldoPendiente, '']);
  }

  return {
    ok: true,
    id_cliente: idCliente,
    nombre: nombre,
    saldoCreado: movimientoFolio ? { folio: movimientoFolio, monto: saldoPendiente } : null
  };
}

// Asigna el siguiente folio L-N en ledger_credito sin depender del trigger onEdit
// (necesario porque appendRow programatico NO dispara onEdit).
function asignarFolioLedger() {
  const ledger = getSheet('ledger_credito');
  const idColIdx = getColumnIndex('ledger_credito', 'id_movimiento');
  let maxN = 0;
  if (ledger.getLastRow() >= 2) {
    const values = ledger.getRange(2, idColIdx, ledger.getLastRow() - 1, 1).getValues();
    values.forEach(([v]) => {
      if (typeof v === 'string' && v.indexOf('L-') === 0) {
        const n = parseInt(v.substring(2), 10);
        if (!isNaN(n) && n > maxN) maxN = n;
      }
    });
  }
  return `L-${maxN + 1}`;
}

function contarClientesActivos() {
  const clientes = getSheet('clientes');
  const idColIdx = getColumnIndex('clientes', 'id_cliente');
  const activoCol = getColumnIndex('clientes', 'activo');
  const maxRow = clientes.getMaxRows();
  if (maxRow < 2) return 0;
  const ids = clientes.getRange(2, idColIdx, maxRow - 1, 1).getValues();
  const activos = clientes.getRange(2, activoCol, maxRow - 1, 1).getValues();
  let count = 0;
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] !== '' && ids[i][0] !== null && String(activos[i][0]).toUpperCase() === 'SI') {
      count++;
    }
  }
  return count;
}

// Recovery: si la version anterior del formulario escribio clientes en filas muy abajo
// (por el bug de getLastRow + formulas prellenadas), los mueve de regreso al inicio.
function recuperarClientesOrfanos() {
  const clientes = getSheet('clientes');
  const idColIdx = getColumnIndex('clientes', 'id_cliente');
  const colSaldo = getColumnIndex('clientes', 'saldo_actual');
  const lastCol = clientes.getLastColumn();
  const maxRow = clientes.getMaxRows();

  // 1. Encontrar todos los renglones con id_cliente llenado (en cualquier fila)
  const idValues = clientes.getRange(2, idColIdx, maxRow - 1, 1).getValues();
  const sourceRows = [];
  idValues.forEach((row, i) => {
    if (row[0] !== '' && row[0] !== null) sourceRows.push(i + 2);
  });

  if (sourceRows.length === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No hay clientes que recuperar', 'Recovery', 3);
    return;
  }

  // 2. Leer la data de cada renglon (toda la fila excepto saldo_actual)
  const dataLeida = sourceRows.map(rowNum => {
    return clientes.getRange(rowNum, 1, 1, lastCol).getValues()[0];
  });

  // 3. Limpiar los renglones origen, columna por columna excepto saldo_actual
  sourceRows.forEach(rowNum => {
    for (let c = 1; c <= lastCol; c++) {
      if (c === colSaldo) continue;
      clientes.getRange(rowNum, c).setValue('');
    }
  });

  // 4. Reescribir los datos al inicio (filas 2, 3, 4...) saltando saldo_actual
  dataLeida.forEach((rowData, idx) => {
    const targetRow = idx + 2;
    for (let c = 0; c < lastCol; c++) {
      if (c + 1 === colSaldo) continue;
      clientes.getRange(targetRow, c + 1).setValue(rowData[c]);
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `${sourceRows.length} cliente(s) recuperado(s) y compactado(s) al inicio`,
    'Recovery',
    5
  );
}
