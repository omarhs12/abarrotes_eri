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

  // Dedup id_cliente
  const clientes = getSheet('clientes');
  if (clientes.getLastRow() >= 2) {
    const idColIdx = getColumnIndex('clientes', 'id_cliente');
    const existing = clientes.getRange(2, idColIdx, clientes.getLastRow() - 1, 1).getValues();
    const dup = existing.find(r => String(r[0]).trim() === idCliente);
    if (dup) return { ok: false, error: `ID '${idCliente}' ya existe en clientes` };
  }

  // Insert cliente row. saldo_actual queda como formula (no escribimos esa columna).
  // appendRow no respeta formulas existentes en la columna saldo_actual, asi que escribimos
  // columna por columna excepto saldo_actual. Pero appendRow escribe TODA la fila.
  // Solucion: appendRow con saldo_actual vacio, despues la formula se reaplica automaticamente
  // gracias a setupFormulas (que ya prellena las filas 2-5000). Como la fila nueva ya tiene
  // formula preparada, appendRow va a pisarla con vacio. Hay que volver a setear la formula
  // despues. Solucion mas simple: NO escribir esa columna usando setValues en lugar de appendRow.

  const headers = getSchemaFor('clientes').headers;
  const colSaldo = getColumnIndex('clientes', 'saldo_actual');
  const targetRow = clientes.getLastRow() + 1;

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
  if (clientes.getLastRow() < 2) return 0;
  const activoCol = getColumnIndex('clientes', 'activo');
  const rows = clientes.getRange(2, activoCol, clientes.getLastRow() - 1, 1).getValues();
  return rows.filter(r => String(r[0]).toUpperCase() === 'SI').length;
}
