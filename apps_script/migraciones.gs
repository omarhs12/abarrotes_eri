// ============================================================
// MIGRACIONES Y RECOVERIES — codigo de un solo uso
// ============================================================
//
// Estas funciones existen para corregir estado del Sheet despues de un
// cambio de modelo o de un bug ya solucionado. Una vez ejecutadas y
// verificadas por Omar, BORRARLAS de este archivo (y del menu).
//
// Convencion: cada funcion lleva comentario con fecha de creacion y
// fecha estimada para borrar.
// ============================================================


// Creada 2026-06-02 — borrar tras confirmar que ya no hay clientes huerfanos.
// Recuperacion del bug donde guardarClienteDesdeFormulario escribia en fila
// 5001+ por mal calculo de getLastRow vs formulas prellenadas en saldo_actual.
// El bug ya esta arreglado en forms_cliente.gs; esta funcion solo sirve si
// quedaron clientes "perdidos" abajo. Una corrida ya se hizo el 2026-06-02.
function recuperarClientesOrfanos() {
  const clientes = getSheet('clientes');
  const idColIdx = getColumnIndex('clientes', 'id_cliente');
  const colSaldo = getColumnIndex('clientes', 'saldo_actual');
  const lastCol = clientes.getLastColumn();
  const maxRow = clientes.getMaxRows();

  const idValues = clientes.getRange(2, idColIdx, maxRow - 1, 1).getValues();
  const sourceRows = [];
  idValues.forEach((row, i) => {
    if (row[0] !== '' && row[0] !== null) sourceRows.push(i + 2);
  });

  if (sourceRows.length === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No hay clientes que recuperar', 'Recovery', 3);
    return;
  }

  const dataLeida = sourceRows.map(rowNum => {
    return clientes.getRange(rowNum, 1, 1, lastCol).getValues()[0];
  });

  sourceRows.forEach(rowNum => {
    for (let c = 1; c <= lastCol; c++) {
      if (c === colSaldo) continue;
      clientes.getRange(rowNum, c).setValue('');
    }
  });

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


// Creada 2026-06-03 — borrar tras confirmar que se ejecuto con exito.
// Copia el `nombre` al `sku` en todos los productos existentes.
// Idempotente: solo actualiza filas donde sku != nombre.
// Nuevos productos creados via el formulario ya nacen con sku=nombre.
function migrarSkuAlNombre() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Migrar SKU = Nombre',
    'Esta accion copia la columna "nombre" a la columna "sku" en todos los productos. ' +
    'Si tienes compras o ventas registradas que referencian los SKU actuales, esas referencias quedaran rotas (avisa primero).\n\n' +
    'Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const productos = getSheet('productos');
  if (productos.getLastRow() < 2) {
    ui.alert('Sin productos para migrar.');
    return;
  }

  const colSku = getColumnIndex('productos', 'sku');
  const colNombre = getColumnIndex('productos', 'nombre');
  const rows = productos.getRange(2, 1, productos.getLastRow() - 1, productos.getLastColumn()).getValues();

  let cambios = 0;
  rows.forEach((row, i) => {
    const sku = String(row[colSku - 1] || '').trim();
    const nombre = String(row[colNombre - 1] || '').trim();
    if (nombre && sku !== nombre) {
      productos.getRange(i + 2, colSku).setValue(nombre);
      cambios++;
    }
  });

  ui.alert(`${cambios} producto(s) migrado(s). El SKU ahora es igual al nombre.`);
}
