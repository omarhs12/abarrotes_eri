function recalcularSaldos() {
  const ledger = getSheet('ledger_credito');
  if (ledger.getLastRow() < 2) return;

  const rows = ledger.getRange(2, 1, ledger.getLastRow() - 1, ledger.getLastColumn()).getValues();
  const colFecha = getColumnIndex('ledger_credito', 'fecha') - 1;
  const colCliente = getColumnIndex('ledger_credito', 'cliente_id') - 1;
  const colTipo = getColumnIndex('ledger_credito', 'tipo') - 1;
  const colMonto = getColumnIndex('ledger_credito', 'monto') - 1;
  const colSaldoPost = getColumnIndex('ledger_credito', 'saldo_post') - 1;

  // Ordenar por (cliente, fecha) manteniendo el indice original para escribir de vuelta
  const indexed = rows.map((r, i) => ({ row: r, originalIdx: i }));
  indexed.sort((a, b) => {
    const ca = a.row[colCliente], cb = b.row[colCliente];
    if (ca < cb) return -1;
    if (ca > cb) return 1;
    const da = new Date(a.row[colFecha]).getTime();
    const db = new Date(b.row[colFecha]).getTime();
    return da - db;
  });

  // Recorrer en orden y acumular saldo por cliente
  const saldos = {};
  indexed.forEach(item => {
    const cliente = item.row[colCliente];
    const tipo = item.row[colTipo];
    const monto = Number(item.row[colMonto]) || 0;
    if (saldos[cliente] === undefined) saldos[cliente] = 0;
    if (tipo === 'cargo') saldos[cliente] += monto;
    else if (tipo === 'abono') saldos[cliente] -= monto;
    item.row[colSaldoPost] = saldos[cliente];
  });

  // Reescribir respetando el orden original
  const out = new Array(rows.length);
  indexed.forEach(item => { out[item.originalIdx] = item.row; });
  ledger.getRange(2, 1, out.length, out[0].length).setValues(out);

  Logger.log(`recalcularSaldos: ${out.length} movimientos actualizados`);
}
