const MAX_ROWS_FOR_FORMULAS = 5000;

function setupFormulas() {
  setupClientesSaldoFormula();
  Logger.log('setupFormulas: formulas aplicadas');
}

function setupClientesSaldoFormula() {
  const clientes = getSheet('clientes');
  const colSaldo = getColumnIndex('clientes', 'saldo_actual');
  const colIdCliente = getColumnIndex('clientes', 'id_cliente');
  const idColLetter = columnToLetter(colIdCliente);

  // Construir TODAS las formulas en memoria y escribirlas en UNA llamada batch
  // (setFormulas en un range es ~100x mas rapido que setFormula fila por fila).
  //
  // OJO: las letras C/D/F abajo son indices DUROS al esquema actual de ledger_credito:
  //   C = cliente_id, D = tipo, F = monto.
  // Si en el futuro se reordenan o insertan columnas en ledger_credito, esta formula
  // lee columnas equivocadas SIN error. Antes de tocar ese SCHEMA, actualizar aqui.
  const formulas = [];
  for (let row = 2; row <= MAX_ROWS_FOR_FORMULAS; row++) {
    const idRef = `${idColLetter}${row}`;
    formulas.push([
      `=IF(${idRef}="","",`
      + `SUMIFS(ledger_credito!F:F, ledger_credito!C:C, ${idRef}, ledger_credito!D:D, "cargo")`
      + ` - SUMIFS(ledger_credito!F:F, ledger_credito!C:C, ${idRef}, ledger_credito!D:D, "abono"))`
    ]);
  }
  clientes.getRange(2, colSaldo, MAX_ROWS_FOR_FORMULAS - 1, 1).setFormulas(formulas);
}

function columnToLetter(col) {
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - mod) / 26);
  }
  return letter;
}
