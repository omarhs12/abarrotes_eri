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

  // Para cada fila n de clientes:
  // saldo_actual = SUMIFS(ledger.monto WHERE id_cliente=ID AND tipo=cargo)
  //              - SUMIFS(ledger.monto WHERE id_cliente=ID AND tipo=abono)
  for (let row = 2; row <= MAX_ROWS_FOR_FORMULAS; row++) {
    const idRef = `${idColLetter}${row}`;
    const formula = `=IF(${idRef}="","",`
      + `SUMIFS(ledger_credito!F:F, ledger_credito!C:C, ${idRef}, ledger_credito!D:D, "cargo")`
      + ` - SUMIFS(ledger_credito!F:F, ledger_credito!C:C, ${idRef}, ledger_credito!D:D, "abono"))`;
    clientes.getRange(row, colSaldo).setFormula(formula);
  }
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
