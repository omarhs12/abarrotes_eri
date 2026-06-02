// ============================================================
// Mini test framework para Apps Script
// ============================================================

function assertEqual(actual, expected, message) {
  if (actual !== expected) {
    throw new Error(
      `${message || 'assertEqual'}\n  expected: ${JSON.stringify(expected)}\n  actual:   ${JSON.stringify(actual)}`
    );
  }
}

function assertDeepEqual(actual, expected, message) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a !== e) {
    throw new Error(
      `${message || 'assertDeepEqual'}\n  expected: ${e}\n  actual:   ${a}`
    );
  }
}

function assertTrue(value, message) {
  if (!value) {
    throw new Error(`${message || 'assertTrue'}: value was falsy`);
  }
}

function assertThrows(fn, message) {
  let threw = false;
  try { fn(); } catch (e) { threw = true; }
  if (!threw) {
    throw new Error(`${message || 'assertThrows'}: function did not throw`);
  }
}

// ============================================================
// Test runner — corre todas las funciones que empiezan con test_
// ============================================================

function runAllTests() {
  const results = { passed: 0, failed: 0, errors: [] };
  const globalObj = this;
  const testNames = Object.keys(globalObj).filter(name =>
    name.startsWith('test_') && typeof globalObj[name] === 'function'
  );

  testNames.forEach(name => {
    try {
      globalObj[name]();
      results.passed++;
      Logger.log(`PASS  ${name}`);
    } catch (e) {
      results.failed++;
      results.errors.push({ name: name, message: e.message });
      Logger.log(`FAIL  ${name}: ${e.message}`);
    }
  });

  Logger.log(`\n${results.passed} passed, ${results.failed} failed (de ${testNames.length} totales)`);
  if (results.failed > 0) {
    throw new Error(`${results.failed} test(s) fallaron`);
  }
  return results;
}

// ============================================================
// Smoke test del runner
// ============================================================

function test_runner_smoke() {
  assertEqual(1 + 1, 2, '1+1 deberia ser 2');
  assertTrue(true);
  assertDeepEqual([1, 2], [1, 2]);
  assertThrows(() => { throw new Error('intentional'); });
}

function test_createSheets_creates_all_12() {
  createSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SCHEMA.forEach(s => {
    const sheet = ss.getSheetByName(s.name);
    assertTrue(sheet !== null, `Hoja ${s.name} no se creo`);
    const headers = sheet.getRange(1, 1, 1, s.headers.length).getValues()[0];
    assertDeepEqual(headers, s.headers, `Headers de ${s.name}`);
  });
}

function test_validations_reject_invalid_enum() {
  createSheets();
  setupValidations();

  const productos = getSheet('productos');
  // La columna 'categoria' es enum: abarrotes, bebidas, limpieza, papel, otros
  // Intentar setear un valor invalido debe lanzar.
  const cell = productos.getRange(2, getColumnIndex('productos', 'categoria'));
  const rule = cell.getDataValidation();
  assertTrue(rule !== null, 'productos.categoria debe tener data validation');
  const criteria = rule.getCriteriaType();
  assertEqual(criteria, SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST, 'debe ser VALUE_IN_LIST');
  const values = rule.getCriteriaValues()[0];
  assertDeepEqual(values, ['abarrotes', 'bebidas', 'limpieza', 'papel', 'otros']);
}

function test_setupFormulas_saldo_actual() {
  createSheets();
  setupFormulas();

  const clientes = getSheet('clientes');
  const colSaldo = getColumnIndex('clientes', 'saldo_actual');
  const formula = clientes.getRange(2, colSaldo).getFormula();
  assertTrue(formula.length > 0, 'clientes.saldo_actual debe tener formula');
  assertTrue(formula.indexOf('ledger_credito') !== -1, 'la formula debe referenciar ledger_credito');
}

function test_assignFolio_compras() {
  createSheets();
  const compras = getSheet('compras');
  const colId = getColumnIndex('compras', 'id_compra');
  const colFecha = getColumnIndex('compras', 'fecha');

  // Limpiar filas de datos
  if (compras.getLastRow() > 1) {
    compras.getRange(2, 1, compras.getLastRow() - 1, compras.getLastColumn()).clearContent();
  }

  // Simular escritura en columna fecha de fila 2
  compras.getRange(2, colFecha).setValue('2026-06-01');
  const event = {
    range: compras.getRange(2, colFecha),
    source: SpreadsheetApp.getActiveSpreadsheet()
  };
  assignFolio(event);
  const id = compras.getRange(2, colId).getValue();
  assertEqual(id, 'C-1', 'primer folio debe ser C-1');

  // Segunda fila
  compras.getRange(3, colFecha).setValue('2026-06-02');
  event.range = compras.getRange(3, colFecha);
  assignFolio(event);
  assertEqual(compras.getRange(3, colId).getValue(), 'C-2', 'segundo folio debe ser C-2');
}

function test_assignFolio_idempotent() {
  createSheets();
  const compras = getSheet('compras');
  const colId = getColumnIndex('compras', 'id_compra');
  const colFecha = getColumnIndex('compras', 'fecha');

  if (compras.getLastRow() > 1) {
    compras.getRange(2, 1, compras.getLastRow() - 1, compras.getLastColumn()).clearContent();
  }

  compras.getRange(2, colFecha).setValue('2026-06-01');
  let event = { range: compras.getRange(2, colFecha), source: SpreadsheetApp.getActiveSpreadsheet() };
  assignFolio(event);
  assertEqual(compras.getRange(2, colId).getValue(), 'C-1');

  // Re-disparar onEdit en la misma fila no debe reasignar
  event.range = compras.getRange(2, colFecha);
  assignFolio(event);
  assertEqual(compras.getRange(2, colId).getValue(), 'C-1', 'folio no debe reasignarse');
}

function test_recalcularInventario_basic() {
  createSheets();
  const comprasDet = getSheet('compras_detalle');
  const ventasDet = getSheet('ventas_detalle');
  const inv = getSheet('inventario');

  // Limpiar
  [comprasDet, ventasDet, inv].forEach(sheet => {
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
  });

  // Cargar datos: compra de 10 unidades de SKU-A lote L1, caducidad 2026-12-31, costo 5.0, destino inventario
  comprasDet.appendRow(['C-1', 'SKU-A', 'L1', 10, 5.0, '2026-12-31', 'inventario']);
  // Compra de 5 unidades lote L2 caducidad 2026-08-15
  comprasDet.appendRow(['C-2', 'SKU-A', 'L2', 5, 5.0, '2026-08-15', 'inventario']);
  // Compra cross-dock no debe contar
  comprasDet.appendRow(['C-3', 'SKU-A', 'L3', 100, 5.0, '2027-01-01', 'cross_dock_CLI-1']);
  // Venta de 3 del lote L1
  ventasDet.appendRow(['V-1', 'SKU-A', 'L1', 3, 8.0, 0]);

  recalcularInventario();

  // Esperado: L1 = 10-3 = 7; L2 = 5; L3 NO debe aparecer
  const rows = inv.getRange(2, 1, inv.getLastRow() - 1, inv.getLastColumn()).getValues();
  const byLote = {};
  rows.forEach(r => { if (r[0]) byLote[r[1]] = r; });
  assertEqual(byLote['L1'][2], 7, 'L1 debe ser 7');
  assertEqual(byLote['L2'][2], 5, 'L2 debe ser 5');
  assertTrue(byLote['L3'] === undefined, 'L3 cross-dock NO debe estar en inventario');
}

function test_recalcularSaldos_llena_saldo_post() {
  createSheets();
  const ledger = getSheet('ledger_credito');

  // Limpiar
  if (ledger.getLastRow() > 1) {
    ledger.getRange(2, 1, ledger.getLastRow() - 1, ledger.getLastColumn()).clearContent();
  }

  // Cargar 3 movimientos del cliente CLI-1
  ledger.appendRow(['L-1', '2026-06-01', 'CLI-1', 'cargo',  'V-1', 1500, '']);
  ledger.appendRow(['L-2', '2026-06-02', 'CLI-1', 'abono',  'pago efectivo', 500, '']);
  ledger.appendRow(['L-3', '2026-06-03', 'CLI-1', 'cargo',  'V-2', 800, '']);

  recalcularSaldos();

  // saldo_post de L-1=1500, L-2=1000, L-3=1800
  const rows = ledger.getRange(2, 1, 3, ledger.getLastColumn()).getValues();
  const colSaldoPost = getColumnIndex('ledger_credito', 'saldo_post') - 1;
  assertEqual(rows[0][colSaldoPost], 1500, 'L-1 saldo_post');
  assertEqual(rows[1][colSaldoPost], 1000, 'L-2 saldo_post');
  assertEqual(rows[2][colSaldoPost], 1800, 'L-3 saldo_post');
}
