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

function test_generarListaPreciosTexto_basico() {
  createSheets();
  const productos = getSheet('productos');
  if (productos.getLastRow() > 1) {
    productos.getRange(2, 1, productos.getLastRow() - 1, productos.getLastColumn()).clearContent();
  }
  productos.appendRow(['FRIJ-1', 'Frijol bayo 1kg', 'abarrotes', 'kg', 1, 35, 0, 'SI']);
  productos.appendRow(['COCA-600', 'Coca cola 600ml', 'bebidas', 'pza', 0.7, 18, 0, 'SI']);
  productos.appendRow(['INACTIVO', 'No debe aparecer', 'abarrotes', 'pza', 0, 99, 0, 'NO']);

  const texto = generarListaPreciosTexto();
  assertTrue(texto.indexOf('Frijol bayo 1kg') !== -1, 'producto activo debe estar en la lista');
  assertTrue(texto.indexOf('Coca cola 600ml') !== -1, 'segundo producto activo debe estar');
  assertTrue(texto.indexOf('No debe aparecer') === -1, 'producto inactivo NO debe aparecer');
  assertTrue(texto.indexOf('ABARROTES') !== -1, 'categoria abarrotes debe aparecer');
  assertTrue(texto.indexOf('BEBIDAS') !== -1, 'categoria bebidas debe aparecer');
  assertTrue(texto.indexOf('35.00') !== -1, 'precio formateado debe aparecer');
}

function test_alertasCaducidad_filtra_por_dias() {
  createSheets();
  const inv = getSheet('inventario');
  if (inv.getLastRow() > 1) {
    inv.getRange(2, 1, inv.getLastRow() - 1, inv.getLastColumn()).clearContent();
  }

  // hoy = 2026-06-01 segun el contexto del proyecto. Calculamos relativo a Date.now().
  const today = new Date();
  const in3days = new Date(today.getTime() + 3 * 24 * 3600 * 1000);
  const in30days = new Date(today.getTime() + 30 * 24 * 3600 * 1000);
  const in100days = new Date(today.getTime() + 100 * 24 * 3600 * 1000);

  inv.appendRow(['SKU-X', 'LX1', 5, in3days, 10, '']);
  inv.appendRow(['SKU-X', 'LX2', 5, in30days, 10, '']);
  inv.appendRow(['SKU-X', 'LX3', 5, in100days, 10, '']);

  const r7 = alertasCaducidad(7);
  assertEqual(r7.length, 1, 'umbral 7d: solo LX1');
  assertEqual(r7[0].lote, 'LX1');

  const r60 = alertasCaducidad(60);
  assertEqual(r60.length, 2, 'umbral 60d: LX1 y LX2');
}

function test_config_defaults_set() {
  createSheets();
  setConfigDefaults();

  const config = getSheet('config');
  const rows = config.getRange(2, 1, config.getLastRow() - 1, 3).getValues();
  const map = {};
  rows.forEach(r => { if (r[0]) map[r[0]] = r[1]; });

  assertTrue(map['capacidad_pickup_kg'] !== undefined, 'capacidad_pickup_kg debe existir');
  assertTrue(map['umbral_alerta_caducidad_dias'] !== undefined, 'umbral_alerta_caducidad_dias debe existir');
  assertEqual(map['moneda'], 'MXN');
}

function test_carga_inicial_crea_compra_y_detalle() {
  createSheets();

  // Limpiar compras y compras_detalle de pruebas previas
  const compras = getSheet('compras');
  const comprasDet = getSheet('compras_detalle');
  [compras, comprasDet].forEach(sheet => {
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
  });

  // Necesita un producto registrado primero
  const productos = getSheet('productos');
  if (productos.getLastRow() > 1) {
    productos.getRange(2, 1, productos.getLastRow() - 1, productos.getLastColumn()).clearContent();
  }
  productos.appendRow(['SKU-CI', 'Producto Carga Inicial', 'abarrotes', 'pza', 0.5, 20, 0, 'SI']);

  // Guardar dos lineas via la API del form
  const r1 = guardarLineaCargaInicial({ sku: 'SKU-CI', lote: 'L1', cantidad: 10, costo_unitario: 5 });
  assertTrue(r1.ok, 'primera linea debe guardarse');
  assertTrue(r1.idCompra.indexOf('INICIAL-') === 0, 'idCompra debe arrancar con INICIAL-');

  const r2 = guardarLineaCargaInicial({ sku: 'SKU-CI', lote: 'L2', cantidad: 5, costo_unitario: 5 });
  assertTrue(r2.ok, 'segunda linea debe guardarse');
  assertEqual(r2.idCompra, r1.idCompra, 'ambas lineas usan el mismo folio INICIAL del dia');

  // Verificar que en compras hay 1 sola fila INICIAL
  const cRows = compras.getRange(2, 1, compras.getLastRow() - 1, compras.getLastColumn()).getValues();
  const inicialRows = cRows.filter(r => typeof r[0] === 'string' && r[0].indexOf('INICIAL-') === 0);
  assertEqual(inicialRows.length, 1, 'solo debe haber 1 fila INICIAL en compras');

  // Verificar 2 lineas en compras_detalle
  const dRows = comprasDet.getRange(2, 1, comprasDet.getLastRow() - 1, comprasDet.getLastColumn()).getValues();
  const detalleRows = dRows.filter(r => r[0] === r1.idCompra);
  assertEqual(detalleRows.length, 2, 'deben haber 2 lineas en compras_detalle');

  // SKU inexistente debe fallar
  const r3 = guardarLineaCargaInicial({ sku: 'NO-EXISTE', lote: 'L3', cantidad: 1 });
  assertTrue(!r3.ok, 'SKU inexistente debe rechazarse');
}

function test_ordenarProductosPorNombre() {
  createSheets();
  const productos = getSheet('productos');
  if (productos.getLastRow() > 1) {
    productos.getRange(2, 1, productos.getLastRow() - 1, productos.getLastColumn()).clearContent();
  }
  productos.appendRow(['Z01', 'Zanahoria', 'abarrotes', 'kg', 0.5, 20, 0, 'SI']);
  productos.appendRow(['A01', 'Aceite', 'abarrotes', 'lt', 1, 50, 0, 'SI']);
  productos.appendRow(['M01', 'Manzana', 'abarrotes', 'kg', 0.3, 30, 0, 'SI']);

  ordenarProductosPorNombre();

  const rows = productos.getRange(2, 1, 3, productos.getLastColumn()).getValues();
  assertEqual(rows[0][1], 'Aceite', 'primero debe ser Aceite');
  assertEqual(rows[1][1], 'Manzana', 'segundo debe ser Manzana');
  assertEqual(rows[2][1], 'Zanahoria', 'tercero debe ser Zanahoria');
  // El SKU debe acompañar a su fila correspondiente (no quedarse fijo)
  assertEqual(rows[0][0], 'A01', 'SKU debe seguir al producto al ordenar');
  assertEqual(rows[1][0], 'M01');
  assertEqual(rows[2][0], 'Z01');
}

function test_guardarCliente_contado_sin_saldo() {
  createSheets();
  const clientes = getSheet('clientes');
  const ledger = getSheet('ledger_credito');
  [clientes, ledger].forEach(s => {
    if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
  });

  const r = guardarClienteDesdeFormulario({
    id_cliente: 'C100', nombre: 'Test Contado', contacto: '555-1234',
    tipo: 'contado', limite_credito: 0, saldo_pendiente: 0, activo: 'SI'
  });
  assertTrue(r.ok, 'cliente contado se debe guardar');
  assertTrue(r.saldoCreado === null, 'no debe crear movimiento en ledger');

  // Verificar fila en clientes
  const row = clientes.getRange(2, 1, 1, clientes.getLastColumn()).getValues()[0];
  assertEqual(row[0], 'C100');
  assertEqual(row[1], 'Test Contado');
  assertEqual(row[3], 'contado');
}

function test_guardarCliente_credito_con_saldo_inicial() {
  createSheets();
  const clientes = getSheet('clientes');
  const ledger = getSheet('ledger_credito');
  [clientes, ledger].forEach(s => {
    if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
  });

  const r = guardarClienteDesdeFormulario({
    id_cliente: 'C200', nombre: 'Test Credito', contacto: '555-5678',
    tipo: 'credito', limite_credito: 5000, saldo_pendiente: 1500, activo: 'SI'
  });
  assertTrue(r.ok, 'cliente credito debe guardarse');
  assertTrue(r.saldoCreado !== null, 'debe haber creado movimiento en ledger');
  assertEqual(r.saldoCreado.monto, 1500);

  // Verificar fila en ledger
  const ledgerRow = ledger.getRange(2, 1, 1, ledger.getLastColumn()).getValues()[0];
  assertTrue(ledgerRow[0].indexOf('L-') === 0, 'folio L-N asignado');
  assertEqual(ledgerRow[2], 'C200', 'cliente_id correcto');
  assertEqual(ledgerRow[3], 'cargo', 'tipo cargo');
  assertEqual(ledgerRow[4], 'saldo inicial', 'referencia saldo inicial');
  assertEqual(ledgerRow[5], 1500, 'monto');
}

function test_guardarCliente_dedup() {
  createSheets();
  const clientes = getSheet('clientes');
  if (clientes.getLastRow() > 1) clientes.getRange(2, 1, clientes.getLastRow() - 1, clientes.getLastColumn()).clearContent();

  const r1 = guardarClienteDesdeFormulario({
    id_cliente: 'C300', nombre: 'Primero', tipo: 'contado', limite_credito: 0, saldo_pendiente: 0, activo: 'SI'
  });
  assertTrue(r1.ok);

  const r2 = guardarClienteDesdeFormulario({
    id_cliente: 'C300', nombre: 'Duplicado', tipo: 'contado', limite_credito: 0, saldo_pendiente: 0, activo: 'SI'
  });
  assertTrue(!r2.ok, 'segundo intento debe rechazarse por id duplicado');
}

function test_guardarCliente_rechaza_saldo_sin_credito() {
  createSheets();
  const clientes = getSheet('clientes');
  if (clientes.getLastRow() > 1) clientes.getRange(2, 1, clientes.getLastRow() - 1, clientes.getLastColumn()).clearContent();

  const r = guardarClienteDesdeFormulario({
    id_cliente: 'C400', nombre: 'Contado con saldo', tipo: 'contado',
    limite_credito: 0, saldo_pendiente: 500, activo: 'SI'
  });
  assertTrue(!r.ok, 'cliente contado con saldo pendiente debe rechazarse');
}

// Reproduce el bug donde setupFormulas() prellena formulas hasta fila 5000,
// haciendo que getLastRow() devuelva 5000 y guardarCliente escriba fuera del area visible.
// Con el fix, el cliente debe quedar en la fila 2.
function test_guardarCliente_respeta_formulas_prellenadas() {
  createSheets();
  setupFormulas(); // <-- crucial: simula el estado real de produccion
  const clientes = getSheet('clientes');
  // Limpiar cualquier id_cliente en la columna (formulas en saldo_actual se quedan)
  const idColIdx = getColumnIndex('clientes', 'id_cliente');
  const maxRow = clientes.getMaxRows();
  clientes.getRange(2, idColIdx, maxRow - 1, 1).clearContent();

  const r = guardarClienteDesdeFormulario({
    id_cliente: 'JAQUI', nombre: 'Jaqui', contacto: '614-555-0001',
    tipo: 'credito', limite_credito: 3000, saldo_pendiente: 0, activo: 'SI'
  });
  assertTrue(r.ok, 'cliente debe guardarse aun con formulas prellenadas');

  // El cliente debe estar en la fila 2 (no en la 5001)
  const jaquiEnFila2 = clientes.getRange(2, idColIdx).getValue();
  assertEqual(jaquiEnFila2, 'JAQUI', 'cliente debe quedar en fila 2, no fila 5001');
}

function test_recuperarClientesOrfanos_mueve_filas_abajo_arriba() {
  createSheets();
  const clientes = getSheet('clientes');
  const idColIdx = getColumnIndex('clientes', 'id_cliente');
  const nombreColIdx = getColumnIndex('clientes', 'nombre');
  const maxRow = clientes.getMaxRows();
  clientes.getRange(2, 1, maxRow - 1, clientes.getLastColumn()).clearContent();

  // Simular bug: cliente escrito en fila 5001 (huerfano)
  clientes.getRange(5001, idColIdx).setValue('JAQUI');
  clientes.getRange(5001, nombreColIdx).setValue('Jaqui');

  recuperarClientesOrfanos();

  assertEqual(clientes.getRange(2, idColIdx).getValue(), 'JAQUI', 'Jaqui debe estar en fila 2');
  assertEqual(clientes.getRange(2, nombreColIdx).getValue(), 'Jaqui');
  assertEqual(clientes.getRange(5001, idColIdx).getValue(), '', 'fila 5001 debe quedar vacia');
}
