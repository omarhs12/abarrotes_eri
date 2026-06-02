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
