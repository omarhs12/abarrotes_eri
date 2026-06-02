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
