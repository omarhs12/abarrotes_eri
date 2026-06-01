# Fase 1 — Sheets Backoffice Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Construir el backoffice de Abarrotes Eri en Google Sheets: 12 hojas con encabezados, validaciones, fórmulas, y Apps Script (folios, recálculo de inventario FEFO, recálculo de saldos). Al terminar, Omar puede capturar compras/inventario/clientes/ventas manualmente y los cálculos se mantienen consistentes.

**Architecture:** Repositorio local con código de Apps Script versionado vía `clasp` (Google's CLI). El Sheet vive en Drive de Omar; el `.clasp.json` referencia el `scriptId` del proyecto Apps Script atado al Sheet. Todo el código en `apps_script/` se sincroniza con `clasp push`. Pruebas como funciones `test_*` ejecutables desde el editor de Apps Script o `clasp run`.

**Tech Stack:** Google Sheets, Google Apps Script (V8 runtime), `@google/clasp` (Node CLI), Git.

---

## File Structure

```
abarrotes_eri/
├── README.md                              # operación y onboarding
├── .gitignore
├── docs/superpowers/
│   ├── specs/2026-06-01-abarrotes-eri-design.md     # ya existe
│   └── plans/2026-06-01-fase-1-sheets-backoffice.md # este plan
├── apps_script/
│   ├── .clasp.json                        # referencia al scriptId (gitignored si tiene IDs sensibles)
│   ├── .claspignore
│   ├── appsscript.json                    # manifest
│   ├── schema.gs                          # definicion de las 12 hojas (single source)
│   ├── setup_sheets.gs                    # createSheets() — crea hojas y headers
│   ├── validations.gs                     # setupValidations()
│   ├── formulas.gs                        # setupFormulas()
│   ├── folios.gs                          # assignFolio() onEdit
│   ├── inventario.gs                      # recalcularInventario()
│   ├── saldos.gs                          # recalcularSaldos()
│   ├── alertas.gs                         # alertasCaducidad()
│   ├── menu.gs                            # onOpen() menu personalizado
│   ├── tests.gs                           # assert helpers + test_* functions
│   └── lib.gs                             # helpers compartidos (getSheet, etc.)
└── config/
    └── sheet_id.md                        # documentacion del Sheet ID y URL
```

---

## Task 1: Repo skeleton (README, .gitignore, estructura)

**Files:**
- Create: `README.md`
- Create: `.gitignore`
- Create: `apps_script/` (directorio)
- Create: `config/` (directorio)

- [ ] **Step 1: Crear `.gitignore`**

```gitignore
# Apps Script credentials
.clasprc.json
apps_script/.clasp.json

# Node
node_modules/
package-lock.json
npm-debug.log

# OS
.DS_Store
Thumbs.db

# Editor
.vscode/
.idea/

# Local secrets
.env
.env.local
```

- [ ] **Step 2: Crear `README.md` inicial**

```markdown
# Abarrotes Eri

Sistema operativo (mini-ERP) para reventa de abarrotes en Chihuahua.

## Componentes
- **Google Sheets**: fuente de verdad de datos (productos, compras, inventario, clientes, ventas, credito, MSI).
- **Apps Script** (en `apps_script/`): folios automaticos, recalculo de inventario FEFO, saldos.
- **PWA Android** (Fase 2, pendiente): notas de venta con impresion termica Bluetooth.

## Spec
Ver [docs/superpowers/specs/2026-06-01-abarrotes-eri-design.md](docs/superpowers/specs/2026-06-01-abarrotes-eri-design.md).

## Setup
Ver seccion **Onboarding** mas abajo (se completa al final de Fase 1).

## Onboarding
TBD al cierre de Fase 1.
```

- [ ] **Step 3: Crear directorios vacios**

```bash
mkdir -p apps_script config
touch apps_script/.gitkeep config/.gitkeep
```

- [ ] **Step 4: Commit**

```bash
cd /c/proyectos/abarrotes_eri
git add .gitignore README.md apps_script/.gitkeep config/.gitkeep
git -c user.name="Omar" -c user.email="omarhs12@gmail.com" commit -m "chore: repo skeleton (README, gitignore, estructura de carpetas)"
```

---

## Task 2: Crear el Google Sheet y documentar el ID

**Files:**
- Create: `config/sheet_id.md`

Esta tarea es **manual** (creacion de un recurso en Google Drive), pero queda documentada.

- [ ] **Step 1: Crear el Sheet en Drive**

1. Abrir https://sheets.new (crea Sheet vacio)
2. Renombrarlo a `Abarrotes Eri - Operativo`
3. Copiar la URL completa de la barra de direcciones
4. El ID es el segmento entre `/d/` y `/edit` (un string largo alfanumerico)

- [ ] **Step 2: Documentar el ID en `config/sheet_id.md`**

```markdown
# Google Sheet — Abarrotes Eri Operativo

**Nombre:** Abarrotes Eri - Operativo
**Sheet ID:** `<PEGAR_AQUI_EL_ID>`
**URL:** https://docs.google.com/spreadsheets/d/<PEGAR_AQUI_EL_ID>/edit
**Dueno:** omarhs12@gmail.com

## Hojas (creadas via Apps Script `createSheets()`)
- productos
- proveedores
- compras
- compras_detalle
- inventario
- clientes
- ventas
- ventas_detalle
- ledger_credito
- msi
- tarjetas
- config

## Notas
- Compartir como editor a quien capture compras nocturnas.
- No borrar manualmente hojas; usar `createSheets()` para reinicializar si necesario.
```

- [ ] **Step 3: Commit**

```bash
git add config/sheet_id.md
git -c user.name="Omar" -c user.email="omarhs12@gmail.com" commit -m "docs: documentar Sheet ID y URL del workbook operativo"
```

---

## Task 3: Setup de clasp + Apps Script project bound

**Files:**
- Create: `apps_script/appsscript.json`
- Create: `apps_script/.claspignore`
- Create: `apps_script/.clasp.json` (no commiteado — esta en `.gitignore`)

- [ ] **Step 1: Instalar clasp global**

```bash
npm install -g @google/clasp
clasp --version
```
Expected: version output (>= 2.4.x).

- [ ] **Step 2: Login**

```bash
clasp login
```
Expected: abre navegador, autoriza con cuenta `omarhs12@gmail.com`, mensaje "You are logged in".

- [ ] **Step 3: Crear proyecto Apps Script atado al Sheet**

```bash
cd /c/proyectos/abarrotes_eri/apps_script
clasp create --type sheets --title "Abarrotes Eri - Apps Script" --parentId <SHEET_ID>
```
Reemplazar `<SHEET_ID>` con el ID guardado en `config/sheet_id.md`.

Expected: crea `.clasp.json` con `scriptId` y `appsscript.json` con manifest minimo. Tambien crea un `Code.js` que ignoraremos (lo borramos en step siguiente).

- [ ] **Step 4: Limpiar archivos default y reemplazar manifest**

Borrar `apps_script/Code.js` si existe.

Reemplazar contenido de `apps_script/appsscript.json` con:

```json
{
  "timeZone": "America/Chihuahua",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```

- [ ] **Step 5: Crear `.claspignore` para excluir archivos no-Apps-Script**

```
**/**
!appsscript.json
!*.gs
```

- [ ] **Step 6: Probar `clasp push`**

```bash
clasp push
```
Expected: `└─ appsscript.json`. Sin errores.

Verificar: abrir el Sheet en Drive, menu `Extensiones > Apps Script` debe abrir el proyecto.

- [ ] **Step 7: Commit**

```bash
cd /c/proyectos/abarrotes_eri
git add apps_script/appsscript.json apps_script/.claspignore
git -c user.name="Omar" -c user.email="omarhs12@gmail.com" commit -m "chore: setup clasp con Apps Script project bound al Sheet"
```

Nota: `apps_script/.clasp.json` queda local (gitignored porque contiene `scriptId`).

---

## Task 4: Test framework helpers (assertions + runner)

**Files:**
- Create: `apps_script/tests.gs`

Apps Script no tiene framework de tests nativo. Implementamos un mini-runner.

- [ ] **Step 1: Escribir test inicial que falla intencionalmente para validar el runner**

Crear `apps_script/tests.gs`:

```javascript
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
```

- [ ] **Step 2: Push a Apps Script**

```bash
cd /c/proyectos/abarrotes_eri/apps_script
clasp push
```

- [ ] **Step 3: Ejecutar `runAllTests` desde Apps Script editor**

1. Abrir `Extensiones > Apps Script` en el Sheet
2. Seleccionar funcion `runAllTests` en el dropdown superior
3. Click `Ejecutar`
4. Autorizar la primera vez
5. Ver "Registros de ejecucion" (Ctrl+Enter o menu Ver)

Expected: `PASS  test_runner_smoke` y `1 passed, 0 failed (de 1 totales)`.

- [ ] **Step 4: Commit**

```bash
cd /c/proyectos/abarrotes_eri
git add apps_script/tests.gs
git -c user.name="Omar" -c user.email="omarhs12@gmail.com" commit -m "test: agregar mini framework de tests para Apps Script"
```

---

## Task 5: Schema definicion + setupSheets() (crear las 12 hojas con headers)

**Files:**
- Create: `apps_script/schema.gs`
- Create: `apps_script/lib.gs`
- Create: `apps_script/setup_sheets.gs`
- Modify: `apps_script/tests.gs`

- [ ] **Step 1: Crear `apps_script/lib.gs` con helpers comunes**

```javascript
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  const sheet = getSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error(`Hoja '${name}' no existe. Corre createSheets() primero.`);
  }
  return sheet;
}

function getSheetOrNull(name) {
  return getSpreadsheet().getSheetByName(name);
}
```

- [ ] **Step 2: Crear `apps_script/schema.gs` con la definicion de las 12 hojas**

```javascript
// SCHEMA: source of truth de la estructura de hojas.
// Cada hoja: { name, headers: [...] }.
// El orden de los headers define el orden de columnas.

const SCHEMA = [
  {
    name: 'productos',
    headers: ['sku', 'nombre', 'categoria', 'unidad_venta', 'peso_kg_unidad', 'precio_venta', 'stock_minimo', 'activo']
  },
  {
    name: 'proveedores',
    headers: ['id', 'nombre', 'contacto', 'forma_pago_default', 'notas']
  },
  {
    name: 'compras',
    headers: ['id_compra', 'fecha', 'proveedor_id', 'forma_pago', 'subtotal', 'iva', 'total', 'estatus', 'notas']
  },
  {
    name: 'compras_detalle',
    headers: ['id_compra', 'sku', 'lote', 'cantidad', 'costo_unitario', 'caducidad', 'destino']
  },
  {
    name: 'inventario',
    headers: ['sku', 'lote', 'cantidad_actual', 'caducidad', 'costo_unitario', 'ubicacion']
  },
  {
    name: 'clientes',
    headers: ['id_cliente', 'nombre', 'contacto', 'tipo', 'limite_credito', 'saldo_actual', 'activo']
  },
  {
    name: 'ventas',
    headers: ['id_venta', 'fecha', 'cliente_id', 'subtotal', 'descuento', 'total', 'peso_total_kg', 'forma_pago', 'tipo', 'estatus_impresion']
  },
  {
    name: 'ventas_detalle',
    headers: ['id_venta', 'sku', 'lote', 'cantidad', 'precio_unitario', 'peso_linea_kg']
  },
  {
    name: 'ledger_credito',
    headers: ['id_movimiento', 'fecha', 'cliente_id', 'tipo', 'referencia', 'monto', 'saldo_post']
  },
  {
    name: 'msi',
    headers: ['id_compra', 'tarjeta', 'plazo_meses', 'monto_total', 'cuota_mensual', 'fecha_inicio', 'mes_actual', 'estatus']
  },
  {
    name: 'tarjetas',
    headers: ['tarjeta', 'fecha_corte', 'fecha_pago_limite', 'saldo_aprox']
  },
  {
    name: 'config',
    headers: ['llave', 'valor', 'descripcion']
  }
];

function getSchemaFor(sheetName) {
  const found = SCHEMA.find(s => s.name === sheetName);
  if (!found) throw new Error(`Hoja '${sheetName}' no esta en el SCHEMA`);
  return found;
}

function getColumnIndex(sheetName, columnName) {
  const schema = getSchemaFor(sheetName);
  const idx = schema.headers.indexOf(columnName);
  if (idx === -1) {
    throw new Error(`Columna '${columnName}' no existe en hoja '${sheetName}'`);
  }
  return idx + 1; // Sheets es 1-indexed
}
```

- [ ] **Step 3: Escribir test que falla — `test_createSheets_creates_all_12`**

Agregar al final de `apps_script/tests.gs`:

```javascript
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
```

- [ ] **Step 4: Push y correr — confirmar que FALLA**

```bash
cd /c/proyectos/abarrotes_eri/apps_script && clasp push
```

Ejecutar `runAllTests` desde el editor.
Expected: `FAIL  test_createSheets_creates_all_12: createSheets is not defined`.

- [ ] **Step 5: Implementar `createSheets()` en `apps_script/setup_sheets.gs`**

```javascript
function createSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  SCHEMA.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
    }

    // Reset row 1 con headers definitivos
    sheet.getRange(1, 1, 1, s.headers.length)
      .setValues([s.headers])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, s.headers.length);
  });

  // Borrar hoja default "Hoja 1" / "Sheet1" si quedo vacia
  const defaults = ['Hoja 1', 'Sheet1', 'Hoja1'];
  defaults.forEach(name => {
    const s = ss.getSheetByName(name);
    if (s && s.getLastRow() === 0) ss.deleteSheet(s);
  });

  Logger.log(`createSheets: ${SCHEMA.length} hojas listas`);
}
```

- [ ] **Step 6: Push y correr — confirmar que PASA**

```bash
clasp push
```

Ejecutar `runAllTests` desde el editor.
Expected: `2 passed, 0 failed`. Verificar visualmente que el Sheet tiene las 12 pestanas.

- [ ] **Step 7: Commit**

```bash
cd /c/proyectos/abarrotes_eri
git add apps_script/schema.gs apps_script/lib.gs apps_script/setup_sheets.gs apps_script/tests.gs
git -c user.name="Omar" -c user.email="omarhs12@gmail.com" commit -m "feat(sheets): createSheets() crea las 12 hojas con headers definidos en SCHEMA"
```

---

## Task 6: Validaciones de datos (enums, fechas, numeros)

**Files:**
- Create: `apps_script/validations.gs`
- Modify: `apps_script/tests.gs`

- [ ] **Step 1: Escribir test que falla — `test_validations_reject_invalid_enum`**

Agregar a `apps_script/tests.gs`:

```javascript
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
```

- [ ] **Step 2: Push y correr — confirmar que FALLA**

```bash
clasp push
```

Expected: `FAIL  test_validations_reject_invalid_enum: setupValidations is not defined`.

- [ ] **Step 3: Implementar `setupValidations()` en `apps_script/validations.gs`**

```javascript
const VALIDATIONS = {
  productos: {
    categoria: ['abarrotes', 'bebidas', 'limpieza', 'papel', 'otros'],
    unidad_venta: ['pza', 'kg', 'lt', 'caja'],
    activo: ['SI', 'NO']
  },
  proveedores: {
    forma_pago_default: ['efectivo', 'tarjeta_a', 'tarjeta_b', 'tarjeta_c', 'MSI_3', 'MSI_6', 'MSI_9', 'MSI_12', 'MSI_18']
  },
  compras: {
    forma_pago: ['efectivo', 'tarjeta_a', 'tarjeta_b', 'tarjeta_c', 'MSI_3', 'MSI_6', 'MSI_9', 'MSI_12', 'MSI_18'],
    estatus: ['pendiente_pago', 'pagado', 'MSI_activo', 'MSI_finalizado']
  },
  compras_detalle: {
    // destino se valida con prefijo en codigo, no con dropdown (puede ser 'inventario' o 'cross_dock_<id>')
  },
  clientes: {
    tipo: ['contado', 'credito'],
    activo: ['SI', 'NO']
  },
  ventas: {
    forma_pago: ['efectivo', 'credito', 'transferencia'],
    tipo: ['normal', 'cross_dock'],
    estatus_impresion: ['impreso', 'pendiente']
  },
  ledger_credito: {
    tipo: ['cargo', 'abono']
  },
  msi: {
    estatus: ['activo', 'finalizado']
  }
};

const MAX_ROWS_FOR_VALIDATION = 5000;

function setupValidations() {
  Object.keys(VALIDATIONS).forEach(sheetName => {
    const sheet = getSheet(sheetName);
    const cols = VALIDATIONS[sheetName];

    Object.keys(cols).forEach(colName => {
      const colIdx = getColumnIndex(sheetName, colName);
      const allowed = cols[colName];

      const range = sheet.getRange(2, colIdx, MAX_ROWS_FOR_VALIDATION - 1, 1);
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(allowed, true)
        .setAllowInvalid(false)
        .setHelpText(`Valores permitidos: ${allowed.join(', ')}`)
        .build();
      range.setDataValidation(rule);
    });
  });

  Logger.log(`setupValidations: validaciones aplicadas en ${Object.keys(VALIDATIONS).length} hojas`);
}
```

- [ ] **Step 4: Push y correr — confirmar que PASA**

```bash
clasp push
```

Expected: `3 passed, 0 failed`.

- [ ] **Step 5: Commit**

```bash
cd /c/proyectos/abarrotes_eri
git add apps_script/validations.gs apps_script/tests.gs
git -c user.name="Omar" -c user.email="omarhs12@gmail.com" commit -m "feat(sheets): setupValidations() agrega data validation en columnas enum"
```

---

## Task 7: Formulas para campos derivados (saldo_actual)

**Files:**
- Create: `apps_script/formulas.gs`
- Modify: `apps_script/tests.gs`

- [ ] **Step 1: Escribir test que falla — `test_setupFormulas_saldo_actual`**

```javascript
function test_setupFormulas_saldo_actual() {
  createSheets();
  setupFormulas();

  const clientes = getSheet('clientes');
  const colSaldo = getColumnIndex('clientes', 'saldo_actual');
  const formula = clientes.getRange(2, colSaldo).getFormula();
  assertTrue(formula.length > 0, 'clientes.saldo_actual debe tener formula');
  assertTrue(formula.indexOf('ledger_credito') !== -1, 'la formula debe referenciar ledger_credito');
}
```

- [ ] **Step 2: Push y correr — confirmar que FALLA**

```bash
clasp push
```

Expected: `FAIL  test_setupFormulas_saldo_actual: setupFormulas is not defined`.

- [ ] **Step 3: Implementar `setupFormulas()` en `apps_script/formulas.gs`**

```javascript
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
```

- [ ] **Step 4: Push y correr — confirmar que PASA**

```bash
clasp push
```

Ejecutar `runAllTests`.
Expected: `4 passed, 0 failed`.

- [ ] **Step 5: Verificacion manual end-to-end**

1. En hoja `clientes` agregar fila: `id_cliente=C001`, `nombre=Test`, `tipo=credito`, `limite_credito=5000`, `activo=SI`. saldo_actual debe mostrar `0`.
2. En hoja `ledger_credito` agregar: `id_movimiento=L001`, `fecha=hoy`, `cliente_id=C001`, `tipo=cargo`, `referencia=test`, `monto=1500`.
3. Volver a `clientes` — `saldo_actual` debe mostrar `1500`.
4. En `ledger_credito` agregar: `id_movimiento=L002`, `cliente_id=C001`, `tipo=abono`, `monto=500`.
5. `saldo_actual` debe mostrar `1000`.

Borrar las filas de prueba al terminar.

- [ ] **Step 6: Commit**

```bash
cd /c/proyectos/abarrotes_eri
git add apps_script/formulas.gs apps_script/tests.gs
git -c user.name="Omar" -c user.email="omarhs12@gmail.com" commit -m "feat(sheets): setupFormulas() calcula saldo_actual de clientes desde ledger"
```

---

## Task 8: assignFolio() — onEdit trigger para id_compra y id_movimiento

**Files:**
- Create: `apps_script/folios.gs`
- Modify: `apps_script/tests.gs`

Logica: cuando alguien escribe en una fila de `compras` (en cualquier columna que no sea `id_compra`), si la columna `id_compra` esta vacia, asignar el siguiente folio incremental con prefijo `C-`. Misma logica para `ledger_credito` con prefijo `L-`.

- [ ] **Step 1: Escribir test — `test_assignFolio_compras`**

```javascript
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
```

- [ ] **Step 2: Push y correr — confirmar que FALLAN ambos**

```bash
clasp push
```

Expected: ambos test_assignFolio_* fallan con `assignFolio is not defined`.

- [ ] **Step 3: Implementar `assignFolio` en `apps_script/folios.gs`**

```javascript
const FOLIO_CONFIG = {
  compras: { idCol: 'id_compra', prefix: 'C-' },
  ledger_credito: { idCol: 'id_movimiento', prefix: 'L-' }
};

function assignFolio(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const cfg = FOLIO_CONFIG[sheetName];
  if (!cfg) return;

  const row = e.range.getRow();
  if (row === 1) return; // header

  const idColIdx = getColumnIndex(sheetName, cfg.idCol);
  const idCell = sheet.getRange(row, idColIdx);
  if (idCell.getValue() !== '') return; // ya tiene folio

  // Buscar el maximo folio actual en la columna
  const lastRow = sheet.getLastRow();
  let maxN = 0;
  if (lastRow >= 2) {
    const values = sheet.getRange(2, idColIdx, lastRow - 1, 1).getValues();
    values.forEach(([v]) => {
      if (typeof v === 'string' && v.startsWith(cfg.prefix)) {
        const n = parseInt(v.substring(cfg.prefix.length), 10);
        if (!isNaN(n) && n > maxN) maxN = n;
      }
    });
  }
  idCell.setValue(`${cfg.prefix}${maxN + 1}`);
}

// Instalar trigger onEdit programaticamente (correr una sola vez)
function installFolioTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Borrar triggers previos para evitar duplicados
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'assignFolio') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('assignFolio')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  Logger.log('Trigger onEdit -> assignFolio instalado');
}
```

- [ ] **Step 4: Push y correr — confirmar que PASAN**

```bash
clasp push
```

Ejecutar `runAllTests`.
Expected: `6 passed, 0 failed`.

- [ ] **Step 5: Instalar trigger en el Sheet**

Ejecutar `installFolioTrigger` desde el editor de Apps Script una sola vez.
Verificar: `Activadores` (icono de reloj) muestra `assignFolio` con evento `Al editar`.

- [ ] **Step 6: Verificacion manual end-to-end**

1. En hoja `compras` escribir fecha en una nueva fila — debe llenar `id_compra=C-X` automaticamente.
2. Borrar la fila al terminar.

- [ ] **Step 7: Commit**

```bash
cd /c/proyectos/abarrotes_eri
git add apps_script/folios.gs apps_script/tests.gs
git -c user.name="Omar" -c user.email="omarhs12@gmail.com" commit -m "feat(sheets): assignFolio() onEdit asigna folios incrementales en compras y ledger"
```

---

## Task 9: recalcularInventario() — derivar stock por lote (compras - ventas)

**Files:**
- Create: `apps_script/inventario.gs`
- Modify: `apps_script/tests.gs`

Logica: la hoja `inventario` es derivada. Borra y reconstruye desde `compras_detalle` (entradas con `destino=inventario`) menos `ventas_detalle` (salidas).

- [ ] **Step 1: Escribir test — `test_recalcularInventario_basic`**

```javascript
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
```

- [ ] **Step 2: Push y correr — confirmar que FALLA**

```bash
clasp push
```

Expected: `FAIL  test_recalcularInventario_basic: recalcularInventario is not defined`.

- [ ] **Step 3: Implementar `recalcularInventario` en `apps_script/inventario.gs`**

```javascript
function recalcularInventario() {
  const comprasDet = getSheet('compras_detalle');
  const ventasDet = getSheet('ventas_detalle');
  const inv = getSheet('inventario');

  // Mapa: "sku|lote" -> { sku, lote, cantidad, caducidad, costo, ubicacion }
  const lotes = {};

  // ENTRADAS desde compras_detalle (solo destino=inventario)
  if (comprasDet.getLastRow() >= 2) {
    const cRows = comprasDet.getRange(2, 1, comprasDet.getLastRow() - 1, comprasDet.getLastColumn()).getValues();
    cRows.forEach(r => {
      const [idCompra, sku, lote, cantidad, costo, caducidad, destino] = r;
      if (!sku || !lote) return;
      if (destino !== 'inventario') return;
      const key = `${sku}|${lote}`;
      if (!lotes[key]) {
        lotes[key] = { sku: sku, lote: lote, cantidad: 0, caducidad: caducidad, costo: costo, ubicacion: '' };
      }
      lotes[key].cantidad += Number(cantidad) || 0;
    });
  }

  // SALIDAS desde ventas_detalle (consume del lote especifico)
  if (ventasDet.getLastRow() >= 2) {
    const vRows = ventasDet.getRange(2, 1, ventasDet.getLastRow() - 1, ventasDet.getLastColumn()).getValues();
    vRows.forEach(r => {
      const [idVenta, sku, lote, cantidad] = r;
      if (!sku || !lote) return;
      const key = `${sku}|${lote}`;
      if (lotes[key]) {
        lotes[key].cantidad -= Number(cantidad) || 0;
      }
    });
  }

  // Reescribir hoja inventario
  if (inv.getLastRow() > 1) {
    inv.getRange(2, 1, inv.getLastRow() - 1, inv.getLastColumn()).clearContent();
  }
  const out = Object.values(lotes).map(l => [l.sku, l.lote, l.cantidad, l.caducidad, l.costo, l.ubicacion]);
  if (out.length > 0) {
    inv.getRange(2, 1, out.length, 6).setValues(out);
  }

  Logger.log(`recalcularInventario: ${out.length} lotes en inventario`);
}

// Helper para consultar stock disponible total de un SKU (para validacion futura)
function stockDisponible(sku) {
  const inv = getSheet('inventario');
  if (inv.getLastRow() < 2) return 0;
  const rows = inv.getRange(2, 1, inv.getLastRow() - 1, inv.getLastColumn()).getValues();
  let total = 0;
  rows.forEach(r => {
    if (r[0] === sku) total += Number(r[2]) || 0;
  });
  return total;
}

// Helper: lotes de un SKU ordenados FEFO (caducidad ascendente)
function lotesFEFO(sku) {
  const inv = getSheet('inventario');
  if (inv.getLastRow() < 2) return [];
  const rows = inv.getRange(2, 1, inv.getLastRow() - 1, inv.getLastColumn()).getValues();
  return rows
    .filter(r => r[0] === sku && Number(r[2]) > 0)
    .map(r => ({ sku: r[0], lote: r[1], cantidad: Number(r[2]), caducidad: r[3] }))
    .sort((a, b) => {
      const da = new Date(a.caducidad).getTime();
      const db = new Date(b.caducidad).getTime();
      return da - db;
    });
}
```

- [ ] **Step 4: Push y correr — confirmar que PASA**

```bash
clasp push
```

Expected: `7 passed, 0 failed`.

- [ ] **Step 5: Commit**

```bash
cd /c/proyectos/abarrotes_eri
git add apps_script/inventario.gs apps_script/tests.gs
git -c user.name="Omar" -c user.email="omarhs12@gmail.com" commit -m "feat(sheets): recalcularInventario() reconstruye stock por lote desde compras y ventas (excluye cross-dock)"
```

---

## Task 10: recalcularSaldos() — alternativa programatica al formula-based saldo

**Files:**
- Create: `apps_script/saldos.gs`
- Modify: `apps_script/tests.gs`

Aunque `saldo_actual` ya se calcula con formula (Task 7), tener un recalculo programatico permite:
1. Reconstruir si las formulas se rompen
2. Cerrar el saldo en `ledger_credito.saldo_post` despues de cada movimiento (audit trail)

- [ ] **Step 1: Escribir test — `test_recalcularSaldos_llena_saldo_post`**

```javascript
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
```

- [ ] **Step 2: Push y correr — confirmar que FALLA**

```bash
clasp push
```

Expected: `FAIL  test_recalcularSaldos_llena_saldo_post: recalcularSaldos is not defined`.

- [ ] **Step 3: Implementar `recalcularSaldos` en `apps_script/saldos.gs`**

```javascript
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
```

- [ ] **Step 4: Push y correr — confirmar que PASA**

```bash
clasp push
```

Expected: `8 passed, 0 failed`.

- [ ] **Step 5: Commit**

```bash
cd /c/proyectos/abarrotes_eri
git add apps_script/saldos.gs apps_script/tests.gs
git -c user.name="Omar" -c user.email="omarhs12@gmail.com" commit -m "feat(sheets): recalcularSaldos() llena saldo_post en cada movimiento del ledger por cliente cronologicamente"
```

---

## Task 11: alertasCaducidad() — listar lotes que vencen pronto

**Files:**
- Create: `apps_script/alertas.gs`
- Modify: `apps_script/tests.gs`

- [ ] **Step 1: Escribir test — `test_alertasCaducidad_filtra_por_dias`**

```javascript
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
```

- [ ] **Step 2: Push y correr — confirmar que FALLA**

```bash
clasp push
```

Expected: `FAIL  test_alertasCaducidad_filtra_por_dias: alertasCaducidad is not defined`.

- [ ] **Step 3: Implementar `alertasCaducidad` en `apps_script/alertas.gs`**

```javascript
function alertasCaducidad(diasUmbral) {
  diasUmbral = diasUmbral || 30;
  const inv = getSheet('inventario');
  if (inv.getLastRow() < 2) return [];

  const rows = inv.getRange(2, 1, inv.getLastRow() - 1, inv.getLastColumn()).getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const umbralMs = diasUmbral * 24 * 3600 * 1000;

  const out = [];
  rows.forEach(r => {
    const [sku, lote, cantidad, caducidad] = r;
    if (!sku || !lote || !caducidad) return;
    if (Number(cantidad) <= 0) return;
    const cadDate = caducidad instanceof Date ? caducidad : new Date(caducidad);
    const diff = cadDate.getTime() - today.getTime();
    if (diff <= umbralMs) {
      const diasRestantes = Math.floor(diff / (24 * 3600 * 1000));
      out.push({ sku: sku, lote: lote, cantidad: Number(cantidad), caducidad: cadDate, diasRestantes: diasRestantes });
    }
  });

  out.sort((a, b) => a.diasRestantes - b.diasRestantes);
  return out;
}

function mostrarAlertasCaducidad() {
  const config = getSheet('config');
  let umbral = 30;
  if (config.getLastRow() >= 2) {
    const rows = config.getRange(2, 1, config.getLastRow() - 1, 2).getValues();
    const found = rows.find(r => r[0] === 'umbral_alerta_caducidad_dias');
    if (found) umbral = Number(found[1]) || 30;
  }

  const alertas = alertasCaducidad(umbral);
  if (alertas.length === 0) {
    SpreadsheetApp.getUi().alert(`Sin alertas. Ningun lote vence en los proximos ${umbral} dias.`);
    return;
  }
  const lines = alertas.map(a =>
    `${a.sku} / lote ${a.lote}: ${a.cantidad} unidades, vence en ${a.diasRestantes} dias (${a.caducidad.toISOString().substring(0, 10)})`
  );
  SpreadsheetApp.getUi().alert(`Lotes proximos a caducar (umbral ${umbral}d):\n\n${lines.join('\n')}`);
}
```

- [ ] **Step 4: Push y correr — confirmar que PASA**

```bash
clasp push
```

Expected: `9 passed, 0 failed`.

- [ ] **Step 5: Commit**

```bash
cd /c/proyectos/abarrotes_eri
git add apps_script/alertas.gs apps_script/tests.gs
git -c user.name="Omar" -c user.email="omarhs12@gmail.com" commit -m "feat(sheets): alertasCaducidad() lista lotes proximos a vencer dado umbral"
```

---

## Task 12: Menu personalizado + onOpen() + config inicial

**Files:**
- Create: `apps_script/menu.gs`
- Modify: `apps_script/tests.gs`

- [ ] **Step 1: Escribir test — `test_config_defaults_set`**

```javascript
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
```

- [ ] **Step 2: Push y correr — confirmar que FALLA**

```bash
clasp push
```

Expected: `FAIL  test_config_defaults_set: setConfigDefaults is not defined`.

- [ ] **Step 3: Implementar `apps_script/menu.gs`**

```javascript
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Abarrotes Eri')
    .addItem('Recalcular inventario', 'recalcularInventario')
    .addItem('Recalcular saldos', 'recalcularSaldos')
    .addItem('Alertas de caducidad', 'mostrarAlertasCaducidad')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Setup (admin)')
        .addItem('Crear/reinicializar hojas', 'createSheets')
        .addItem('Aplicar validaciones', 'setupValidations')
        .addItem('Aplicar formulas', 'setupFormulas')
        .addItem('Config: defaults', 'setConfigDefaults')
        .addItem('Instalar trigger folios', 'installFolioTrigger')
        .addItem('Correr todos los tests', 'runAllTests')
    )
    .addToUi();
}

function setConfigDefaults() {
  const config = getSheet('config');
  const defaults = [
    ['capacidad_pickup_kg', 1500, 'Capacidad maxima de carga de la pickup en kg'],
    ['umbral_alerta_caducidad_dias', 30, 'Avisar caducidades dentro de N dias'],
    ['moneda', 'MXN', 'Moneda del sistema'],
    ['nombre_negocio', 'Abarrotes Eri', 'Nombre que aparece en tickets'],
    ['rfc', '', 'RFC opcional para reportes'],
    ['version_schema', '1', 'Version del schema de hojas']
  ];

  // Solo agregar las llaves que no existan
  const existing = {};
  if (config.getLastRow() >= 2) {
    const rows = config.getRange(2, 1, config.getLastRow() - 1, 1).getValues();
    rows.forEach(r => { if (r[0]) existing[r[0]] = true; });
  }

  defaults.forEach(d => {
    if (!existing[d[0]]) config.appendRow(d);
  });

  Logger.log(`setConfigDefaults: ${defaults.length} llaves verificadas`);
}

// Funcion master para setup inicial completo
function setupAll() {
  createSheets();
  setupValidations();
  setupFormulas();
  setConfigDefaults();
  installFolioTrigger();
  Logger.log('setupAll: completo. Recarga el Sheet para ver el menu nuevo.');
}
```

- [ ] **Step 4: Push y correr — confirmar que PASA**

```bash
clasp push
```

Expected: `10 passed, 0 failed`.

- [ ] **Step 5: Verificacion manual del menu**

1. Recargar el Sheet en el navegador (F5).
2. Aparece un menu "Abarrotes Eri" en la barra superior.
3. Click `Abarrotes Eri > Setup (admin) > Config: defaults` — debe llenar la hoja `config`.
4. Click `Abarrotes Eri > Alertas de caducidad` — debe mostrar dialogo (vacio o con alertas reales).

- [ ] **Step 6: Commit**

```bash
cd /c/proyectos/abarrotes_eri
git add apps_script/menu.gs apps_script/tests.gs
git -c user.name="Omar" -c user.email="omarhs12@gmail.com" commit -m "feat(sheets): menu personalizado + setupAll() + config defaults"
```

---

## Task 13: Setup inicial end-to-end en el Sheet real + README operativo

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Correr `setupAll` en el Apps Script editor**

1. Abrir Apps Script editor del Sheet.
2. Seleccionar funcion `setupAll`.
3. Ejecutar.

Expected: log `setupAll: completo. Recarga el Sheet para ver el menu nuevo.` y todas las hojas existen con headers, validaciones, formulas y config defaults llenas.

- [ ] **Step 2: Sanity check end-to-end manual**

1. Hoja `clientes`: agregar `C001`, `Cliente Test`, `5551234567`, `credito`, `5000`, (saldo_actual vacio → formula), `SI`.
2. Hoja `ledger_credito`: agregar fila — al escribir `2026-06-01` en `fecha`, `id_movimiento` debe auto-llenarse a `L-1`. Tipo=`cargo`, cliente_id=`C001`, monto=`1500`.
3. Volver a `clientes` — `saldo_actual` de C001 debe leer `1500`.
4. Hoja `compras`: agregar fila — `fecha=2026-06-01`, `proveedor_id=PROV1`, `forma_pago=efectivo`, `subtotal=500`, `iva=80`, `total=580`, `estatus=pagado`. `id_compra` debe auto-llenarse a `C-1`.
5. Borrar las filas de prueba al terminar.

- [ ] **Step 3: Reemplazar `README.md` con version operativa completa**

```markdown
# Abarrotes Eri

Sistema operativo (mini-ERP) para reventa de abarrotes en Chihuahua. Operado por Omar.

## Componentes

- **Google Sheets** — fuente de verdad (12 hojas).
  - Sheet ID: ver `config/sheet_id.md`
- **Apps Script** (`apps_script/`) — folios automaticos, recalculo de inventario FEFO, saldos, alertas de caducidad. Versionado con `clasp`.
- **PWA Android** — Fase 2 (pendiente).

## Spec
[docs/superpowers/specs/2026-06-01-abarrotes-eri-design.md](docs/superpowers/specs/2026-06-01-abarrotes-eri-design.md)

## Plan Fase 1
[docs/superpowers/plans/2026-06-01-fase-1-sheets-backoffice.md](docs/superpowers/plans/2026-06-01-fase-1-sheets-backoffice.md)

---

## Setup (one-time)

### Requisitos
- Node.js >= 18
- `@google/clasp` instalado global: `npm install -g @google/clasp`
- Cuenta de Google con acceso al Sheet

### Pasos

1. Clonar repo.
2. Crear el Sheet manualmente y guardar el ID en `config/sheet_id.md` (si es nuevo).
3. `cd apps_script && clasp login`.
4. Si es proyecto nuevo: `clasp create --type sheets --title "Abarrotes Eri - Apps Script" --parentId <SHEET_ID>`.
5. `clasp push` para subir el codigo.
6. Abrir el Sheet, menu `Abarrotes Eri > Setup (admin) > Crear/reinicializar hojas`.
7. Click `Setup (admin) > Aplicar validaciones`.
8. Click `Setup (admin) > Aplicar formulas`.
9. Click `Setup (admin) > Config: defaults`.
10. Click `Setup (admin) > Instalar trigger folios` (una sola vez).

O ejecutar `setupAll()` desde el editor de Apps Script para correr todos los pasos en uno.

---

## Operacion diaria

### Capturar una compra

1. Hoja `compras`: nueva fila con fecha, proveedor_id, forma_pago, subtotal, iva, total, estatus. (`id_compra` se llena solo).
2. Hoja `compras_detalle`: una fila por linea de la compra. Llenar `id_compra`, `sku`, `lote` (formato sugerido `YYYYMMDD-N`), cantidad, costo_unitario, caducidad, destino (`inventario` o `cross_dock_<cliente_id>`).
3. Menu `Abarrotes Eri > Recalcular inventario`.

### Registrar un abono de cliente

1. Hoja `ledger_credito`: nueva fila con fecha, cliente_id, tipo=`abono`, referencia (libre), monto. (`id_movimiento` se llena solo, `saldo_post` se actualiza al correr recalculo).
2. (Opcional) Menu `Abarrotes Eri > Recalcular saldos` para llenar `saldo_post` historicamente.

### Agregar un cliente

1. Hoja `clientes`: nueva fila. `saldo_actual` es formula automatica (no escribir manualmente).

### Agregar un producto nuevo

1. Hoja `productos`: nueva fila. `peso_kg_unidad` es importante para el calculo de carga de la pickup.

### Ver alertas de caducidad

Menu `Abarrotes Eri > Alertas de caducidad`. El umbral se configura en hoja `config` (`umbral_alerta_caducidad_dias`).

### Cross-docking (mercancia directo de proveedor a cliente)

1. Hoja `compras` y `compras_detalle`: registrar la compra con `destino=cross_dock_<id_cliente>` en cada linea de detalle.
2. Hoja `ventas` y `ventas_detalle`: registrar la venta correspondiente al mismo cliente.
3. **No recalcular inventario para esa compra** — `recalcularInventario()` ya excluye lineas con destino distinto a `inventario`.

---

## Desarrollo

```bash
cd apps_script
clasp push       # subir cambios locales al proyecto Apps Script
clasp pull       # bajar cambios hechos desde el editor web
clasp open       # abrir editor web
```

### Correr tests

Desde el editor de Apps Script, seleccionar `runAllTests` y Ejecutar. Los resultados aparecen en "Registros de ejecucion".

O via CLI (requiere Apps Script API habilitada): `clasp run runAllTests`.

---

## Roadmap

- **Fase 1 (este plan):** Sheets + Apps Script ✓
- **Fase 2:** PWA notas de venta con impresion BT
- **Fase 3:** Alertas WhatsApp, dashboards, escaneo de codigo de barras
```

- [ ] **Step 4: Commit final**

```bash
cd /c/proyectos/abarrotes_eri
git add README.md
git -c user.name="Omar" -c user.email="omarhs12@gmail.com" commit -m "docs: README operativo completo de Fase 1"
```

---

## Closeout

Al cerrar Fase 1:
- 12 hojas operativas en el Sheet, con headers, validaciones y formulas.
- Apps Script versionado en git, con tests pasando.
- Trigger onEdit instalado para folios automaticos.
- Menu personalizado en el Sheet con acciones operativas.
- README documenta setup y operacion diaria.
- 10 tests pasando.

**Siguiente:** operar manualmente 2-3 semanas (capturar compras reales, ventas reales) para validar el modelo de datos antes de iniciar Fase 2 (PWA). Si aparecen ajustes al schema durante este periodo, hacerlos en `schema.gs` y re-correr `createSheets()` (idempotente — no destruye datos existentes en columnas que sigan ahi, pero verifica antes en una copia).
