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

  setupValidacionesLedger();
  setupValidacionesSKU();

  Logger.log(`setupValidations: validaciones aplicadas en ${Object.keys(VALIDATIONS).length} hojas + ledger dinamico + SKU dinamico`);
}

// Validaciones especiales de ledger_credito:
//   - cliente_id: dropdown DINAMICO desde clientes.id_cliente (se actualiza solo)
//   - referencia: lista sugerida con valores libres permitidos
function setupValidacionesLedger() {
  const ledger = getSheet('ledger_credito');
  const colCliente = getColumnIndex('ledger_credito', 'cliente_id');
  const colRef = getColumnIndex('ledger_credito', 'referencia');

  // Dropdown dinamico para cliente_id (range-based desde clientes.id_cliente)
  const clientes = getSheetOrNull('clientes');
  if (clientes) {
    const idColIdx = getColumnIndex('clientes', 'id_cliente');
    const clientesRange = clientes.getRange(2, idColIdx, clientes.getMaxRows() - 1, 1);
    const clienteRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(clientesRange, true)
      .setAllowInvalid(false)
      .setHelpText('Selecciona un cliente existente de la hoja clientes')
      .build();
    ledger.getRange(2, colCliente, MAX_ROWS_FOR_VALIDATION - 1, 1).setDataValidation(clienteRule);
  }

  // Lista sugerida para referencia (permite valores libres tambien)
  const refRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Inicial', 'abono', 'cargo'], true)
    .setAllowInvalid(true)
    .setHelpText('Sugerencias: Inicial / abono / cargo. Tambien puedes escribir libre (ej. "transferencia 12345", V-1)')
    .build();
  ledger.getRange(2, colRef, MAX_ROWS_FOR_VALIDATION - 1, 1).setDataValidation(refRule);
}

// Dropdown dinamico de SKU en hojas que referencian productos:
//   - compras_detalle.sku
//   - ventas_detalle.sku
//   - inventario.sku (aunque sea derivada, mantiene consistencia visual)
function setupValidacionesSKU() {
  const productos = getSheetOrNull('productos');
  if (!productos) return;

  const skuColIdx = getColumnIndex('productos', 'sku');
  const skuRange = productos.getRange(2, skuColIdx, productos.getMaxRows() - 1, 1);
  const skuRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(skuRange, true)
    .setAllowInvalid(false)
    .setHelpText('Selecciona un SKU existente de productos')
    .build();

  ['compras_detalle', 'ventas_detalle', 'inventario'].forEach(sheetName => {
    const sheet = getSheetOrNull(sheetName);
    if (!sheet) return;
    const colIdx = getColumnIndex(sheetName, 'sku');
    sheet.getRange(2, colIdx, MAX_ROWS_FOR_VALIDATION - 1, 1).setDataValidation(skuRule);
  });
}
