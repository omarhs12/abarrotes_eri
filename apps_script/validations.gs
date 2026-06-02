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
