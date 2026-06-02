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
