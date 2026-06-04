// Servidor del formulario "Nuevo producto"

function abrirFormularioProducto() {
  const html = HtmlService.createHtmlOutputFromFile('form_producto')
    .setWidth(420)
    .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Nuevo producto');
}

function guardarProductoDesdeFormulario(data) {
  if (!data.nombre) return { ok: false, error: 'Nombre es obligatorio' };

  // Decision 2026-06-03: SKU = nombre.
  // Omar prefiere identificar productos por nombre completo. La columna `sku`
  // se mantiene en el schema por compatibilidad pero su valor siempre = nombre.
  const nombre = String(data.nombre).trim();
  const sku = nombre;
  const categoria = data.categoria || 'otros';
  const unidad = data.unidad_venta || 'pza';
  const peso = Number(data.peso_kg_unidad) || 0;
  const precio = Number(data.precio_venta) || 0;
  const stockMin = Number(data.stock_minimo) || 0;
  const activo = data.activo === 'NO' ? 'NO' : 'SI';

  // Detectar duplicado por nombre (que ahora tambien es el SKU)
  const productos = getSheet('productos');
  if (productos.getLastRow() >= 2) {
    const nombreColIdx = getColumnIndex('productos', 'nombre');
    const existing = productos.getRange(2, nombreColIdx, productos.getLastRow() - 1, 1).getValues();
    const dup = existing.find(r => String(r[0]).trim() === nombre);
    if (dup) return { ok: false, error: `Producto '${nombre}' ya existe` };
  }

  productos.appendRow([sku, nombre, categoria, unidad, peso, precio, stockMin, activo]);
  return { ok: true, sku: sku, nombre: nombre };
}

function contarProductosActivos() {
  const productos = getSheet('productos');
  if (productos.getLastRow() < 2) return 0;
  const activoCol = getColumnIndex('productos', 'activo');
  const rows = productos.getRange(2, activoCol, productos.getLastRow() - 1, 1).getValues();
  return rows.filter(r => String(r[0]).toUpperCase() === 'SI').length;
}
