// Servidor del formulario "Nuevo producto"

function abrirFormularioProducto() {
  const html = HtmlService.createHtmlOutputFromFile('form_producto')
    .setWidth(420)
    .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Nuevo producto');
}

function guardarProductoDesdeFormulario(data) {
  // Validaciones mínimas
  if (!data.sku) return { ok: false, error: 'SKU es obligatorio' };
  if (!data.nombre) return { ok: false, error: 'Nombre es obligatorio' };

  const sku = String(data.sku).trim();
  const nombre = String(data.nombre).trim();
  const categoria = data.categoria || 'otros';
  const unidad = data.unidad_venta || 'pza';
  const peso = Number(data.peso_kg_unidad) || 0;
  const precio = Number(data.precio_venta) || 0;
  const stockMin = Number(data.stock_minimo) || 0;
  const activo = data.activo === 'NO' ? 'NO' : 'SI';

  // Detectar SKU duplicado
  const productos = getSheet('productos');
  if (productos.getLastRow() >= 2) {
    const skuColIdx = getColumnIndex('productos', 'sku');
    const existing = productos.getRange(2, skuColIdx, productos.getLastRow() - 1, 1).getValues();
    const dup = existing.find(r => String(r[0]).trim() === sku);
    if (dup) return { ok: false, error: `SKU '${sku}' ya existe` };
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
