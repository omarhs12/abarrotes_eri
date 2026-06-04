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

// Migracion: copia nombre -> sku para todos los productos.
// Idempotente: solo actualiza filas donde sku != nombre.
function migrarSkuAlNombre() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Migrar SKU = Nombre',
    'Esta accion copia la columna "nombre" a la columna "sku" en todos los productos. ' +
    'Si tienes compras o ventas registradas que referencian los SKU actuales, esas referencias quedaran rotas (avisa primero).\n\n' +
    'Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const productos = getSheet('productos');
  if (productos.getLastRow() < 2) {
    ui.alert('Sin productos para migrar.');
    return;
  }

  const colSku = getColumnIndex('productos', 'sku');
  const colNombre = getColumnIndex('productos', 'nombre');
  const rows = productos.getRange(2, 1, productos.getLastRow() - 1, productos.getLastColumn()).getValues();

  let cambios = 0;
  rows.forEach((row, i) => {
    const sku = String(row[colSku - 1] || '').trim();
    const nombre = String(row[colNombre - 1] || '').trim();
    if (nombre && sku !== nombre) {
      productos.getRange(i + 2, colSku).setValue(nombre);
      cambios++;
    }
  });

  ui.alert(`${cambios} producto(s) migrado(s). El SKU ahora es igual al nombre.`);
}

function contarProductosActivos() {
  const productos = getSheet('productos');
  if (productos.getLastRow() < 2) return 0;
  const activoCol = getColumnIndex('productos', 'activo');
  const rows = productos.getRange(2, activoCol, productos.getLastRow() - 1, 1).getValues();
  return rows.filter(r => String(r[0]).toUpperCase() === 'SI').length;
}
