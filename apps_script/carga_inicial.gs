// Formulario "Carga inicial de inventario": meter stock fisico existente
// sin que sea una compra real. Crea una entrada en `compras` con folio
// INICIAL-N y lineas en `compras_detalle` con destino=inventario.

const CARGA_INICIAL_PREFIX = 'INICIAL-';

function abrirFormularioCargaInicial() {
  const html = HtmlService.createHtmlOutputFromFile('form_carga_inicial')
    .setWidth(440)
    .setHeight(580);
  SpreadsheetApp.getUi().showModalDialog(html, 'Carga inicial de inventario');
}

// Devuelve el id_compra INICIAL activo. Si no existe, lo crea.
// Misma corrida de captura debe acumular bajo el mismo folio.
function obtenerOCrearCompraInicial() {
  const compras = getSheet('compras');
  const idColIdx = getColumnIndex('compras', 'id_compra');
  let maxN = 0;
  let existingFolio = null;

  if (compras.getLastRow() >= 2) {
    const ids = compras.getRange(2, idColIdx, compras.getLastRow() - 1, 1).getValues();
    ids.forEach(([v]) => {
      if (typeof v === 'string' && v.indexOf(CARGA_INICIAL_PREFIX) === 0) {
        const n = parseInt(v.substring(CARGA_INICIAL_PREFIX.length), 10);
        if (!isNaN(n) && n > maxN) {
          maxN = n;
          existingFolio = v;
        }
      }
    });
  }

  // Si ya hay un INICIAL-N hoy, reusarlo. Si no, crear INICIAL-(maxN+1).
  if (existingFolio) {
    // Verificar si la fecha del folio existente es hoy
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayMs = today.getTime();

    const fechaColIdx = getColumnIndex('compras', 'fecha');
    if (compras.getLastRow() >= 2) {
      const rows = compras.getRange(2, 1, compras.getLastRow() - 1, compras.getLastColumn()).getValues();
      const match = rows.find(r => r[idColIdx - 1] === existingFolio);
      if (match) {
        const f = match[fechaColIdx - 1];
        const fDate = f instanceof Date ? f : new Date(f);
        fDate.setHours(0, 0, 0, 0);
        if (fDate.getTime() === todayMs) return existingFolio;
      }
    }
  }

  // Crear nuevo INICIAL-N
  const nuevoFolio = `${CARGA_INICIAL_PREFIX}${maxN + 1}`;
  const today = new Date();
  // Headers de compras: id_compra, fecha, proveedor_id, forma_pago, subtotal, iva, total, estatus, notas
  compras.appendRow([nuevoFolio, today, 'CARGA_INICIAL', 'efectivo', 0, 0, 0, 'pagado', 'Carga inicial de inventario fisico existente']);
  return nuevoFolio;
}

function guardarLineaCargaInicial(data) {
  if (!data.sku) return { ok: false, error: 'SKU obligatorio' };
  if (!data.lote) return { ok: false, error: 'Lote obligatorio' };
  if (!data.cantidad || Number(data.cantidad) <= 0) return { ok: false, error: 'Cantidad debe ser > 0' };

  // Validar que el SKU existe en productos
  const productos = getSheet('productos');
  let skuExiste = false;
  if (productos.getLastRow() >= 2) {
    const skuColIdx = getColumnIndex('productos', 'sku');
    const skus = productos.getRange(2, skuColIdx, productos.getLastRow() - 1, 1).getValues();
    skuExiste = skus.some(([v]) => String(v).trim() === String(data.sku).trim());
  }
  if (!skuExiste) return { ok: false, error: `SKU '${data.sku}' no existe en productos. Crealo primero con "Nuevo producto".` };

  const sku = String(data.sku).trim();
  const lote = String(data.lote).trim();
  const cantidad = Number(data.cantidad);
  const costo = Number(data.costo_unitario) || 0;
  const caducidad = data.caducidad ? new Date(data.caducidad) : '';

  const idCompra = obtenerOCrearCompraInicial();

  const comprasDet = getSheet('compras_detalle');
  // Headers: id_compra, sku, lote, cantidad, costo_unitario, caducidad, destino
  comprasDet.appendRow([idCompra, sku, lote, cantidad, costo, caducidad, 'inventario']);

  return { ok: true, sku: sku, lote: lote, cantidad: cantidad, idCompra: idCompra };
}

function finalizarCargaInicial() {
  recalcularInventario();
  return { ok: true };
}

function contarLineasCargaInicialHoy() {
  const comprasDet = getSheet('compras_detalle');
  if (comprasDet.getLastRow() < 2) return 0;
  // Buscar el INICIAL-N de hoy
  const idCompra = obtenerOCrearCompraInicial();
  const rows = comprasDet.getRange(2, 1, comprasDet.getLastRow() - 1, comprasDet.getLastColumn()).getValues();
  return rows.filter(r => r[0] === idCompra).length;
}

// Helper para autocompletar SKU en formularios. Devuelve array de { sku, nombre, unidad }
// solo de productos activos, ordenados por nombre.
function listarProductosActivos() {
  const productos = getSheet('productos');
  if (productos.getLastRow() < 2) return [];
  const colSku = getColumnIndex('productos', 'sku') - 1;
  const colNombre = getColumnIndex('productos', 'nombre') - 1;
  const colUnidad = getColumnIndex('productos', 'unidad_venta') - 1;
  const colActivo = getColumnIndex('productos', 'activo') - 1;
  const rows = productos.getRange(2, 1, productos.getLastRow() - 1, productos.getLastColumn()).getValues();
  return rows
    .filter(r => r[colSku] && String(r[colActivo]).toUpperCase() === 'SI')
    .map(r => ({ sku: String(r[colSku]), nombre: String(r[colNombre]), unidad: String(r[colUnidad] || '') }))
    .sort((a, b) => a.nombre.localeCompare(b.nombre));
}
