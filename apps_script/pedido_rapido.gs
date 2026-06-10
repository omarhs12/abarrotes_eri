// Pedido rapido — MVP 2 dias.
// Captura ad-hoc de pedidos: arma ticket con productos+cantidades, manda por WhatsApp o imprime.
// No persiste nada. No toca inventario, ventas, ni ledger.

function cargarCatalogoActivos() {
  const productos = getSheet('productos');
  if (productos.getLastRow() < 2) return [];

  const rows = productos.getRange(2, 1, productos.getLastRow() - 1, productos.getLastColumn()).getValues();

  const items = rows
    .map(r => {
      const [sku, nombre, categoria, unidad, peso, precio, stockMin, activo] = r;
      return {
        nombre: String(nombre || ''),
        categoria: String(categoria || 'otros'),
        precio: parsePrecio(precio),
        activo: String(activo || '').toUpperCase()
      };
    })
    .filter(it => it.nombre && it.activo === 'SI')
    .sort((a, b) => {
      const c = a.categoria.localeCompare(b.categoria);
      if (c !== 0) return c;
      return a.nombre.localeCompare(b.nombre);
    });

  return items;
}

// Acepta numero (123.45), o string ('$349.00', '1,234.56'). Devuelve numero.
function parsePrecio(v) {
  if (typeof v === 'number') return v;
  const s = String(v || '').replace(/[$,\s]/g, '');
  const n = parseFloat(s);
  return isFinite(n) ? n : 0;
}
