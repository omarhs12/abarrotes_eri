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

// lineas: [{ nombre, precio, cantidad }, ...] (solo con cantidad > 0)
// Devuelve string formateado para WhatsApp.
function generarTicketPedidoTexto(lineas) {
  const nombreNegocio = leerConfig('nombre_negocio', 'Abarrotes Eri');
  const today = new Date();
  const fecha = `${pad2(today.getDate())}/${pad2(today.getMonth() + 1)}/${today.getFullYear()}`;

  if (!lineas || !lineas.length) {
    return `*${nombreNegocio}*\nPedido ${fecha}\n\n_Sin productos._`;
  }

  // Truncar nombres muy largos para no romper alineado en mobile.
  const MAX_NAME = 30;
  const items = lineas.map(l => {
    const nombre = String(l.nombre || '');
    const display = nombre.length > MAX_NAME ? nombre.slice(0, MAX_NAME - 1) + '…' : nombre;
    const cantidad = Number(l.cantidad) || 0;
    const precio = Number(l.precio) || 0;
    const subtotal = cantidad * precio;
    return { display, cantidad, subtotal };
  });

  // Para alinear con puntos: longitud objetivo basada en nombre + cantidad prefix mas largo.
  const prefixes = items.map(it => `• ${it.cantidad} x ${it.display}`);
  const maxPrefixLen = prefixes.reduce((m, p) => Math.max(m, p.length), 0);

  const lineasTxt = items.map((it, i) => {
    const prefix = prefixes[i];
    const subtotalStr = `$${formatearPrecio(it.subtotal)}`;
    const dots = '.'.repeat(Math.max(3, maxPrefixLen + 2 - prefix.length));
    return `${prefix} ${dots} ${subtotalStr}`;
  });

  const total = items.reduce((s, it) => s + it.subtotal, 0);

  return [
    `*${nombreNegocio}*`,
    `Pedido ${fecha}`,
    '',
    ...lineasTxt,
    '',
    `*Total: $${formatearPrecio(total)}*`
  ].join('\n');
}

function pad2(n) {
  return String(n).padStart(2, '0');
}
