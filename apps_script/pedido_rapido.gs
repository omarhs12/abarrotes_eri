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

const PEDIDO_PRINT_SHEET = '_pedido_print';

// Crea/reescribe la hoja `_pedido_print` con el pedido formateado.
function generarPedidoImprimible(lineas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let print = ss.getSheetByName(PEDIDO_PRINT_SHEET);
  if (!print) {
    print = ss.insertSheet(PEDIDO_PRINT_SHEET);
  }
  print.clear();

  const nombreNegocio = leerConfig('nombre_negocio', 'Abarrotes Eri');
  const today = new Date();
  const fecha = `${pad2(today.getDate())}/${pad2(today.getMonth() + 1)}/${today.getFullYear()}`;

  // Header
  print.getRange(1, 1).setValue(nombreNegocio.toUpperCase());
  print.getRange(1, 1, 1, 4).merge()
    .setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#fef7e0');

  print.getRange(2, 1).setValue(`Pedido del ${fecha}`);
  print.getRange(2, 1, 1, 4).merge()
    .setFontSize(11).setFontStyle('italic').setHorizontalAlignment('center');

  // Tabla header
  print.getRange(4, 1).setValue('Cant').setFontWeight('bold').setHorizontalAlignment('center');
  print.getRange(4, 2).setValue('Producto').setFontWeight('bold');
  print.getRange(4, 3).setValue('P. Unit').setFontWeight('bold').setHorizontalAlignment('right');
  print.getRange(4, 4).setValue('Subtotal').setFontWeight('bold').setHorizontalAlignment('right');
  print.getRange(4, 1, 1, 4).setBackground('#e6f4ea').setBorder(true, true, true, true, false, false);

  // Lineas
  let row = 5;
  let total = 0;
  (lineas || []).forEach(l => {
    const cantidad = Number(l.cantidad) || 0;
    const precio = Number(l.precio) || 0;
    const subtotal = cantidad * precio;
    total += subtotal;
    print.getRange(row, 1).setValue(cantidad).setHorizontalAlignment('center');
    print.getRange(row, 2).setValue(String(l.nombre || ''));
    print.getRange(row, 3).setValue(precio).setNumberFormat('"$"#,##0.00').setHorizontalAlignment('right');
    print.getRange(row, 4).setValue(subtotal).setNumberFormat('"$"#,##0.00').setHorizontalAlignment('right');
    row++;
  });

  // Total
  row++;
  print.getRange(row, 3).setValue('TOTAL:').setFontWeight('bold').setHorizontalAlignment('right');
  print.getRange(row, 4).setValue(total).setNumberFormat('"$"#,##0.00')
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('right');

  // Footer
  row += 2;
  print.getRange(row, 1).setValue('Gracias por su compra');
  print.getRange(row, 1, 1, 4).merge()
    .setFontStyle('italic').setHorizontalAlignment('center');

  // Anchos
  print.setColumnWidth(1, 50);
  print.setColumnWidth(2, 320);
  print.setColumnWidth(3, 90);
  print.setColumnWidth(4, 100);
  print.setHiddenGridlines(true);

  return print;
}
