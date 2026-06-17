// Pedido rapido — armado de pedidos.
// WhatsApp: texto formateado (no persiste nada).
// Imprimir: hoja `Pedidos-DD-MM-AA` por dia, pedidos acumulados (mas reciente arriba),
//           cada ticket duplicado lado a lado para carta horizontal.

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
        peso: Number(peso) || 0,
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

// Lista de nombres de clientes activos para el dropdown del formulario.
function cargarClientesActivos() {
  const sheet = getSheetOrNull('clientes');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  return rows
    .map(r => ({
      nombre: String(r[1] || ''),               // col 'nombre'
      activo: String(r[6] || '').toUpperCase()  // col 'activo'
    }))
    .filter(c => c.nombre && c.activo === 'SI')
    .sort((a, b) => a.nombre.localeCompare(b.nombre))
    .map(c => c.nombre);
}

// Acepta numero (123.45), o string ('$349.00', '1,234.56'). Devuelve numero.
function parsePrecio(v) {
  if (typeof v === 'number') return v;
  const s = String(v || '').replace(/[$,\s]/g, '');
  const n = parseFloat(s);
  return isFinite(n) ? n : 0;
}

// lineas: [{ nombre, precio, peso, cantidad }, ...] (solo con cantidad > 0)
// Devuelve string formateado para WhatsApp.
function generarTicketPedidoTexto(lineas) {
  const nombreNegocio = leerConfig('nombre_negocio', 'Abarrotes Eri');
  const today = new Date();
  const fecha = `${pad2(today.getDate())}/${pad2(today.getMonth() + 1)}/${today.getFullYear()}`;

  if (!lineas || !lineas.length) {
    return `*${nombreNegocio}*\nPedido ${fecha}\n\n_Sin productos._`;
  }

  const MAX_NAME = 30;
  const items = lineas.map(l => {
    const nombre = String(l.nombre || '');
    const display = nombre.length > MAX_NAME ? nombre.slice(0, MAX_NAME - 1) + '…' : nombre;
    const cantidad = Number(l.cantidad) || 0;
    const precio = Number(l.precio) || 0;
    const peso = Number(l.peso) || 0;
    const subtotal = cantidad * precio;
    const kg = cantidad * peso;
    return { display, cantidad, subtotal, kg };
  });

  const prefixes = items.map(it => `• ${it.cantidad} x ${it.display}`);
  const maxPrefixLen = prefixes.reduce((m, p) => Math.max(m, p.length), 0);

  const lineasTxt = items.map((it, i) => {
    const prefix = prefixes[i];
    const subtotalStr = `$${formatearPrecio(it.subtotal)}`;
    const dots = '.'.repeat(Math.max(3, maxPrefixLen + 2 - prefix.length));
    return `${prefix} ${dots} ${subtotalStr}`;
  });

  const total = items.reduce((s, it) => s + it.subtotal, 0);
  const totalKg = items.reduce((s, it) => s + it.kg, 0);

  return [
    `*${nombreNegocio}*`,
    `Pedido ${fecha}`,
    '',
    ...lineasTxt,
    '',
    `*Total: $${formatearPrecio(total)}*`,
    `Peso: ${formatearPrecio(totalKg)} kg`
  ].join('\n');
}

function pad2(n) {
  return String(n).padStart(2, '0');
}

// Convierte indice 1-based de columna a letra A1 (1 -> A, 26 -> Z, 27 -> AA, ...).
function columnLetter(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// === HOJA DE PEDIDOS POR DIA ===
// Nombre `Pedidos-DD-MM-AA`. Resumen acumulado del dia en las primeras 3 filas.
// Cada nuevo pedido inserta su bloque justo despues del resumen (mas reciente arriba).
// Cada ticket se renderea DUPLICADO lado a lado (cols 1-4 + cols 6-9) en carta horizontal.

const BLOCK_ROWS = 34;       // filas por ticket (1 pagina carta horizontal con margenes razonables)
const ITEMS_DEFAULT = 16;    // items que caben sin desbordar; mas items expanden el bloque
const COLS_PER_TICKET = 4;
const COL_GAP_IDX = 5;       // columna separadora entre original y copia
const PESO_MAX_KG = 600;     // mismo limite que el form

function nombreHojaPedidos() {
  const d = new Date();
  return `Pedidos-${pad2(d.getDate())}-${pad2(d.getMonth() + 1)}-${String(d.getFullYear()).slice(-2)}`;
}

function formatFolio(date, seq) {
  const dd = pad2(date.getDate());
  const mm = pad2(date.getMonth() + 1);
  const aa = String(date.getFullYear()).slice(-2);
  return `P-${dd}${mm}${aa}-${pad2(seq)}`;
}

function computeBlockRows(itemCount) {
  return BLOCK_ROWS + Math.max(0, itemCount - ITEMS_DEFAULT);
}

function obtenerOCrearHojaPedidos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombre = nombreHojaPedidos();
  let sheet = ss.getSheetByName(nombre);
  if (sheet) return { sheet, esNueva: false };

  sheet = ss.insertSheet(nombre);
  inicializarHojaPedidos(sheet);
  guardarResumen({ total: 0, kg: 0, count: 0 });
  return { sheet, esNueva: true };
}

function inicializarHojaPedidos(sheet) {
  // Anchos columna pensados para carta horizontal con duplicado lado a lado
  // + concentrado lateral en cols 11-13 (no entra en la pagina impresa)
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 95);
  sheet.setColumnWidth(COL_GAP_IDX, 30);
  sheet.setColumnWidth(6, 50);
  sheet.setColumnWidth(7, 250);
  sheet.setColumnWidth(8, 80);
  sheet.setColumnWidth(9, 95);
  sheet.setColumnWidth(10, 30);   // separador entre duplicado y concentrado
  sheet.setColumnWidth(11, 55);   // cant
  sheet.setColumnWidth(12, 260);  // producto
  sheet.setColumnWidth(13, 110);  // peso total

  sheet.setHiddenGridlines(true);
}

function leerResumen() {
  const props = PropertiesService.getDocumentProperties();
  const key = `pedidos_${nombreHojaPedidos()}`;
  const raw = props.getProperty(key);
  const defecto = { total: 0, kg: 0, count: 0, productos: {} };
  if (!raw) return defecto;
  try {
    const r = JSON.parse(raw);
    return {
      total: r.total || 0,
      kg: r.kg || 0,
      count: r.count || 0,
      productos: r.productos || {}
    };
  } catch (e) {
    return defecto;
  }
}

// Acumula cantidades por producto en el resumen del dia.
// productosAcum: objeto { nombreProducto: { cantidad, peso } }.
function acumularProductos(productosAcum, items) {
  items.forEach(it => {
    const key = it.nombre;
    if (!key) return;
    if (!productosAcum[key]) {
      productosAcum[key] = { cantidad: 0, peso: it.peso || 0 };
    }
    productosAcum[key].cantidad += it.cantidad;
    // Si el primer pedido tuvo peso 0 pero uno posterior trae peso, lo actualizamos.
    if (!productosAcum[key].peso && it.peso) productosAcum[key].peso = it.peso;
  });
}

function guardarResumen(resumen) {
  const props = PropertiesService.getDocumentProperties();
  const key = `pedidos_${nombreHojaPedidos()}`;
  props.setProperty(key, JSON.stringify(resumen));
}

// Renderea el concentrado de carga del dia a la derecha del bloque del pedido mas reciente
// (cols 11-13). Se sobrescribe cada vez con el acumulado completo de resumen.productos.
function renderConcentradoLateral(sheet, blockStartRow, actualBlockRows, resumen) {
  const COL = 11;
  const COLS = 3;
  const endRow = blockStartRow + actualBlockRows - 1;

  // Encabezado verde
  sheet.getRange(blockStartRow, COL).setValue('CONCENTRADO DE CARGA');
  sheet.getRange(blockStartRow, COL, 1, COLS).merge()
    .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#0f9d58').setFontColor('white');

  // Subtitulo con peso total / limite
  const totalKg = Number(resumen.kg) || 0;
  sheet.getRange(blockStartRow + 1, COL)
    .setValue(`${formatearPrecio(totalKg)} kg  /  Max ${PESO_MAX_KG} kg`);
  sheet.getRange(blockStartRow + 1, COL, 1, COLS).merge()
    .setFontSize(11).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground(totalKg > PESO_MAX_KG ? '#fce8e6' : '#e8f0fe');

  // Header de la tabla
  const tableRow = blockStartRow + 3;
  sheet.getRange(tableRow, COL).setValue('Cant').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(tableRow, COL + 1).setValue('Producto').setFontWeight('bold');
  sheet.getRange(tableRow, COL + 2).setValue('Peso').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange(tableRow, COL, 1, COLS).setBackground('#e6f4ea')
    .setBorder(true, true, true, true, false, false);

  // Items ordenados alfabeticamente
  const productos = resumen.productos || {};
  const nombres = Object.keys(productos).sort((a, b) => a.localeCompare(b));
  const filaTotal = endRow;        // total al fondo del bloque
  const maxItemRow = filaTotal - 2; // dejar al menos 1 fila de gap antes del total
  let r = tableRow + 1;
  let sumaKg = 0;
  nombres.forEach(n => {
    const p = productos[n];
    const cant = Number(p.cantidad) || 0;
    const pesoUnit = Number(p.peso) || 0;
    const pesoTotal = cant * pesoUnit;
    sumaKg += pesoTotal;
    if (r > maxItemRow) return;
    sheet.getRange(r, COL).setValue(cant).setHorizontalAlignment('center').setFontWeight('bold');
    sheet.getRange(r, COL + 1).setValue(n);
    sheet.getRange(r, COL + 2).setValue(pesoTotal).setNumberFormat('#,##0.00" kg"').setHorizontalAlignment('right');
    r++;
  });

  // Total al final del bloque
  sheet.getRange(filaTotal, COL + 1).setValue('TOTAL:').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange(filaTotal, COL + 2).setValue(sumaKg).setNumberFormat('#,##0.00" kg"')
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('right')
    .setBackground(sumaKg > PESO_MAX_KG ? '#fce8e6' : '#fff2cc');

  // Borde verde alrededor del bloque del concentrado
  sheet.getRange(blockStartRow, COL, actualBlockRows, COLS)
    .setBorder(true, true, true, true, false, false, '#0f9d58', SpreadsheetApp.BorderStyle.SOLID);
}

// Renderea un ticket completo a partir de (startRow, startCol) ocupando actualBlockRows filas.
// data: { folio, hora, cliente, items: [{cantidad, nombre, precio}], total, totalKg }
// La copia (startCol > 1) referencia las celdas del original via formulas para que
// editar cantidad/precio en el original se propague automaticamente.
function renderTicket(sheet, startRow, startCol, actualBlockRows, data) {
  const nombreNegocio = leerConfig('nombre_negocio', 'Abarrotes Eri');
  const telefono = leerConfig('telefono_negocio', '');
  const c = startCol;
  const cols = COLS_PER_TICKET;
  const endRow = startRow + actualBlockRows - 1;
  const esCopia = startCol > 1;
  // Letras de columna para las formulas
  const colCantLetter    = columnLetter(c);
  const colPrecioLetter  = columnLetter(c + 2);
  const colSubtotLetter  = columnLetter(c + 3);

  // ---- Banner del negocio ----
  sheet.getRange(startRow, c).setValue(nombreNegocio.toUpperCase());
  sheet.getRange(startRow, c, 1, cols).merge()
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#fef7e0');

  // ---- Folio + hora ----
  sheet.getRange(startRow + 1, c).setValue(`Folio: ${data.folio}`).setFontSize(10);
  sheet.getRange(startRow + 1, c + cols - 1).setValue(data.hora)
    .setFontSize(10).setHorizontalAlignment('right');

  // Cliente: con nombre o linea en blanco para escribir
  const clienteTxt = data.cliente ? `Cliente: ${data.cliente}` : 'Cliente: ';
  sheet.getRange(startRow + 2, c).setValue(clienteTxt)
    .setFontSize(11).setFontWeight('bold');
  sheet.getRange(startRow + 2, c, 1, cols).merge()
    .setBorder(false, false, true, false, false, false);

  // ---- Tabla ----
  const tableHeaderRow = startRow + 4;
  sheet.getRange(tableHeaderRow, c).setValue('Cant').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(tableHeaderRow, c + 1).setValue('Producto').setFontWeight('bold');
  sheet.getRange(tableHeaderRow, c + 2).setValue('P.Unit').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange(tableHeaderRow, c + 3).setValue('Subtotal').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange(tableHeaderRow, c, 1, cols).setBackground('#e6f4ea')
    .setBorder(true, true, true, true, false, false);

  // Items — original tiene valores editables; copia referencia al original.
  // El subtotal siempre es formula local (cant * precio del mismo lado).
  const firstItemRow = tableHeaderRow + 1;
  let r = firstItemRow;
  data.items.forEach(it => {
    if (esCopia) {
      sheet.getRange(r, c).setFormula(`=A${r}`).setHorizontalAlignment('center');
      sheet.getRange(r, c + 1).setFormula(`=B${r}`);
      sheet.getRange(r, c + 2).setFormula(`=C${r}`).setNumberFormat('"$"#,##0.00').setHorizontalAlignment('right');
    } else {
      sheet.getRange(r, c).setValue(it.cantidad).setHorizontalAlignment('center');
      sheet.getRange(r, c + 1).setValue(it.nombre);
      sheet.getRange(r, c + 2).setValue(it.precio).setNumberFormat('"$"#,##0.00').setHorizontalAlignment('right');
    }
    sheet.getRange(r, c + 3)
      .setFormula(`=${colCantLetter}${r}*${colPrecioLetter}${r}`)
      .setNumberFormat('"$"#,##0.00').setHorizontalAlignment('right');
    r++;
  });
  const lastItemRow = r - 1;

  // PESO: 2 renglones debajo del ultimo producto (separado del TOTAL)
  const pesoRow = lastItemRow + 2;
  sheet.getRange(pesoRow, c + 2).setValue('PESO:').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange(pesoRow, c + 3).setValue(data.totalKg).setNumberFormat('#,##0.00" kg"')
    .setFontWeight('bold').setHorizontalAlignment('right')
    .setBackground(data.totalKg > PESO_MAX_KG ? '#fce8e6' : '#fff2cc');
  if (data.totalKg > PESO_MAX_KG) {
    sheet.getRange(pesoRow, c + 2).setNote(`Excede el limite de ${PESO_MAX_KG} kg de la troca`);
  }

  // ---- Footer (offsets desde el FIN del bloque para alineacion consistente) ----
  const totalRow  = endRow - 9;
  const notasRow  = endRow - 7;
  const firmaRow  = endRow - 3;
  const piePieRow = endRow;

  // TOTAL — formula para que se actualice al editar cantidades/precios.
  // Copia referencia al total del original (D{totalRow}) para que ambos siempre coincidan.
  sheet.getRange(totalRow, c + 2).setValue('TOTAL:').setFontWeight('bold').setHorizontalAlignment('right');
  const totalCell = sheet.getRange(totalRow, c + 3)
    .setNumberFormat('"$"#,##0.00').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('right');
  if (data.items.length === 0) {
    totalCell.setValue(0);
  } else if (esCopia) {
    totalCell.setFormula(`=D${totalRow}`);
  } else {
    totalCell.setFormula(`=SUM(${colSubtotLetter}${firstItemRow}:${colSubtotLetter}${lastItemRow})`);
  }

  // Notas (lineas en blanco para escribir a mano)
  sheet.getRange(notasRow, c).setValue('Notas:').setFontSize(10).setFontStyle('italic');
  sheet.getRange(notasRow, c + 1, 1, cols - 1).setBorder(false, false, true, false, false, false);
  sheet.getRange(notasRow + 1, c, 1, cols).setBorder(false, false, true, false, false, false);

  // Firma
  sheet.getRange(firmaRow, c).setValue('Firma:').setFontSize(10).setFontStyle('italic');
  sheet.getRange(firmaRow, c + 1, 1, cols - 1).setBorder(false, false, true, false, false, false);

  // Pie de pagina con datos del negocio
  const piePie = telefono ? `Tel/WhatsApp: ${telefono}   ·   Gracias por su compra` : 'Gracias por su compra';
  sheet.getRange(piePieRow, c).setValue(piePie);
  sheet.getRange(piePieRow, c, 1, cols).merge()
    .setFontSize(10).setFontStyle('italic').setHorizontalAlignment('center');

  // Borde alrededor del bloque para que se vea como ticket
  sheet.getRange(startRow, c, actualBlockRows, cols)
    .setBorder(true, true, true, true, false, false, '#666', SpreadsheetApp.BorderStyle.SOLID);
}

function generarPedidoImprimible(lineas, cliente) {
  const { sheet, esNueva } = obtenerOCrearHojaPedidos();
  const resumen = esNueva ? { total: 0, kg: 0, count: 0, productos: {} } : leerResumen();

  const items = (lineas || []).map(l => ({
    cantidad: Number(l.cantidad) || 0,
    nombre: String(l.nombre || ''),
    precio: Number(l.precio) || 0,
    peso: Number(l.peso) || 0
  }));
  const total = items.reduce((s, it) => s + it.cantidad * it.precio, 0);
  const totalKg = items.reduce((s, it) => s + it.cantidad * it.peso, 0);

  const now = new Date();
  const folio = formatFolio(now, resumen.count + 1);
  const hora = `${pad2(now.getHours())}:${pad2(now.getMinutes())}`;

  const data = {
    folio,
    hora,
    cliente: String(cliente || '').trim(),
    items,
    total,
    totalKg
  };

  const actualBlockRows = computeBlockRows(items.length);
  const insertAt = 1;

  // Insertar filas al inicio y limpiar formato heredado en cols 1-13
  sheet.insertRowsBefore(insertAt, actualBlockRows);
  sheet.getRange(insertAt, 1, actualBlockRows, 13).clearFormat();

  // Renderear original (cols 1-4) y copia (cols 6-9)
  renderTicket(sheet, insertAt, 1, actualBlockRows, data);
  renderTicket(sheet, insertAt, COL_GAP_IDX + 1, actualBlockRows, data);

  // Columna separadora entre original y copia
  sheet.getRange(insertAt, COL_GAP_IDX, actualBlockRows, 1).setBackground('#f0f0f0');

  // Actualizar resumen acumulado y concentrado por producto
  resumen.total += total;
  resumen.kg += totalKg;
  resumen.count += 1;
  acumularProductos(resumen.productos, items);
  guardarResumen(resumen);

  // Limpiar el concentrado que quedo en el bloque anterior (ahora desplazado hacia abajo)
  // y dibujarlo al lado del bloque mas reciente (cols 10-13).
  const desdeRow = insertAt + actualBlockRows;
  const maxRow = sheet.getMaxRows();
  if (desdeRow <= maxRow) {
    sheet.getRange(desdeRow, 10, maxRow - desdeRow + 1, 4)
      .clearContent().clearFormat();
  }
  renderConcentradoLateral(sheet, insertAt, actualBlockRows, resumen);

  return sheet;
}

function abrirFormularioPedido() {
  const html = HtmlService.createHtmlOutputFromFile('form_pedido')
    .setWidth(1000)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Nuevo pedido');
}

// Web App entry point — sirve el mismo form como URL accesible desde celular.
// Deploy: clasp deploy → URL https://script.google.com/macros/s/.../exec
function doGet() {
  return HtmlService.createHtmlOutputFromFile('form_pedido')
    .setTitle('Abarrotes Eri - Nuevo pedido')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Llamado desde el frontend cuando el usuario aprieta WhatsApp.
function ticketPedidoParaWhatsApp(lineas) {
  return generarTicketPedidoTexto(lineas || []);
}

// Llamado desde el frontend cuando aprieta Imprimir.
// cliente: nombre opcional del cliente seleccionado en el dropdown.
function imprimirPedido(lineas, cliente) {
  const sheet = generarPedidoImprimible(lineas || [], cliente || '');
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  return { ok: true, sheetName: sheet.getName() };
}
