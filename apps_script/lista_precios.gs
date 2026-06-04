// Exporta lista de precios formateada para compartir via WhatsApp / imprimir.

function exportarListaPrecios() {
  const texto = generarListaPreciosTexto();
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 10px; font-size: 13px; }
      textarea { width: 100%; height: 380px; box-sizing: border-box; font-family: monospace; font-size: 12px; padding: 8px; }
      .row { margin-top: 10px; display: flex; gap: 8px; }
      .btn { padding: 8px 14px; cursor: pointer; border: 1px solid #ccc; border-radius: 4px; background: #4285f4; color: white; flex: 1; }
      .btn-secondary { background: #fff; color: #333; }
      .hint { color: #666; font-size: 12px; margin-top: 6px; }
      .copied { color: #1e8e3e; font-weight: bold; }
    </style>
    <textarea id="txt" readonly></textarea>
    <div class="hint" id="status">Copia el texto y pégalo en WhatsApp, o presiona "Copiar".</div>
    <div class="row">
      <button class="btn" onclick="copiar()">Copiar al portapapeles</button>
      <button class="btn btn-secondary" onclick="google.script.host.close()">Cerrar</button>
    </div>
    <script>
      const texto = ${JSON.stringify(texto)};
      document.getElementById('txt').value = texto;

      function copiar() {
        const ta = document.getElementById('txt');
        ta.select();
        ta.setSelectionRange(0, 99999); // móvil
        try {
          document.execCommand('copy');
          document.getElementById('status').innerHTML = '<span class="copied">✓ Copiado. Pégalo en WhatsApp o donde quieras.</span>';
        } catch (e) {
          document.getElementById('status').textContent = 'No se pudo copiar automáticamente. Selecciona el texto y copia a mano.';
        }
      }
    </script>
  `).setWidth(520).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Lista de precios');
}

// Genera el texto formateado. Separada para poder testearla.
function generarListaPreciosTexto() {
  const productos = getSheet('productos');
  const nombreNegocio = leerConfig('nombre_negocio', 'Abarrotes Eri');
  const moneda = leerConfig('moneda', 'MXN');

  if (productos.getLastRow() < 2) {
    return `*${nombreNegocio} — Lista de precios*\n\nSin productos cargados.`;
  }

  const rows = productos.getRange(2, 1, productos.getLastRow() - 1, productos.getLastColumn()).getValues();

  // Filtrar solo activos y agrupar por categoria
  const porCategoria = {};
  rows.forEach(r => {
    const [sku, nombre, categoria, unidad, peso, precio, stockMin, activo] = r;
    if (!sku || String(activo).toUpperCase() !== 'SI') return;
    const cat = categoria || 'otros';
    if (!porCategoria[cat]) porCategoria[cat] = [];
    porCategoria[cat].push({ nombre: String(nombre), unidad: String(unidad || ''), precio: Number(precio) || 0 });
  });

  // Orden de categorias (alfabético, con 'otros' al final)
  const cats = Object.keys(porCategoria).sort((a, b) => {
    if (a === 'otros') return 1;
    if (b === 'otros') return -1;
    return a.localeCompare(b);
  });

  const iconos = {
    abarrotes: '📦',
    bebidas: '🥤',
    limpieza: '🧴',
    papel: '🧻',
    otros: '🛒'
  };

  const today = new Date();
  const fecha = `${today.getDate()}/${today.getMonth() + 1}/${today.getFullYear()}`;

  let out = `*${nombreNegocio} — Lista de precios*\n_Actualizada ${fecha}_\n\n`;
  cats.forEach(cat => {
    const items = porCategoria[cat].sort((a, b) => a.nombre.localeCompare(b.nombre));
    const icono = iconos[cat] || '•';
    out += `${icono} *${cat.toUpperCase()}*\n`;
    items.forEach(it => {
      out += `• ${it.nombre} — $${formatearPrecio(it.precio)}\n`;
    });
    out += '\n';
  });
  out += `_Precios en ${moneda}. Sujeto a cambios sin previo aviso._`;
  return out;
}

function formatearPrecio(n) {
  return Number(n).toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function leerConfig(llave, fallback) {
  const config = getSheetOrNull('config');
  if (!config || config.getLastRow() < 2) return fallback;
  const rows = config.getRange(2, 1, config.getLastRow() - 1, 2).getValues();
  const found = rows.find(r => r[0] === llave);
  return found ? found[1] : fallback;
}

// ============================================================
// Lista de precios IMPRIMIBLE (hoja formateada)
// ============================================================

const PRINT_SHEET_NAME = '_lista_precios_print';

// Crea/reescribe la hoja `_lista_precios_print` con formato bonito para imprimir.
// Devuelve la hoja.
function generarListaPreciosImprimible() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let print = ss.getSheetByName(PRINT_SHEET_NAME);
  if (!print) {
    print = ss.insertSheet(PRINT_SHEET_NAME);
  }
  print.clear();

  const nombreNegocio = leerConfig('nombre_negocio', 'Abarrotes Eri');
  const moneda = leerConfig('moneda', 'MXN');
  const today = new Date();
  const fecha = `${today.getDate()}/${today.getMonth() + 1}/${today.getFullYear()}`;

  // Header
  print.getRange(1, 1).setValue(nombreNegocio.toUpperCase());
  print.getRange(1, 1, 1, 3).merge()
    .setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#fef7e0');

  print.getRange(2, 1).setValue(`Lista de precios — ${fecha}`);
  print.getRange(2, 1, 1, 3).merge()
    .setFontSize(11).setFontStyle('italic').setHorizontalAlignment('center');

  // Leer productos
  const productos = getSheet('productos');
  if (productos.getLastRow() < 2) {
    print.getRange(4, 1).setValue('Sin productos cargados.');
    return print;
  }

  const rows = productos.getRange(2, 1, productos.getLastRow() - 1, productos.getLastColumn()).getValues();
  const porCategoria = {};
  rows.forEach(r => {
    const [sku, nombre, categoria, unidad, peso, precio, stockMin, activo] = r;
    if (!sku || String(activo).toUpperCase() !== 'SI') return;
    const cat = categoria || 'otros';
    if (!porCategoria[cat]) porCategoria[cat] = [];
    porCategoria[cat].push({ nombre: String(nombre), precio: Number(precio) || 0 });
  });

  const cats = Object.keys(porCategoria).sort((a, b) => {
    if (a === 'otros') return 1;
    if (b === 'otros') return -1;
    return a.localeCompare(b);
  });

  let row = 4;
  cats.forEach(cat => {
    // Header de categoria
    print.getRange(row, 1).setValue(cat.toUpperCase());
    print.getRange(row, 1, 1, 3).merge()
      .setFontWeight('bold').setBackground('#e6f4ea').setFontSize(13)
      .setHorizontalAlignment('left');
    row++;

    // Items ordenados alfabeticamente
    const items = porCategoria[cat].sort((a, b) => a.nombre.localeCompare(b.nombre));
    items.forEach(it => {
      print.getRange(row, 1).setValue('•').setHorizontalAlignment('center');
      print.getRange(row, 2).setValue(it.nombre);
      print.getRange(row, 3).setValue(it.precio)
        .setNumberFormat('"$"#,##0.00')
        .setHorizontalAlignment('right');
      row++;
    });

    row++; // linea en blanco entre categorias
  });

  // Footer
  print.getRange(row, 1).setValue(`Precios en ${moneda}. Sujeto a cambios sin previo aviso.`);
  print.getRange(row, 1, 1, 3).merge()
    .setFontStyle('italic').setFontSize(10).setHorizontalAlignment('center');

  // Anchos de columna
  print.setColumnWidth(1, 30);
  print.setColumnWidth(2, 400);
  print.setColumnWidth(3, 110);

  // Ocultar gridlines para que se vea limpio al imprimir
  print.setHiddenGridlines(true);

  return print;
}

// Genera la hoja imprimible y la activa (usuario ve la hoja, puede Ctrl+P)
function abrirListaPreciosImprimir() {
  const print = generarListaPreciosImprimible();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(print);

  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Lista lista para imprimir',
    'La hoja "_lista_precios_print" ya esta formateada.\n\nPara imprimir: Archivo > Imprimir (Ctrl+P).\nEn el dialogo de imprimir: elige "Hoja actual" para imprimir solo esta hoja.',
    ui.ButtonSet.OK
  );
}

// ============================================================
// Lista de precios PDF (subida a Drive)
// ============================================================

const PDF_FOLDER_NAME = 'Abarrotes Eri PDFs';

function exportarListaPreciosPDF() {
  const printSheet = generarListaPreciosImprimible();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetId = ss.getId();
  const gid = printSheet.getSheetId();

  const url = `https://docs.google.com/spreadsheets/d/${sheetId}/export?` + [
    'format=pdf',
    'size=letter',
    'portrait=true',
    'fitw=true',
    `gid=${gid}`,
    'sheetnames=false',
    'printtitle=false',
    'pagenumbers=false',
    'gridlines=false',
    'fzr=false',
    'top_margin=0.5',
    'bottom_margin=0.5',
    'left_margin=0.5',
    'right_margin=0.5'
  ].join('&');

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });

  const blob = response.getBlob();
  const fecha = new Date();
  const pad = n => String(n).padStart(2, '0');
  const fechaStr = `${fecha.getFullYear()}-${pad(fecha.getMonth() + 1)}-${pad(fecha.getDate())}`;
  blob.setName(`Lista_precios_${fechaStr}.pdf`);

  const folder = getOrCreatePdfFolder();
  const file = folder.createFile(blob);
  const fileUrl = file.getUrl();
  const downloadUrl = `https://drive.google.com/uc?id=${file.getId()}&export=download`;

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 12px; font-size: 13px; }
      a { color: #4285f4; text-decoration: none; word-break: break-all; }
      .btn { padding: 8px 14px; cursor: pointer; border: 1px solid #ccc; border-radius: 4px; background: #4285f4; color: white; margin-right: 6px; }
      .btn-sec { background: #fff; color: #333; }
      .ok { background: #e6f4ea; color: #1e8e3e; padding: 8px; border-radius: 4px; margin-bottom: 12px; }
      .row { margin-top: 10px; }
      .copied { color: #1e8e3e; font-weight: bold; }
      input { width: 100%; padding: 6px; font-size: 12px; box-sizing: border-box; margin-top: 4px; }
    </style>
    <div class="ok">PDF generado y guardado en Drive</div>
    <p><b>Archivo:</b> ${file.getName()}</p>
    <p><b>Carpeta en Drive:</b> ${PDF_FOLDER_NAME}</p>
    <div class="row">
      <p><b>Enlace para compartir:</b></p>
      <input id="url" type="text" readonly value="${fileUrl}" />
    </div>
    <div class="row">
      <button class="btn" onclick="copiar()">Copiar enlace</button>
      <a href="${fileUrl}" target="_blank"><button class="btn btn-sec">Abrir en Drive</button></a>
      <button class="btn btn-sec" onclick="google.script.host.close()">Cerrar</button>
    </div>
    <div id="status" class="row"></div>
    <script>
      function copiar() {
        const ta = document.getElementById('url');
        ta.select();
        ta.setSelectionRange(0, 99999);
        try {
          document.execCommand('copy');
          document.getElementById('status').innerHTML = '<span class="copied">Enlace copiado. Pegalo en WhatsApp o correo.</span>';
        } catch (e) {
          document.getElementById('status').textContent = 'No se pudo copiar automaticamente. Selecciona y copia a mano.';
        }
      }
    </script>
  `).setWidth(540).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Lista de precios PDF');
}

function getOrCreatePdfFolder() {
  const folders = DriveApp.getFoldersByName(PDF_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PDF_FOLDER_NAME);
}
