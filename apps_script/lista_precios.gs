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
