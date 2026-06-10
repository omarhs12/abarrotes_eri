# Pedido Rapido Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Agregar al menu de Abarrotes Eri un formulario `Nuevo pedido` que arme un ticket de venta rapida con productos+cantidades, calcule el total, y permita mandarlo por WhatsApp (texto al portapapeles) o imprimirlo (hoja formateada).

**Architecture:** Sigue el patron de los formularios existentes (`form_producto.html` + `forms_producto.gs` + `lista_precios.gs`). Backend en `pedido_rapido.gs` con: `abrirFormularioPedido()`, `cargarCatalogoActivos()`, `generarTicketPedidoTexto(lineas)`, `generarPedidoImprimible(lineas)`. Frontend en `form_pedido.html` con tabla buscable y calculo de total en vivo. No persiste nada — es solo armar y mandar.

**Tech Stack:** Google Apps Script (V8), HtmlService modales, `clasp` para deploy.

**Test strategy:** El spec excluye tests formales (MVP 2 dias). Cada tarea termina con un protocolo de verificacion manual: comandos para `clasp push` y pasos en el Sheet para validar. No se modifica `tests.gs`.

---

## File Structure

**Files to create:**
- `apps_script/pedido_rapido.gs` — todo el backend del feature.
- `apps_script/form_pedido.html` — UI del modal.

**Files to modify:**
- `apps_script/menu.gs` — agregar item `Nuevo pedido` arriba en el menu.

**No tocar:**
- `tests.gs`, `inventario.gs`, `ventas`/`ventas_detalle` (la feature no persiste nada).
- Cualquier hoja existente excepto la creacion automatica de `_pedido_print`.

---

## Task 1: Backend skeleton + cargarCatalogoActivos()

**Files:**
- Create: `apps_script/pedido_rapido.gs`

- [ ] **Step 1: Crear `apps_script/pedido_rapido.gs` con la funcion `cargarCatalogoActivos()`**

```javascript
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
      return { nombre: String(nombre || ''), categoria: String(categoria || 'otros'), precio: parsePrecio(precio), activo: String(activo || '').toUpperCase() };
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
```

- [ ] **Step 2: Verificar carga manual via Apps Script editor**

Run en orden:
```bash
cd c:/proyectos/abarrotes_eri/apps_script
clasp push
```

Expected: `Pushed N files.` sin errores.

Luego en https://script.google.com/d/18IS7xdZL2RjBFtGL30lk9_hVIH4OHTyhuHIop3mp6wZPgL-pob-1VPQ7/edit :
- Abrir `pedido_rapido.gs`.
- Seleccionar funcion `cargarCatalogoActivos` y `Ejecutar`.
- Abrir registros (`Ver > Registros`). No deberia haber errores.
- En la consola del editor, agregar temporalmente al final del archivo:
  ```javascript
  function _debugCatalogo() {
    Logger.log(JSON.stringify(cargarCatalogoActivos().slice(0, 3)));
  }
  ```
  Ejecutar `_debugCatalogo`, verificar que el log muestre 3 items con `{nombre, categoria, precio, activo}`, precios como numeros (no strings con `$`).
- Borrar `_debugCatalogo` al confirmar que jala.

- [ ] **Step 3: Commit**

```bash
cd c:/proyectos/abarrotes_eri && git add apps_script/pedido_rapido.gs && git commit -m "feat(pedido_rapido): cargarCatalogoActivos lee productos activos con parsePrecio robusto"
```

---

## Task 2: generarTicketPedidoTexto() — formato WhatsApp

**Files:**
- Modify: `apps_script/pedido_rapido.gs`

- [ ] **Step 1: Agregar `generarTicketPedidoTexto(lineas)` al final del archivo**

```javascript
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

  // Para alinear con puntos: longitud objetivo = nombre mas largo + 2 spaces.
  const maxLen = items.reduce((m, it) => Math.max(m, it.display.length), 0);

  const lineasTxt = items.map(it => {
    const prefix = `• ${it.cantidad} x ${it.display}`;
    const subtotalStr = `$${formatearPrecio(it.subtotal)}`;
    // Padding con puntos hasta que el prefix llegue a longitud objetivo (maxLen + 7 por "N x ").
    const targetLen = maxLen + 8;
    const dots = '.'.repeat(Math.max(3, targetLen - prefix.length));
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
```

Notas:
- `formatearPrecio` y `leerConfig` ya existen en `lista_precios.gs` — se reusan tal cual (mismo namespace global de Apps Script).
- No introducir un segundo `formatearPrecio` local — usaria el global y daria warning de redeclaracion.

- [ ] **Step 2: Verificacion manual**

`clasp push` desde `apps_script/`, luego en el editor:

```javascript
function _debugTicket() {
  const t = generarTicketPedidoTexto([
    { nombre: 'Coca Cola 2L c/8', precio: 275, cantidad: 3 },
    { nombre: 'Arizona Lata 460ml c/24', precio: 349, cantidad: 2 },
    { nombre: 'Electrolit 625ml c/6', precio: 115, cantidad: 1 }
  ]);
  Logger.log(t);
}
```

Ejecutar `_debugTicket`. Expected en logs:
```
*Abarrotes Eri*
Pedido 09/06/2026

• 3 x Coca Cola 2L c/8 ............. $825.00
• 2 x Arizona Lata 460ml c/24 ...... $698.00
• 1 x Electrolit 625ml c/6 ......... $115.00

*Total: $1,638.00*
```

Borrar `_debugTicket` al confirmar.

- [ ] **Step 3: Commit**

```bash
cd c:/proyectos/abarrotes_eri && git add apps_script/pedido_rapido.gs && git commit -m "feat(pedido_rapido): generarTicketPedidoTexto formatea ticket WhatsApp con alineado de puntos"
```

---

## Task 3: generarPedidoImprimible() — hoja formateada

**Files:**
- Modify: `apps_script/pedido_rapido.gs`

- [ ] **Step 1: Agregar generador de hoja imprimible**

```javascript
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
```

- [ ] **Step 2: Verificacion manual**

`clasp push`. En el editor:

```javascript
function _debugPrint() {
  const sheet = generarPedidoImprimible([
    { nombre: 'Coca Cola 2L c/8', precio: 275, cantidad: 3 },
    { nombre: 'Arizona Lata 460ml c/24', precio: 349, cantidad: 2 },
    { nombre: 'Electrolit 625ml c/6', precio: 115, cantidad: 1 }
  ]);
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
}
```

Ejecutar. Abrir el Sheet: deberia aparecer hoja `_pedido_print` con header amarillo, tabla, total $1,638.00, y "Gracias por su compra". Borrar `_debugPrint` al confirmar.

- [ ] **Step 3: Commit**

```bash
cd c:/proyectos/abarrotes_eri && git add apps_script/pedido_rapido.gs && git commit -m "feat(pedido_rapido): generarPedidoImprimible escribe hoja _pedido_print formateada"
```

---

## Task 4: abrirFormularioPedido() + handlers del backend

**Files:**
- Modify: `apps_script/pedido_rapido.gs`

- [ ] **Step 1: Agregar el opener del modal y los handlers que llama el frontend**

```javascript
function abrirFormularioPedido() {
  const html = HtmlService.createHtmlOutputFromFile('form_pedido')
    .setWidth(640)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Nuevo pedido');
}

// Llamado desde el frontend cuando el usuario aprieta WhatsApp.
function ticketPedidoParaWhatsApp(lineas) {
  return generarTicketPedidoTexto(lineas || []);
}

// Llamado desde el frontend cuando aprieta Imprimir.
// Genera la hoja, la activa, y devuelve el nombre para que el frontend cierre el dialogo.
function imprimirPedido(lineas) {
  const sheet = generarPedidoImprimible(lineas || []);
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  return { ok: true, sheetName: sheet.getName() };
}
```

- [ ] **Step 2: Verificacion manual**

`clasp push`. En el editor ejecutar `abrirFormularioPedido` → deberia mostrar dialogo (puede salir error si `form_pedido.html` no existe todavia, eso lo arreglamos en Task 5). Si sale error "HTML file not found", es esperado.

- [ ] **Step 3: Commit**

```bash
cd c:/proyectos/abarrotes_eri && git add apps_script/pedido_rapido.gs && git commit -m "feat(pedido_rapido): abrirFormularioPedido + handlers ticketPedidoParaWhatsApp/imprimirPedido"
```

---

## Task 5: form_pedido.html — UI del dialogo

**Files:**
- Create: `apps_script/form_pedido.html`

- [ ] **Step 1: Crear el HTML completo**

```html
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; padding: 10px; font-size: 13px; margin: 0; }
    .top { position: sticky; top: 0; background: white; padding-bottom: 8px; border-bottom: 1px solid #eee; }
    #search { width: 100%; padding: 8px; box-sizing: border-box; font-size: 14px; border: 1px solid #ccc; border-radius: 4px; }
    table { width: 100%; border-collapse: collapse; margin-top: 8px; }
    th, td { padding: 6px 4px; border-bottom: 1px solid #eee; text-align: left; font-size: 13px; }
    th { background: #f8f9fa; position: sticky; top: 50px; }
    .precio { text-align: right; white-space: nowrap; }
    .qty { width: 60px; text-align: center; padding: 4px; border: 1px solid #ccc; border-radius: 3px; font-size: 13px; }
    .subtotal { text-align: right; white-space: nowrap; color: #1e8e3e; font-weight: bold; }
    .row-empty { color: #999; }
    .footer { position: sticky; bottom: 0; background: white; padding-top: 10px; border-top: 2px solid #4285f4; margin-top: 10px; }
    .total { font-size: 18px; font-weight: bold; text-align: right; padding: 8px 4px; }
    .btns { display: flex; gap: 8px; margin-top: 8px; }
    .btn { padding: 10px 14px; cursor: pointer; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; flex: 1; }
    .btn-primary { background: #4285f4; color: white; border-color: #4285f4; }
    .btn-secondary { background: #fff; color: #333; }
    .btn:disabled { opacity: 0.5; cursor: wait; }
    .status { font-size: 12px; color: #666; padding: 6px 0; min-height: 18px; }
    .copied { color: #1e8e3e; font-weight: bold; }
    .err { color: #c5221f; }
    .cat-row { background: #f1f3f4; font-weight: bold; }
  </style>
</head>
<body>
  <div class="top">
    <input id="search" type="text" placeholder="Buscar producto..." autocomplete="off" autofocus />
  </div>

  <div id="loading" style="padding: 20px; text-align: center; color: #666;">Cargando catalogo...</div>

  <table id="tbl" style="display: none;">
    <thead>
      <tr><th>Producto</th><th class="precio">P.Unit</th><th>Cant</th><th class="precio">Subtotal</th></tr>
    </thead>
    <tbody id="tbody"></tbody>
  </table>

  <div class="footer">
    <div class="total">Total: <span id="total">$0.00</span></div>
    <div class="status" id="status">Captura cantidades; el total se actualiza en vivo.</div>
    <div class="btns">
      <button class="btn btn-primary" id="btnWA" onclick="copiarWhatsApp()" disabled>WhatsApp</button>
      <button class="btn btn-primary" id="btnPrint" onclick="imprimir()" disabled>Imprimir</button>
      <button class="btn btn-secondary" onclick="google.script.host.close()">Cancelar</button>
    </div>
    <textarea id="clip" style="position: absolute; left: -9999px; top: -9999px;" aria-hidden="true"></textarea>
  </div>

  <script>
    let CATALOGO = []; // [{nombre, categoria, precio, activo}]
    const cantidades = {}; // nombre -> cantidad numerica

    google.script.run
      .withSuccessHandler(onCatalogo)
      .withFailureHandler(err => {
        document.getElementById('loading').textContent = 'Error cargando catalogo: ' + (err.message || err);
      })
      .cargarCatalogoActivos();

    function onCatalogo(items) {
      CATALOGO = items || [];
      document.getElementById('loading').style.display = 'none';
      document.getElementById('tbl').style.display = '';
      render();
      document.getElementById('search').addEventListener('input', render);
    }

    function render() {
      const q = document.getElementById('search').value.trim().toLowerCase();
      const tbody = document.getElementById('tbody');
      tbody.innerHTML = '';

      const filtrados = CATALOGO.filter(it => !q || it.nombre.toLowerCase().includes(q));
      if (!filtrados.length) {
        tbody.innerHTML = '<tr><td colspan="4" class="row-empty">Sin resultados</td></tr>';
      } else {
        let lastCat = null;
        filtrados.forEach(it => {
          if (it.categoria !== lastCat) {
            const trCat = document.createElement('tr');
            trCat.className = 'cat-row';
            trCat.innerHTML = '<td colspan="4">' + escapeHtml(it.categoria.toUpperCase()) + '</td>';
            tbody.appendChild(trCat);
            lastCat = it.categoria;
          }
          const tr = document.createElement('tr');
          const cant = cantidades[it.nombre] || '';
          const subtotal = (Number(cantidades[it.nombre]) || 0) * it.precio;
          tr.innerHTML =
            '<td>' + escapeHtml(it.nombre) + '</td>' +
            '<td class="precio">$' + fmt(it.precio) + '</td>' +
            '<td><input type="number" class="qty" min="0" step="1" data-nombre="' + escapeAttr(it.nombre) + '" value="' + cant + '" /></td>' +
            '<td class="subtotal">' + (subtotal > 0 ? '$' + fmt(subtotal) : '') + '</td>';
          tbody.appendChild(tr);
        });
      }

      tbody.querySelectorAll('input.qty').forEach(inp => {
        inp.addEventListener('input', onQtyChange);
      });

      actualizarTotal();
    }

    function onQtyChange(e) {
      const nombre = e.target.dataset.nombre;
      const v = parseInt(e.target.value, 10);
      if (!v || v <= 0) {
        delete cantidades[nombre];
      } else {
        cantidades[nombre] = v;
      }
      // Actualizar subtotal de la fila sin re-renderizar todo
      const tr = e.target.closest('tr');
      const it = CATALOGO.find(x => x.nombre === nombre);
      const sub = (v || 0) * (it ? it.precio : 0);
      tr.querySelector('.subtotal').textContent = sub > 0 ? '$' + fmt(sub) : '';
      actualizarTotal();
    }

    function actualizarTotal() {
      let total = 0;
      CATALOGO.forEach(it => {
        const c = Number(cantidades[it.nombre]) || 0;
        total += c * it.precio;
      });
      document.getElementById('total').textContent = '$' + fmt(total);
      const tieneLineas = Object.keys(cantidades).length > 0;
      document.getElementById('btnWA').disabled = !tieneLineas;
      document.getElementById('btnPrint').disabled = !tieneLineas;
    }

    function lineasParaEnviar() {
      return Object.keys(cantidades).map(nombre => {
        const it = CATALOGO.find(x => x.nombre === nombre);
        return { nombre: nombre, precio: it ? it.precio : 0, cantidad: cantidades[nombre] };
      });
    }

    function copiarWhatsApp() {
      setStatus('Generando ticket...', '');
      const btn = document.getElementById('btnWA');
      btn.disabled = true;
      google.script.run
        .withSuccessHandler(texto => {
          const ta = document.getElementById('clip');
          ta.value = texto;
          ta.select();
          ta.setSelectionRange(0, 99999);
          try {
            document.execCommand('copy');
            setStatus('Copiado. Pegalo en WhatsApp.', 'copied');
          } catch (e) {
            setStatus('No se pudo copiar automaticamente. Copia este texto:\n' + texto, 'err');
          }
          btn.disabled = false;
        })
        .withFailureHandler(err => {
          setStatus('Error: ' + (err.message || err), 'err');
          btn.disabled = false;
        })
        .ticketPedidoParaWhatsApp(lineasParaEnviar());
    }

    function imprimir() {
      setStatus('Preparando hoja...', '');
      const btn = document.getElementById('btnPrint');
      btn.disabled = true;
      google.script.run
        .withSuccessHandler(res => {
          setStatus('Hoja "' + res.sheetName + '" lista. Cierro y haz Ctrl+P.', 'copied');
          setTimeout(() => google.script.host.close(), 1200);
        })
        .withFailureHandler(err => {
          setStatus('Error: ' + (err.message || err), 'err');
          btn.disabled = false;
        })
        .imprimirPedido(lineasParaEnviar());
    }

    function setStatus(msg, cls) {
      const el = document.getElementById('status');
      el.textContent = msg;
      el.className = 'status ' + (cls || '');
    }

    function fmt(n) {
      return Number(n).toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }

    function escapeHtml(s) {
      return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    }

    function escapeAttr(s) {
      return escapeHtml(s);
    }
  </script>
</body>
</html>
```

- [ ] **Step 2: Verificacion manual end-to-end**

`clasp push`. En el editor, ejecutar `abrirFormularioPedido`. Validar:
1. Modal abre con buscador y catalogo de 21 items agrupados por categoria.
2. Escribir "coca" filtra a las Cocas activas.
3. Limpiar buscador, escribir cantidad 3 en Coca Cola 2L c/8 → subtotal $825.00 → Total cambia a $825.00.
4. Agregar cantidad 2 en Arizona Lata 460ml c/24 → Total $1,523.00.
5. Click `WhatsApp` → status dice "Copiado" → pegar en cualquier textarea, debe ser el formato del Task 2.
6. Click `Imprimir` → modal cierra, sheet activa es `_pedido_print` con la tabla.

- [ ] **Step 3: Commit**

```bash
cd c:/proyectos/abarrotes_eri && git add apps_script/form_pedido.html && git commit -m "feat(pedido_rapido): form_pedido.html con buscador, total en vivo, botones WA/imprimir"
```

---

## Task 6: Wire en menu.gs

**Files:**
- Modify: `apps_script/menu.gs`

- [ ] **Step 1: Agregar `Nuevo pedido` como primer item del menu**

Editar `menu.gs:1-35` cambiando:

```javascript
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Abarrotes Eri')
    .addItem('Nuevo producto', 'abrirFormularioProducto')
    .addItem('Nuevo cliente', 'abrirFormularioCliente')
    .addItem('Carga inicial inventario', 'abrirFormularioCargaInicial')
```

a:

```javascript
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Abarrotes Eri')
    .addItem('Nuevo pedido', 'abrirFormularioPedido')
    .addSeparator()
    .addItem('Nuevo producto', 'abrirFormularioProducto')
    .addItem('Nuevo cliente', 'abrirFormularioCliente')
    .addItem('Carga inicial inventario', 'abrirFormularioCargaInicial')
```

- [ ] **Step 2: Verificacion**

`clasp push`. Recargar el Sheet en el navegador (F5). El menu `Abarrotes Eri` debe ahora tener `Nuevo pedido` arriba, seguido de un separador y los items existentes.

Click `Nuevo pedido` → abre el modal del Task 5.

- [ ] **Step 3: Commit**

```bash
cd c:/proyectos/abarrotes_eri && git add apps_script/menu.gs && git commit -m "feat(menu): item 'Nuevo pedido' arriba con separador"
```

---

## Task 7: Verificacion end-to-end real y push final

- [ ] **Step 1: Push completo (idempotente)**

```bash
cd c:/proyectos/abarrotes_eri/apps_script && clasp push
```

Expected: `Pushed N files.`

- [ ] **Step 2: Smoke test en el Sheet real**

En https://docs.google.com/spreadsheets/d/1z4-cLilxp40Pi6i0QfPmApFhZ-5tuYgbfW79w-jSmlA/edit :
1. Recargar (F5).
2. Menu `Abarrotes Eri > Nuevo pedido`.
3. Armar pedido con 3-5 productos.
4. Click `WhatsApp`. Pegar en cualquier textarea, verificar formato.
5. Click `Nuevo pedido` otra vez (reabre limpio), armar otro pedido.
6. Click `Imprimir`. Verificar que hoja `_pedido_print` aparece formateada.
7. Ctrl+P → vista previa, "Hoja actual" → ver que cabe en una pagina.

- [ ] **Step 3: Confirmar que nada del sistema cambio**

```bash
cd c:/proyectos/abarrotes_eri && git log --oneline -10
```

Expected: 6 commits nuevos (Tasks 1-6), todos con prefijo `feat(pedido_rapido)` o `feat(menu)`. Sin cambios en `inventario.gs`, `saldos.gs`, `forms_producto.gs`, etc.

- [ ] **Step 4: Push a remoto**

```bash
cd c:/proyectos/abarrotes_eri && git push origin main
```

---

## Resumen de archivos al terminar

```
apps_script/
├── pedido_rapido.gs     (NUEVO — ~150 lineas)
├── form_pedido.html     (NUEVO — ~180 lineas)
└── menu.gs              (modificado — +2 lineas)

docs/superpowers/
├── specs/2026-06-09-pedido-rapido-design.md  (ya existe)
└── plans/2026-06-09-pedido-rapido.md         (este archivo)
```
