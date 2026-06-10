---
name: Pedido rapido (MVP 2 dias)
status: approved
date: 2026-06-09
---

# Pedido rapido — MVP 2 dias

## Contexto

Omar necesita capturar pedidos HOY (proximos 2 dias) sin cargar inventario inicial, FEFO ni ledger. Operacion: selecciona productos con cantidades, ve el total, manda el ticket por WhatsApp o lo imprime. Nada se persiste — es solo armar y mandar.

## Scope (in)

- Menu nuevo: **Abarrotes Eri > Nuevo pedido** (arriba del menu).
- Formulario HTML dialog que lee `productos` filtrado por `activo='SI'`.
- Buscador por nombre, tabla con cantidad editable por fila, total en vivo.
- Salida WhatsApp: texto formateado al portapapeles.
- Salida Imprimir: hoja temporal `_pedido_print` formateada → dialogo de impresion.

## Fuera de scope (deliberado)

- Cliente, forma de pago, descuentos.
- Folio persistente o guardar el pedido en `ventas`/`ventas_detalle`.
- Tocar inventario, ledger_credito o cualquier otra hoja del sistema.
- Tests unitarios formales (es throwaway de 2 dias, validacion manual).

## Arquitectura

Sigue el patron de los formularios existentes (`form_producto.html` + `forms_producto.gs`).

**Archivos nuevos:**
- `apps_script/pedido_rapido.gs` — backend (carga catalogo, formatea ticket WhatsApp, prepara hoja imprimible).
- `apps_script/form_pedido.html` — frontend (UI del dialogo).

**Archivo modificado:**
- `apps_script/menu.gs` — agregar `Nuevo pedido` como primer item.

## Flujo

1. Usuario abre `Nuevo pedido` desde el menu.
2. Backend `abrirFormularioPedido()` lee `productos` (sku, nombre, precio_venta) filtrando `activo='SI'`, ordenado por categoria + nombre. Pasa array al HTML.
3. UI muestra:
   - Input search arriba (filtra por substring en nombre, case-insensitive).
   - Tabla: `[nombre]  [precio unit]  [cantidad input]  [subtotal]`.
   - Footer: `Total: $X,XXX.XX` (recalcula con cada cambio).
   - Botones: `WhatsApp` | `Imprimir` | `Cancelar`.
4. Click `WhatsApp`:
   - Frontend manda al backend el array `[{nombre, precio, cantidad}, ...]` con `cantidad > 0`.
   - Backend devuelve el texto formateado.
   - Frontend lo copia al portapapeles y muestra "Copiado, pega en WhatsApp".
5. Click `Imprimir`:
   - Backend recibe el array, crea/limpia hoja oculta `_pedido_print`, escribe encabezado + lineas + total formateado.
   - Frontend abre la hoja en una pestaña nueva con `?gid=...&printview=true` o similar; alternativa: `SpreadsheetApp.setActiveSheet(_pedido_print)` y dejar que el usuario haga Ctrl+P. Decision durante implementacion (lo que sea mas simple en Apps Script).
6. Click `Cancelar`: cierra dialogo.

## Formato ticket WhatsApp

```
*Abarrotes Eri*
Pedido 09/06/2026

• 3 x Coca Cola 2L c/8 .......... $825.00
• 2 x Arizona Lata 460ml c/24 ... $698.00
• 1 x Electrolit 625ml c/6 ...... $115.00

*Total: $1,638.00*
```

Reglas:
- Fecha en formato `dd/mm/yyyy`.
- Lineas alineadas con puntos (padding dinamico segun el nombre mas largo).
- Numeros con coma como separador de miles, 2 decimales.
- Si el nombre es muy largo (>30 chars), trunca con ellipsis para no romper el alineado en moviles.

## Formato hoja imprimible

Hoja `_pedido_print` (oculta por default, se muestra al imprimir):
- Fila 1: "Abarrotes Eri" (centrado, bold, 16pt).
- Fila 2: `Pedido del DD/MM/YYYY` (centrado, 11pt).
- Fila 4+: tabla `Cant | Producto | P.Unit | Subtotal`.
- Fila final + 2: `TOTAL: $X,XXX.XX` (bold, derecha).
- Pie: "Gracias por su compra".

## Riesgos / consideraciones

- **Datos `precio_venta` con simbolo `$`**: los productos que mandaste muestran `$349.00`. Si en la hoja `productos` esta guardado como texto (no como numero con formato), `parseFloat` necesita limpiar `$` y comas. La logica del backend debe ser robusta a ambos formatos (numero o string con simbolos).
- **Catalogo grande**: 21 productos activos cabe sin paginacion. Si en el futuro crece, agregar virtual scroll.
- **No hay validacion de stock**: si Omar vende mas de lo que tiene, el sistema no se entera. Aceptable para 2 dias.

## Criterio de exito

- Omar arma un pedido de 5 productos en <60 segundos desde abrir el menu.
- Texto WhatsApp se copia con un click y se pega legible en WhatsApp movil.
- Hoja imprimible sale en una pagina A4/carta normal.
- Cero impacto en hojas existentes (compras, ventas, inventario, ledger).
