# Abarrotes Eri

Sistema operativo (mini-ERP) para reventa de abarrotes en Chihuahua. Operado por Omar.

## Componentes

- **Google Sheets** — fuente de verdad (12 hojas).
  - Sheet ID: ver `config/sheet_id.md`
- **Apps Script** (`apps_script/`) — folios automaticos, recalculo de inventario FEFO, saldos, alertas de caducidad. Versionado con `clasp`.
- **PWA Android** — Fase 2 (pendiente).

## Spec
[docs/superpowers/specs/2026-06-01-abarrotes-eri-design.md](docs/superpowers/specs/2026-06-01-abarrotes-eri-design.md)

## Plan Fase 1
[docs/superpowers/plans/2026-06-01-fase-1-sheets-backoffice.md](docs/superpowers/plans/2026-06-01-fase-1-sheets-backoffice.md)

---

## Setup (one-time)

### Requisitos
- Node.js >= 18
- `@google/clasp` instalado global: `npm install -g @google/clasp`
- Cuenta de Google con acceso al Sheet

### Pasos

1. Clonar repo.
2. Crear el Sheet manualmente y guardar el ID en `config/sheet_id.md` (si es nuevo).
3. `cd apps_script && clasp login`.
4. Si es proyecto nuevo: `clasp create --type sheets --title "Abarrotes Eri - Apps Script" --parentId <SHEET_ID>`.
5. `clasp push` para subir el codigo.
6. Abrir el Sheet, menu `Abarrotes Eri > Setup (admin) > Crear/reinicializar hojas`.
7. Click `Setup (admin) > Aplicar validaciones`.
8. Click `Setup (admin) > Aplicar formulas`.
9. Click `Setup (admin) > Config: defaults`.
10. Click `Setup (admin) > Instalar trigger folios` (una sola vez).

O ejecutar `setupAll()` desde el editor de Apps Script para correr todos los pasos en uno.

---

## Operacion diaria

### Capturar una compra

1. Hoja `compras`: nueva fila con fecha, proveedor_id, forma_pago, subtotal, iva, total, estatus. (`id_compra` se llena solo).
2. Hoja `compras_detalle`: una fila por linea de la compra. Llenar `id_compra`, `sku`, `lote` (formato sugerido `YYYYMMDD-N`), cantidad, costo_unitario, caducidad, destino (`inventario` o `cross_dock_<cliente_id>`).
3. Menu `Abarrotes Eri > Recalcular inventario`.

### Registrar un abono de cliente

1. Hoja `ledger_credito`: nueva fila con fecha, cliente_id, tipo=`abono`, referencia (libre), monto. (`id_movimiento` se llena solo, `saldo_post` se actualiza al correr recalculo).
2. (Opcional) Menu `Abarrotes Eri > Recalcular saldos` para llenar `saldo_post` historicamente.

### Agregar un cliente

1. Hoja `clientes`: nueva fila. `saldo_actual` es formula automatica (no escribir manualmente).

### Agregar un producto nuevo

1. Hoja `productos`: nueva fila. `peso_kg_unidad` es importante para el calculo de carga de la pickup.

### Ver alertas de caducidad

Menu `Abarrotes Eri > Alertas de caducidad`. El umbral se configura en hoja `config` (`umbral_alerta_caducidad_dias`).

### Cross-docking (mercancia directo de proveedor a cliente)

1. Hoja `compras` y `compras_detalle`: registrar la compra con `destino=cross_dock_<id_cliente>` en cada linea de detalle.
2. Hoja `ventas` y `ventas_detalle`: registrar la venta correspondiente al mismo cliente.
3. **No recalcular inventario para esa compra** — `recalcularInventario()` ya excluye lineas con destino distinto a `inventario`.

---

## Desarrollo

```bash
cd apps_script
clasp push       # subir cambios locales al proyecto Apps Script
clasp pull       # bajar cambios hechos desde el editor web
clasp open       # abrir editor web
```

### Correr tests

Desde el editor de Apps Script, seleccionar `runAllTests` y Ejecutar. Los resultados aparecen en "Registros de ejecucion".

O via CLI (requiere Apps Script API habilitada): `clasp run runAllTests`.

---

## Roadmap

- **Fase 1 (este plan):** Sheets + Apps Script
- **Fase 2:** PWA notas de venta con impresion BT
- **Fase 3:** Alertas WhatsApp, dashboards, escaneo de codigo de barras
