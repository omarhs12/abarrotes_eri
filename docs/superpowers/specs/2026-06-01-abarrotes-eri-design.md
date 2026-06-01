# Diseño: Abarrotes Eri — Sistema híbrido Sheets + PWA

**Fecha:** 2026-06-01
**Autor:** Omar (con asistencia de Claude)
**Estado:** Aprobado para implementación

## 1. Contexto y problema

Omar opera reventa de abarrotes, bebidas, productos de limpieza, papel, etc. en Chihuahua. La operación incluye:

- Compras a múltiples proveedores con diferentes formas de pago (efectivo, varias tarjetas, MSI)
- Inventario con productos de distintas caducidades (lotes)
- Ventas a clientes — algunos al contado, otros a crédito (saldo, cargos, abonos)
- Entregas en pickup cuya **capacidad de carga no debe excederse** (control de peso en kg por venta)
- Cross-docking: a veces mercancía va directo del proveedor al cliente sin pasar por inventario propio
- Notas de venta impresas en ticket térmico desde el celular del operador

**Bottleneck principal:** generar notas de venta es lo que más tiempo le toma a Omar.

## 2. Objetivo

Construir un sistema operativo mínimo que:

1. Lleve catálogo de productos, proveedores y clientes
2. Registre compras con desglose por lote y caducidad
3. Calcule inventario en tiempo real respetando FEFO (First Expired First Out)
4. Permita generar notas de venta desde el celular con:
   - Validación de stock disponible
   - Cálculo de peso acumulado vs. capacidad de pickup
   - Impresión en impresora térmica Bluetooth
   - Manejo de cross-docking (sin afectar inventario)
5. Lleve ledger de cargos/abonos/saldo por cliente con crédito
6. Lleve seguimiento de pagos MSI y cortes de tarjetas

## 3. Arquitectura

```
Google Sheets (fuente de verdad de datos)
   ↕ Sheets API
Cloud Function Node.js (proxy + transacciones atómicas)
   ↕ HTTPS
PWA Android Chrome
   ↕ Web Bluetooth API
Impresora térmica BT (ESC/POS)
```

**Patrón híbrido:** Sheets para captura administrativa de baja frecuencia (compras, catálogos, ajustes); PWA para captura de alta frecuencia con UX optimizada (ventas + impresión).

## 4. Modelo de datos (Google Sheets)

### `productos` (catálogo maestro)
| campo | tipo | nota |
|---|---|---|
| sku | texto (PK) | código único |
| nombre | texto | |
| categoria | enum | abarrotes, bebidas, limpieza, papel, otros |
| unidad_venta | enum | pza, kg, lt, caja |
| peso_kg_unidad | decimal | clave para calcular carga de pickup |
| precio_venta | decimal | |
| stock_minimo | decimal | para alertas |
| activo | bool | |

### `proveedores`
| campo | tipo |
|---|---|
| id | texto (PK) |
| nombre | texto |
| contacto | texto |
| forma_pago_default | enum |
| notas | texto |

### `compras` (encabezado)
| campo | tipo |
|---|---|
| id_compra | auto (folio interno) |
| fecha | fecha |
| proveedor_id | FK |
| forma_pago | enum: efectivo, tarjeta_X, MSI_X_meses |
| subtotal | decimal |
| iva | decimal |
| total | decimal |
| estatus | enum: pendiente_pago, pagado, MSI_activo, MSI_finalizado |
| notas | texto |

### `compras_detalle`
| campo | tipo | nota |
|---|---|---|
| id_compra | FK | |
| sku | FK | |
| lote | texto | identificador único por SKU (formato sugerido `YYYYMMDD-N`) |
| cantidad | decimal | |
| costo_unitario | decimal | |
| caducidad | fecha | |
| destino | enum | `inventario` o `cross_dock_<cliente_id>` |

### `inventario` (vista derivada por Apps Script)
| campo | tipo | nota |
|---|---|---|
| sku | FK | |
| lote | texto | |
| cantidad_actual | decimal | computado: entradas − salidas |
| caducidad | fecha | |
| costo_unitario | decimal | |
| ubicacion | texto | opcional (bodega, camioneta) |

### `clientes`
| campo | tipo |
|---|---|
| id_cliente | texto (PK) |
| nombre | texto |
| contacto | texto |
| tipo | enum: contado, credito |
| limite_credito | decimal |
| saldo_actual | decimal (computado del ledger) |
| activo | bool |

### `ventas` (encabezado — `id_venta` también es el folio del ticket)
| campo | tipo |
|---|---|
| id_venta | auto (folio) |
| fecha | timestamp |
| cliente_id | FK |
| subtotal | decimal |
| descuento | decimal |
| total | decimal |
| peso_total_kg | decimal |
| forma_pago | enum: efectivo, credito, transferencia |
| tipo | enum: normal, cross_dock |
| estatus_impresion | enum: impreso, pendiente |

### `ventas_detalle`
| campo | tipo |
|---|---|
| id_venta | FK |
| sku | FK |
| lote | texto (qué lote específico consumió, FEFO) |
| cantidad | decimal |
| precio_unitario | decimal |
| peso_linea_kg | decimal |

### `ledger_credito`
| campo | tipo |
|---|---|
| id_movimiento | auto |
| fecha | fecha |
| cliente_id | FK |
| tipo | enum: cargo, abono |
| referencia | texto (id_venta si cargo, nota si abono) |
| monto | decimal |
| saldo_post | decimal |

### `msi`
| campo | tipo |
|---|---|
| id_compra | FK |
| tarjeta | texto |
| plazo_meses | int |
| monto_total | decimal |
| cuota_mensual | decimal |
| fecha_inicio | fecha |
| mes_actual | int |
| estatus | enum: activo, finalizado |

### `tarjetas`
| campo | tipo |
|---|---|
| tarjeta | texto (PK) |
| fecha_corte | int (día del mes) |
| fecha_pago_limite | int (día del mes) |
| saldo_aprox | decimal (manual) |

### `config` (parámetros operativos)
| llave | valor |
|---|---|
| capacidad_pickup_kg | decimal |
| umbral_alerta_caducidad_dias | int |
| moneda | texto (MXN) |

## 5. Lógica clave

### FEFO al vender
Al registrar una venta para un SKU, consumir lotes en orden de caducidad ascendente. Si la cantidad excede el stock disponible total del SKU, rechazar la venta (salvo cross-docking).

### Cross-docking
La PWA permite marcar una venta como `tipo = cross_dock`. En ese caso, el Cloud Function:
1. Crea un registro en `compras` y `compras_detalle` con `destino = cross_dock_<cliente_id>`
2. Crea el registro de venta correspondiente
3. **No incrementa ni decrementa inventario**

### Cálculo de peso
`peso_total_kg` por venta = Σ (cantidad × `peso_kg_unidad`) por línea.
La PWA muestra peso acumulado vs. `config.capacidad_pickup_kg`:
- Verde: < 70%
- Amarillo: 70-90%
- Rojo: > 90% (alerta, permite continuar)
- Bloqueo: > 100% (botón "Imprimir y guardar" deshabilitado)

### Saldo de cliente
`saldo_actual` = Σ cargos − Σ abonos del `ledger_credito`. Apps Script lo refresca con fórmula o trigger.

### Folios
- `id_compra`: capturada desde Sheets → Apps Script `onEdit` asigna el siguiente folio
- `id_venta`: capturada desde PWA → Cloud Function asigna el siguiente folio dentro de la misma transacción de escritura
- `lote`: asignado manualmente al capturar compra en Sheets (formato sugerido `YYYYMMDD-N` donde N es secuencial dentro del día)

## 6. PWA — alcance Fase 2

### Pantallas
1. **Login** — PIN simple, token de sesión guardado en localStorage
2. **Nueva nota de venta**:
   - Buscador de cliente (autocomplete desde `clientes`)
   - Indicador de saldo actual y crédito disponible (si aplica)
   - Lista de líneas: buscar producto → cantidad → validar stock → agregar línea
   - Total subtotal/total
   - Peso acumulado con semáforo
   - Toggle "Cross-docking" (oculta validación de stock)
   - Selector de forma de pago
   - Botón "Imprimir y guardar"
3. **Historial del día** (lectura, últimas N ventas locales)
4. **Configuración** — URL del backend, capacidad pickup local, impresora BT pareada

### Flujo de impresión
1. Generar payload ESC/POS con `esc-pos-encoder`
2. Conectar a impresora vía Web Bluetooth (la primera vez se aparea, queda guardada)
3. Enviar payload por GATT characteristic
4. Best-effort: la mayoría de impresoras térmicas BT no confirman, asumir éxito si no hubo error de conexión

### Sincronización con backend
Una sola llamada `POST /venta` al Cloud Function con todo el payload. El Cloud Function:
1. Valida stock atómicamente (relee inventario, verifica disponibilidad)
2. Asigna folio (`id_venta`)
3. Escribe `ventas`, `ventas_detalle`, decrementa lotes en `inventario`, escribe `ledger_credito` (si crédito) — todo en un `batchUpdate` de Sheets API
4. Devuelve folio confirmado a la PWA

### Modo offline
PWA con Service Worker + IndexedDB:
- Sin red: guarda la venta en IndexedDB, imprime localmente, marca como `pendiente_sync`
- Con red de vuelta: sincroniza pendientes con backend
- Indicador visible de ventas pendientes de sync

## 7. Stack técnico

| Capa | Tecnología |
|---|---|
| Backend datos | Google Sheets |
| Lógica interna Sheets | Google Apps Script (folios, recálculo de inventario y saldos) |
| API mediadora | Cloud Function Node.js (Cloud Functions Gen 2) |
| Auth a Sheets | Service Account JSON (solo en Cloud Function) |
| Auth de PWA al backend | PIN + token bearer |
| Frontend | HTML + JS vanilla |
| Hosting | Firebase Hosting (gratis, HTTPS) |
| Almacenamiento offline | IndexedDB |
| Service Worker | Workbox |
| Bluetooth | Web Bluetooth API + `esc-pos-encoder` (npm) |

**Costo estimado de infra:** ~$0-3 USD/mes (cabe en free tier de GCP y Firebase para este volumen).

## 8. Hardware

- **Celular Android** de Omar — Chrome reciente con Web Bluetooth habilitado
- **Impresora térmica Bluetooth ESC/POS** — modelos sugeridos: Munbyn ITPP068 o Goojprt PT-210 (~$50 USD)
  - Recomendado: alimentada por cargador 12V del camión, en soporte fijo
- **Opcional**: router con SIM en pickup (mejora conectividad pero no es bloqueante; BT imprime sin internet)

## 9. Fases de entrega

### Fase 1 — Backoffice en Sheets (1-2 días de setup + 2-3 semanas operando)
- Crear estructura de hojas (12 hojas con encabezados y validaciones)
- Apps Script para folios automáticos y recálculo de inventario/saldos
- Fórmulas y validaciones de datos
- Capturar catálogo inicial: ~50-100 productos, proveedores principales, clientes actuales
- Operar manualmente unas semanas (capturar compras y ventas reales sin PWA)
- **Salida:** modelo de datos validado contra realidad, ajustes hechos antes de Fase 2

### Fase 2 — PWA notas de venta (1-2 semanas)
- Setup Firebase Hosting + Cloud Function
- UI mínima de venta
- Integración con Sheets vía Cloud Function
- Web Bluetooth + impresión ESC/POS
- Modo cross-docking
- Modo offline (Service Worker + IndexedDB)
- **Salida:** Omar genera notas de venta desde el cel con impresión BT en ruta

### Fase 3 — Mejoras (cuando duela)
- Alertas WhatsApp/SMS de caducidades próximas (3, 7, 30 días)
- Dashboard semanal (ventas, márgenes, top clientes/productos)
- Escaneo de código de barras (BarcodeDetector API o Quagga.js)
- Recordatorios MSI (próximo pago)
- Reportes contables
- Migración a terminal Sunmi POS si crece el volumen

## 10. Riesgos y mitigaciones

| Riesgo | Mitigación |
|---|---|
| Concurrencia en escrituras | Cloud Function serializa con `batchUpdate` atómico |
| Pérdida de venta por estar offline | Service Worker + IndexedDB; sync al recuperar red |
| Impresora no responde | Reintentar 2 veces; permitir guardar venta sin imprimir (`estatus_impresion=pendiente`) y reimprimir después |
| Driver ESC/POS de modelo desconocido | Comprar impresora **antes** de iniciar Fase 2 y validar con tests |
| Sheets se vuelve lento con miles de filas | Archivar histórico anual a hojas `_archivo_AAAA`; mantener vista activa pequeña |
| Saldos descuadrados por error humano | Apps Script `recalcular_saldos()` reconstruye desde el ledger |
| Capacidad de pickup mal estimada | Parámetro en hoja `config`; ajustable sin redeploy |
| Service Account expuesta | Vive solo en Cloud Function; PWA nunca la ve |

## 11. Decisiones tomadas

- **Híbrido pragmático** (no Sheets puro ni app completa) — Sheets para captura administrativa, PWA solo donde más duele (notas de venta)
- **Lote como entidad de inventario** (no solo SKU) — necesario para controlar caducidades
- **FEFO** como política de consumo — minimiza pérdidas por caducidad
- **Cloud Function intermedia** — atomicidad de escrituras y protección de credenciales
- **Web Bluetooth, no WiFi/TCP** — limitación del navegador; BT es más simple, más barato, imprime offline
- **Sin framework JS pesado** (vanilla HTML/JS) — minimiza mantenimiento, la app es pequeña

## 12. Decisiones pospuestas

- Marca/modelo exacto de impresora BT (decidir antes de comprar Fase 2 — validar driver ESC/POS)
- Capacidad exacta de pickup en kg (calibrar con la real)
- Estructura final de categorías de productos (refinar durante captura inicial)
- Política de respaldo de Sheets (versionado nativo de Google es suficiente al inicio)
- Diseño visual del ticket impreso (logo, encabezado, formato)
