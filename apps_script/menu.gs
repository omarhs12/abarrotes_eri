function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Abarrotes Eri')
    .addItem('Nuevo producto', 'abrirFormularioProducto')
    .addItem('Carga inicial inventario', 'abrirFormularioCargaInicial')
    .addSeparator()
    .addItem('Recalcular inventario', 'recalcularInventario')
    .addItem('Recalcular saldos', 'recalcularSaldos')
    .addItem('Alertas de caducidad', 'mostrarAlertasCaducidad')
    .addItem('Lista de precios (WhatsApp)', 'exportarListaPrecios')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Ordenar A-Z')
        .addItem('Productos por nombre', 'ordenarProductosPorNombre')
        .addItem('Clientes por nombre', 'ordenarClientesPorNombre')
        .addItem('Proveedores por nombre', 'ordenarProveedoresPorNombre')
    )
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Setup (admin)')
        .addItem('Crear/reinicializar hojas', 'createSheets')
        .addItem('Aplicar validaciones', 'setupValidations')
        .addItem('Aplicar formulas', 'setupFormulas')
        .addItem('Config: defaults', 'setConfigDefaults')
        .addItem('Instalar trigger folios', 'installFolioTrigger')
        .addItem('Correr todos los tests', 'runAllTests')
    )
    .addToUi();
}

function setConfigDefaults() {
  const config = getSheet('config');
  const defaults = [
    ['capacidad_pickup_kg', 1500, 'Capacidad maxima de carga de la pickup en kg'],
    ['umbral_alerta_caducidad_dias', 30, 'Avisar caducidades dentro de N dias'],
    ['moneda', 'MXN', 'Moneda del sistema'],
    ['nombre_negocio', 'Abarrotes Eri', 'Nombre que aparece en tickets'],
    ['rfc', '', 'RFC opcional para reportes'],
    ['version_schema', '1', 'Version del schema de hojas']
  ];

  // Solo agregar las llaves que no existan
  const existing = {};
  if (config.getLastRow() >= 2) {
    const rows = config.getRange(2, 1, config.getLastRow() - 1, 1).getValues();
    rows.forEach(r => { if (r[0]) existing[r[0]] = true; });
  }

  defaults.forEach(d => {
    if (!existing[d[0]]) config.appendRow(d);
  });

  Logger.log(`setConfigDefaults: ${defaults.length} llaves verificadas`);
}

// Funcion master para setup inicial completo
function setupAll() {
  createSheets();
  setupValidations();
  setupFormulas();
  setConfigDefaults();
  installFolioTrigger();
  Logger.log('setupAll: completo. Recarga el Sheet para ver el menu nuevo.');
}
