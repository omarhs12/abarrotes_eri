function alertasCaducidad(diasUmbral) {
  diasUmbral = diasUmbral || 30;
  const inv = getSheet('inventario');
  if (inv.getLastRow() < 2) return [];

  const rows = inv.getRange(2, 1, inv.getLastRow() - 1, inv.getLastColumn()).getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const umbralMs = diasUmbral * 24 * 3600 * 1000;

  const out = [];
  rows.forEach(r => {
    const [sku, lote, cantidad, caducidad] = r;
    if (!sku || !lote || !caducidad) return;
    if (Number(cantidad) <= 0) return;
    const cadDate = caducidad instanceof Date ? caducidad : new Date(caducidad);
    const diff = cadDate.getTime() - today.getTime();
    if (diff <= umbralMs) {
      const diasRestantes = Math.floor(diff / (24 * 3600 * 1000));
      out.push({ sku: sku, lote: lote, cantidad: Number(cantidad), caducidad: cadDate, diasRestantes: diasRestantes });
    }
  });

  out.sort((a, b) => a.diasRestantes - b.diasRestantes);
  return out;
}

function mostrarAlertasCaducidad() {
  const config = getSheet('config');
  let umbral = 30;
  if (config.getLastRow() >= 2) {
    const rows = config.getRange(2, 1, config.getLastRow() - 1, 2).getValues();
    const found = rows.find(r => r[0] === 'umbral_alerta_caducidad_dias');
    if (found) umbral = Number(found[1]) || 30;
  }

  const alertas = alertasCaducidad(umbral);
  if (alertas.length === 0) {
    SpreadsheetApp.getUi().alert(`Sin alertas. Ningun lote vence en los proximos ${umbral} dias.`);
    return;
  }
  const lines = alertas.map(a =>
    `${a.sku} / lote ${a.lote}: ${a.cantidad} unidades, vence en ${a.diasRestantes} dias (${a.caducidad.toISOString().substring(0, 10)})`
  );
  SpreadsheetApp.getUi().alert(`Lotes proximos a caducar (umbral ${umbral}d):\n\n${lines.join('\n')}`);
}
