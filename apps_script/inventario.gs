function recalcularInventario() {
  const comprasDet = getSheet('compras_detalle');
  const ventasDet = getSheet('ventas_detalle');
  const inv = getSheet('inventario');

  // Mapa: "sku|lote" -> { sku, lote, cantidad, caducidad, costo, ubicacion }
  const lotes = {};

  // ENTRADAS desde compras_detalle (solo destino=inventario)
  if (comprasDet.getLastRow() >= 2) {
    const cRows = comprasDet.getRange(2, 1, comprasDet.getLastRow() - 1, comprasDet.getLastColumn()).getValues();
    cRows.forEach(r => {
      const [idCompra, sku, lote, cantidad, costo, caducidad, destino] = r;
      if (!sku || !lote) return;
      if (destino !== 'inventario') return;
      const key = `${sku}|${lote}`;
      if (!lotes[key]) {
        lotes[key] = { sku: sku, lote: lote, cantidad: 0, caducidad: caducidad, costo: costo, ubicacion: '' };
      }
      lotes[key].cantidad += Number(cantidad) || 0;
    });
  }

  // SALIDAS desde ventas_detalle (consume del lote especifico)
  if (ventasDet.getLastRow() >= 2) {
    const vRows = ventasDet.getRange(2, 1, ventasDet.getLastRow() - 1, ventasDet.getLastColumn()).getValues();
    vRows.forEach(r => {
      const [idVenta, sku, lote, cantidad] = r;
      if (!sku || !lote) return;
      const key = `${sku}|${lote}`;
      if (lotes[key]) {
        lotes[key].cantidad -= Number(cantidad) || 0;
      }
    });
  }

  // Reescribir hoja inventario
  if (inv.getLastRow() > 1) {
    inv.getRange(2, 1, inv.getLastRow() - 1, inv.getLastColumn()).clearContent();
  }
  const out = Object.values(lotes).map(l => [l.sku, l.lote, l.cantidad, l.caducidad, l.costo, l.ubicacion]);
  if (out.length > 0) {
    inv.getRange(2, 1, out.length, 6).setValues(out);
  }

  Logger.log(`recalcularInventario: ${out.length} lotes en inventario`);
}

// Helper para consultar stock disponible total de un SKU (para validacion futura)
function stockDisponible(sku) {
  const inv = getSheet('inventario');
  if (inv.getLastRow() < 2) return 0;
  const rows = inv.getRange(2, 1, inv.getLastRow() - 1, inv.getLastColumn()).getValues();
  let total = 0;
  rows.forEach(r => {
    if (r[0] === sku) total += Number(r[2]) || 0;
  });
  return total;
}

// Helper: lotes de un SKU ordenados FEFO (caducidad ascendente)
function lotesFEFO(sku) {
  const inv = getSheet('inventario');
  if (inv.getLastRow() < 2) return [];
  const rows = inv.getRange(2, 1, inv.getLastRow() - 1, inv.getLastColumn()).getValues();
  return rows
    .filter(r => r[0] === sku && Number(r[2]) > 0)
    .map(r => ({ sku: r[0], lote: r[1], cantidad: Number(r[2]), caducidad: r[3] }))
    .sort((a, b) => {
      const da = new Date(a.caducidad).getTime();
      const db = new Date(b.caducidad).getTime();
      return da - db;
    });
}
