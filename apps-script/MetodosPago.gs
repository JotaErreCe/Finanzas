/**
 * MÓDULO v3 — MÉTODOS DE PAGO
 * Archivo NUEVO: no modifica nada del sistema v2.
 *
 * Construye la hoja "Métodos de Pago": cuánto se gasta con cada tarjeta
 * (identificada por sus últimos 4 dígitos) y con transferencias de cada
 * cuenta, por período financiero (del 10 al 9).
 *
 * Las transferencias entre cuentas propias nunca llegan aquí porque el
 * módulo de Transferencias las desvía a "Transferencias Internas" antes
 * de que entren a Transacciones.
 *
 * Trigger requerido (setupV3 lo crea automáticamente):
 *   actualizarMetodosPago | Por tiempo | Cada hora
 */

const CONFIG_MP = {
  HOJA: "Métodos de Pago",
  NUM_PERIODOS: 6 // período actual + 5 anteriores
};

function actualizarMetodosPago() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTx = ss.getSheetByName(CONFIG.HOJA_TRANSACCIONES);
  const tc = obtenerTipoCambio(ss);

  const ultimaFila = hojaTx.getLastRow();
  if (ultimaFila < 5) return;
  const datos = hojaTx.getRange(5, 1, ultimaFila - 4, 10).getValues();

  const periodos = obtenerUltimosPeriodos(CONFIG_MP.NUM_PERIODOS);

  // Acumular por método × período.
  // metodos = { "Crédito GyT •••4447": { totales: [p0..p5], txActual: n } }
  const metodos = {};

  for (const r of datos) {
    const [fecha, , monto, moneda, , metodo, banco, , , notas] = r;
    if (!(fecha instanceof Date) || !monto) continue;

    const idx = periodos.findIndex(p => fecha >= p.inicio && fecha <= p.fin);
    if (idx === -1) continue;

    const etiqueta = etiquetaMetodo(metodo, banco, (notas || "").toString());
    if (!metodos[etiqueta]) {
      metodos[etiqueta] = { totales: new Array(periodos.length).fill(0), txActual: 0 };
    }
    const montoGTQ = moneda === "USD" ? monto * tc : monto;
    metodos[etiqueta].totales[idx] += montoGTQ;
    if (idx === periodos.length - 1) metodos[etiqueta].txActual++;
  }

  escribirHojaMetodos(ss, metodos, periodos);
}

/**
 * Etiqueta legible de cada método:
 *   TC            → "Crédito GyT •••4447"   (tarjeta desde las notas)
 *   Débito        → "Débito GyT •••1234"
 *   Transferencia → "Transferencia •••2741" (cuenta origen desde las notas)
 *   Efectivo      → "Efectivo"
 */
function etiquetaMetodo(metodo, banco, notas) {
  const tarjeta = notas.match(/Tarjeta\s+•••(\d{4})/);
  const cuenta = notas.match(/Cta\s+•••(\d{4})/);

  if (metodo === "TC") return `Crédito ${banco}` + (tarjeta ? ` •••${tarjeta[1]}` : "");
  if (metodo === "Débito") return `Débito ${banco}` + (tarjeta ? ` •••${tarjeta[1]}` : "");
  if (metodo === "Transferencia") return "Transferencia" + (cuenta ? ` •••${cuenta[1]}` : "");
  if (metodo === "Efectivo") return "Efectivo";
  return metodo || "Sin método";
}

/**
 * Devuelve los últimos N períodos financieros (10 → 9), del más viejo
 * al actual. Reutiliza la misma lógica de ciclo que obtenerPeriodoFinanciero,
 * manejando correctamente los cambios de mes y de año.
 */
function obtenerUltimosPeriodos(n) {
  const actual = obtenerPeriodoFinanciero();
  const periodos = [];
  for (let i = n - 1; i >= 0; i--) {
    const inicio = new Date(actual.inicio.getFullYear(), actual.inicio.getMonth() - i, CONFIG.DIA_INICIO_MES, 0, 0, 0);
    const fin = new Date(actual.inicio.getFullYear(), actual.inicio.getMonth() - i + 1, CONFIG.DIA_INICIO_MES - 1, 23, 59, 59);
    periodos.push({
      inicio: inicio,
      fin: fin,
      label: Utilities.formatDate(inicio, "GMT-6", "d MMM") + " – " + Utilities.formatDate(fin, "GMT-6", "d MMM")
    });
  }
  return periodos;
}

function escribirHojaMetodos(ss, metodos, periodos) {
  let hoja = ss.getSheetByName(CONFIG_MP.HOJA);
  if (!hoja) hoja = ss.insertSheet(CONFIG_MP.HOJA);
  hoja.clear();

  const idxActual = periodos.length - 1;
  const totalActual = Object.values(metodos).reduce((s, m) => s + m.totales[idxActual], 0);

  // Ordenar por gasto del período actual, de mayor a menor.
  const orden = Object.entries(metodos).sort((a, b) => b[1].totales[idxActual] - a[1].totales[idxActual]);

  const filas = [];
  filas.push(["MÉTODOS DE PAGO", "", "", "", "", "", "", "", ""]);
  filas.push(["Actualizado: " + Utilities.formatDate(new Date(), "GMT-6", "d MMM yyyy HH:mm"), "", "", "", "", "", "", "", ""]);
  filas.push(["Montos en GTQ. Excluye transferencias entre cuentas propias.", "", "", "", "", "", "", "", ""]);
  filas.push([]);
  filas.push(["Método", ...periodos.map(p => p.label), "# Tx (actual)", "% (actual)"]);

  for (const [nombre, info] of orden) {
    filas.push([
      nombre,
      ...info.totales.map(t => Math.round(t * 100) / 100),
      info.txActual,
      totalActual > 0 ? Math.round((info.totales[idxActual] / totalActual) * 1000) / 10 + "%" : "—"
    ]);
  }

  filas.push([]);
  filas.push([
    "TOTAL",
    ...periodos.map((p, i) => Math.round(Object.values(metodos).reduce((s, m) => s + m.totales[i], 0) * 100) / 100),
    Object.values(metodos).reduce((s, m) => s + m.txActual, 0),
    ""
  ]);

  const numCols = periodos.length + 3;
  hoja.getRange(1, 1, filas.length, numCols)
    .setValues(filas.map(f => {
      while (f.length < numCols) f.push("");
      return f.slice(0, numCols);
    }));

  // Formato básico.
  hoja.getRange(1, 1).setFontWeight("bold").setFontSize(14);
  hoja.getRange(5, 1, 1, numCols).setFontWeight("bold").setBackground("#f0f0f0");
  hoja.getRange(filas.length, 1, 1, numCols).setFontWeight("bold");
  hoja.getRange(6, 2, filas.length - 5, periodos.length).setNumberFormat("Q#,##0.00");
  hoja.setFrozenRows(5);
  hoja.autoResizeColumns(1, numCols);
}
