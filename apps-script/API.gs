/**
 * MÓDULO v3 — API WEB PARA LA APP
 * Archivo NUEVO: no modifica nada del sistema v2.
 *
 * Expone el proyecto como Web App para que la app de GitHub Pages lea y
 * escriba datos. Reemplaza al CSV publicado (que era solo lectura y frágil).
 *
 * GET  ?token=XXX               → JSON con todo el dashboard:
 *                                 período 10→9, rubros, tendencia diaria,
 *                                 métodos de pago, checklist de fijos, atípicos.
 * POST {token, action, ...}     → addGasto | marcarPago
 *      (la app manda el body como texto plano JSON para evitar preflight CORS)
 *
 * DESPLIEGUE (una sola vez):
 *   1. Ejecutar setupV3() — genera el token (queda en los logs).
 *   2. Implementar > Nueva implementación > Aplicación web
 *      - Ejecutar como: Yo
 *      - Acceso: Cualquier usuario
 *   3. Copiar la URL /exec y pegarla junto al token en la app (⚙ Ajustes).
 */

function doGet(e) {
  try {
    if (!validarToken(e.parameter.token)) return respuestaJSON({ error: "Token inválido" });
    return respuestaJSON(construirDashboard());
  } catch (err) {
    return respuestaJSON({ error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (!validarToken(body.token)) return respuestaJSON({ error: "Token inválido" });

    if (body.action === "addGasto") return respuestaJSON(apiAddGasto(body));
    if (body.action === "marcarPago") return respuestaJSON(apiMarcarPago(body));
    return respuestaJSON({ error: "Acción desconocida: " + body.action });
  } catch (err) {
    return respuestaJSON({ error: err.message });
  }
}

function respuestaJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function validarToken(token) {
  const esperado = PropertiesService.getScriptProperties().getProperty("API_TOKEN");
  return esperado && token === esperado;
}

// ============================================================
// LECTURA: DASHBOARD COMPLETO
// ============================================================
function construirDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tc = obtenerTipoCambio(ss);
  const periodo = obtenerPeriodoFinanciero();
  const ahora = new Date();

  return {
    actualizado: Utilities.formatDate(ahora, "GMT-6", "yyyy-MM-dd HH:mm"),
    tipoCambio: tc,
    periodo: {
      inicio: Utilities.formatDate(periodo.inicio, "GMT-6", "yyyy-MM-dd"),
      fin: Utilities.formatDate(periodo.fin, "GMT-6", "yyyy-MM-dd"),
      label: Utilities.formatDate(periodo.inicio, "GMT-6", "d MMM") + " – " +
             Utilities.formatDate(periodo.fin, "GMT-6", "d MMM"),
      diaActual: Math.floor((ahora - periodo.inicio) / 86400000) + 1,
      diasTotales: Math.round((periodo.fin - periodo.inicio) / 86400000) + 1
    },
    rubros: leerRubrosDashboard(ss),
    tendencia: calcularTendencia(ss, tc, periodo),
    metodos: calcularMetodosActual(ss, tc, periodo),
    fijos: leerPagosDelMes(ahora.getFullYear(), ahora.getMonth() + 1),
    rubrosLista: leerListaRubros(ss),
    atipicos: leerAtipicos(ss)
  };
}

/**
 * Lee la hoja "Dashboard Mensual" con la misma lógica que usaba la app
 * con el CSV: filas cuyo col A es un número de rubro y col B el nombre.
 */
function leerRubrosDashboard(ss) {
  const hoja = ss.getSheetByName(CONFIG.HOJA_DASHBOARD);
  const datos = hoja.getDataRange().getValues();
  const rubros = [];
  for (const row of datos) {
    const num = parseInt(row[0]);
    const nombre = (row[1] || "").toString().trim();
    if (isNaN(num) || num < 1 || num > 30 || nombre.length < 3 || nombre === "TOTAL") continue;
    const pres = typeof row[2] === "number" ? row[2] : 0;
    const gast = typeof row[3] === "number" ? row[3] : 0;
    const disp = typeof row[4] === "number" ? row[4] : pres - gast;
    const pct = typeof row[5] === "number" ? row[5] : (pres > 0 ? gast / pres : 0);
    const txCount = typeof row[7] === "number" ? row[7] : 0;
    if (pres > 0 || gast > 0) {
      rubros.push({ nombre: nombre, pres: pres, gast: gast, disp: disp, pct: pct, txCount: txCount });
    }
  }
  return rubros;
}

/**
 * Gasto acumulado por día del período actual y del anterior (alineados por
 * día 1..N del ciclo) para la gráfica de tendencia comparativa.
 */
function calcularTendencia(ss, tc, periodo) {
  const hojaTx = ss.getSheetByName(CONFIG.HOJA_TRANSACCIONES);
  const ultimaFila = hojaTx.getLastRow();
  if (ultimaFila < 5) return [];
  const datos = hojaTx.getRange(5, 1, ultimaFila - 4, 10).getValues();

  const inicioAnterior = new Date(periodo.inicio.getFullYear(), periodo.inicio.getMonth() - 1, CONFIG.DIA_INICIO_MES, 0, 0, 0);
  const finAnterior = new Date(periodo.inicio.getTime() - 1);

  const diasTotales = Math.round((periodo.fin - periodo.inicio) / 86400000) + 1;
  const actual = new Array(diasTotales).fill(0);
  const anterior = new Array(diasTotales).fill(0);

  for (const r of datos) {
    const fecha = r[0];
    if (!(fecha instanceof Date) || !r[2]) continue;
    const montoGTQ = r[3] === "USD" ? r[2] * tc : r[2];

    if (fecha >= periodo.inicio && fecha <= periodo.fin) {
      const dia = Math.floor((fecha - periodo.inicio) / 86400000);
      if (dia >= 0 && dia < diasTotales) actual[dia] += montoGTQ;
    } else if (fecha >= inicioAnterior && fecha <= finAnterior) {
      const dia = Math.floor((fecha - inicioAnterior) / 86400000);
      if (dia >= 0 && dia < diasTotales) anterior[dia] += montoGTQ;
    }
  }

  // Acumulados. El período actual solo hasta hoy.
  const hoy = new Date();
  const diaHoy = Math.min(Math.floor((hoy - periodo.inicio) / 86400000), diasTotales - 1);
  const salida = [];
  let accA = 0, accB = 0;
  for (let i = 0; i < diasTotales; i++) {
    accA += actual[i];
    accB += anterior[i];
    salida.push({
      dia: i + 1,
      actual: i <= diaHoy ? Math.round(accA * 100) / 100 : null,
      anterior: Math.round(accB * 100) / 100
    });
  }
  return salida;
}

/**
 * Desglose de métodos de pago del período actual (para la app; la hoja
 * "Métodos de Pago" del sheet la genera actualizarMetodosPago con histórico).
 */
function calcularMetodosActual(ss, tc, periodo) {
  const hojaTx = ss.getSheetByName(CONFIG.HOJA_TRANSACCIONES);
  const ultimaFila = hojaTx.getLastRow();
  if (ultimaFila < 5) return [];
  const datos = hojaTx.getRange(5, 1, ultimaFila - 4, 10).getValues();

  const metodos = {};
  let total = 0;
  for (const r of datos) {
    const fecha = r[0];
    if (!(fecha instanceof Date) || !r[2]) continue;
    if (fecha < periodo.inicio || fecha > periodo.fin) continue;
    const etiqueta = etiquetaMetodo(r[5], r[6], (r[9] || "").toString());
    const montoGTQ = r[3] === "USD" ? r[2] * tc : r[2];
    if (!metodos[etiqueta]) metodos[etiqueta] = { total: 0, count: 0 };
    metodos[etiqueta].total += montoGTQ;
    metodos[etiqueta].count++;
    total += montoGTQ;
  }

  return Object.entries(metodos)
    .map(([nombre, m]) => ({
      nombre: nombre,
      total: Math.round(m.total * 100) / 100,
      count: m.count,
      pct: total > 0 ? Math.round((m.total / total) * 1000) / 10 : 0
    }))
    .sort((a, b) => b.total - a.total);
}

function leerListaRubros(ss) {
  const hoja = ss.getSheetByName(CONFIG.HOJA_PRESUPUESTO);
  if (hoja.getLastRow() < 4) return [];
  return hoja.getRange(4, 2, hoja.getLastRow() - 3, 1).getValues()
    .map(r => (r[0] || "").toString().trim())
    .filter(n => n.length > 0);
}

function leerAtipicos(ss) {
  const hoja = ss.getSheetByName(CONFIG.HOJA_ATIPICOS);
  if (!hoja || hoja.getLastRow() < 2) return [];
  const n = Math.min(5, hoja.getLastRow() - 1);
  const datos = hoja.getRange(hoja.getLastRow() - n + 1, 1, n, 6).getValues();
  return datos.map(r => ({
    fecha: r[0] instanceof Date ? Utilities.formatDate(r[0], "GMT-6", "yyyy-MM-dd") : "",
    comercio: r[1],
    rubro: r[2],
    monto: typeof r[3] === "number" ? Math.round(r[3] * 100) / 100 : 0,
    veces: r[5]
  })).reverse();
}

// ============================================================
// ESCRITURA
// ============================================================
function apiAddGasto(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTx = ss.getSheetByName(CONFIG.HOJA_TRANSACCIONES);
  const reglas = leerReglas(ss);

  const monto = parseFloat(body.monto);
  if (!monto || isNaN(monto) || monto <= 0) return { error: "Monto inválido" };
  const comercio = (body.comercio || "").toString().trim();
  if (!comercio) return { error: "Falta la descripción" };

  const ahora = new Date();
  const tx = {
    fecha: ahora,
    hora: Utilities.formatDate(ahora, "GMT-6", "HH:mm"),
    monto: monto,
    moneda: body.moneda === "USD" ? "USD" : "GTQ",
    comercio: comercio,
    metodo: ["Transferencia", "Efectivo", "TC", "Débito"].includes(body.metodo) ? body.metodo : "Efectivo",
    banco: "Manual",
    persona: body.persona === "Ms" ? "Ms" : "Jr",
    rubro: body.rubro && body.rubro !== "auto" ? body.rubro : clasificar(comercio, reglas),
    notas: "Ingresado desde app" + (body.notas ? " | " + body.notas : "")
  };

  escribirTransaccion(hojaTx, tx);
  verificarAtipico(ss, tx);
  verificarAlertaPresupuesto(ss, tx);
  return { ok: true, rubro: tx.rubro };
}

function apiMarcarPago(body) {
  const anio = parseInt(body.anio);
  const mes = parseInt(body.mes);
  if (isNaN(anio) || isNaN(mes)) return { error: "Año o mes inválido" };
  const ok = marcarPagoFijo(anio, mes, body.id, body.pagado === true);
  return ok ? { ok: true } : { error: "No se encontró ese pago en Control Pagos" };
}
