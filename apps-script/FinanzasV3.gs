/**
 * ════════════════════════════════════════════════════════════════
 * FINANZAS FAMILIARES — EXTENSIONES v3 (UN SOLO ARCHIVO)
 * ════════════════════════════════════════════════════════════════
 * Este archivo se agrega COMPLETO como un archivo nuevo en el
 * proyecto de Apps Script existente. El Código.gs (v2) NO se toca:
 * todo lo de aquí es independiente y convive con él.
 *
 * Contiene: transferencias ACH, métodos de pago, gastos fijos
 * (Control Maestro), API web para la app y setupV3().
 *
 * Para activar: ejecutar setupV3() una vez y luego desplegar como
 * Aplicación web (ver apps-script/README.md).
 */


// ╔══════════════════════════════════════════════════════════════╗
// ║ Transferencias                                                        
// ╚══════════════════════════════════════════════════════════════╝

/**
 * MÓDULO v3 — TRANSFERENCIAS ACH (GyT)
 * Archivo NUEVO e independiente: no modifica nada del sistema v2.
 *
 * Procesa los correos "transferencia interbancaria ACH" de GyT y:
 *  - Si el destino es una cuenta PROPIA (hoja "Cuentas Propias") → la registra
 *    en "Transferencias Internas" y NO cuenta como gasto (evita duplicados).
 *  - Si el destino es un tercero → la registra en Transacciones con
 *    método "Transferencia", clasificada con las mismas Reglas de siempre.
 *
 * Usa su propia etiqueta de Gmail y sus propias llaves de deduplicación,
 * así que puede convivir con procesarCorreos sin pisarse.
 *
 * Trigger requerido (setupV3 lo crea automáticamente):
 *   procesarTransferencias | Por tiempo | Cada 15 minutos
 */

const CONFIG_TRF = {
  REMITENTE: "bpi@gtc.com.gt",
  ETIQUETA: "Finanzas/Transferencias",
  HOJA_CUENTAS_PROPIAS: "Cuentas Propias",
  HOJA_TRANSF_INTERNAS: "Transferencias Internas",
  DIAS_BUSQUEDA: 3
};

function procesarTransferencias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTx = ss.getSheetByName(CONFIG.HOJA_TRANSACCIONES);
  const reglas = leerReglas(ss);
  const cuentasPropias = leerCuentasPropias(ss);
  const etiqueta = obtenerOCrearEtiqueta(CONFIG_TRF.ETIQUETA);
  const props = PropertiesService.getDocumentProperties();

  const fechaDesde = new Date();
  fechaDesde.setDate(fechaDesde.getDate() - CONFIG_TRF.DIAS_BUSQUEDA);
  const fechaStr = Utilities.formatDate(fechaDesde, "GMT-6", "yyyy/MM/dd");

  // Busca por frase del cuerpo; no usa la etiqueta de procesarCorreos,
  // porque ese flujo pudo haber etiquetado el hilo sin registrar nada.
  const query = `from:${CONFIG_TRF.REMITENTE} "transferencia interbancaria ACH" ` +
                `after:${fechaStr} -label:${CONFIG_TRF.ETIQUETA}`;

  const hilos = GmailApp.search(query, 0, 100);
  let procesadas = 0, internas = 0;

  for (const hilo of hilos) {
    for (const msg of hilo.getMessages()) {
      try {
        const msgId = msg.getId();
        if (props.getProperty("trf_" + msgId)) continue;

        const trf = parseEmailTransferenciaGyT(msg.getPlainBody(), msg.getDate());
        if (!trf) continue;

        const esReenvio = CONFIG.PREFIJO_REENVIO.some(p => msg.getSubject().startsWith(p));
        const persona = esReenvio ? "Ms" : "Jr";

        if (cuentasPropias.includes(trf.cuentaDestino)) {
          registrarTransferenciaInterna(ss, trf, persona);
          internas++;
        } else {
          const tx = {
            fecha: trf.fecha,
            hora: trf.hora,
            monto: trf.monto,
            moneda: trf.moneda,
            comercio: trf.comercio,
            metodo: "Transferencia",
            banco: "GyT",
            persona: persona,
            rubro: clasificar(trf.comercio, reglas),
            notas: trf.notas
          };
          escribirTransaccion(hojaTx, tx);
          verificarAtipico(ss, tx);
          verificarAlertaPresupuesto(ss, tx);
          procesadas++;
        }

        props.setProperty("trf_" + msgId, "1");
      } catch (e) {
        console.error("Error procesando transferencia: " + e.message);
      }
    }
    hilo.addLabel(etiqueta);
  }

  console.log(`Transferencias: ${procesadas} a terceros, ${internas} internas`);
}

/**
 * Parser del correo de transferencia ACH de GyT.
 *
 * Formato observado (jul 2026):
 *   Banco G&T Continental te informa: tu transferencia interbancaria ACH ha sido enviada...
 *   Origen de transferencia ACH
 *     Banco: BANCO G&T CONTINENTAL Q.
 *     Cuenta: 05800102741
 *     Nombre: JOSE ROBERTO...
 *   Destino de transferencia ACH
 *     Banco: BANCO INDUSTRIAL Q.
 *     Cuenta: 1850066539
 *     Tipo Cuenta: MONETARIA
 *     Nombre: ...
 *   Descripción: Carton enviado
 *   Monto: 90.00
 *   Moneda: GTQ
 *   No de confirmación: 11274965
 *   Fecha: 06/07/2026
 *
 * Ojo: los acentos pueden llegar corruptos ("Descripci?n"), por eso los
 * regex aceptan ó, o, ? o cualquier carácter en esa posición.
 */
function parseEmailTransferenciaGyT(cuerpo, fechaCorreo) {
  if (!cuerpo.match(/transferencia\s+interbancaria\s+ACH/i)) return null;

  // Cuentas: la primera es origen, la segunda destino.
  const cuentas = [];
  const reCuenta = /Cuenta:\s*([0-9]{4,})/gi;
  let m;
  while ((m = reCuenta.exec(cuerpo)) !== null) cuentas.push(m[1]);
  if (cuentas.length < 2) {
    console.log("Transferencia ACH sin cuentas identificables. Primeros 300 chars:", cuerpo.substring(0, 300));
    return null;
  }
  const cuentaOrigen = cuentas[0];
  const cuentaDestino = cuentas[1];

  // Nombres: el último "Nombre:" corresponde al destino.
  const nombres = [];
  const reNombre = /Nombre:\s*([^\n\r]+)/gi;
  while ((m = reNombre.exec(cuerpo)) !== null) nombres.push(m[1].trim());
  const nombreDestino = nombres.length >= 2 ? nombres[nombres.length - 1] : "";

  // Bancos destino (segunda aparición).
  const bancos = [];
  const reBanco = /Banco:\s*([^\n\r]+)/gi;
  while ((m = reBanco.exec(cuerpo)) !== null) bancos.push(m[1].trim());
  const bancoDestino = bancos.length >= 2 ? bancos[1] : "";

  const montoMatch = cuerpo.match(/Monto:\s*([\d,]+\.?\d*)/i);
  if (!montoMatch) return null;
  const monto = parseFloat(montoMatch[1].replace(/,/g, ""));
  if (!monto || isNaN(monto)) return null;

  const monedaMatch = cuerpo.match(/Moneda:\s*(GTQ|USD)/i);
  const moneda = monedaMatch ? monedaMatch[1].toUpperCase() : "GTQ";

  const descMatch = cuerpo.match(/Descripci.n:\s*([^\n\r]+)/i);
  const descripcion = descMatch ? descMatch[1].trim() : "";

  const confMatch = cuerpo.match(/confirmaci.n:\s*(\d+)/i);
  const confirmacion = confMatch ? confMatch[1] : "";

  // "Comercio" para clasificar: la descripción es lo más informativo;
  // si viene vacía, se usa el nombre del destinatario.
  const comercio = (descripcion || nombreDestino || "Transferencia ACH")
    .replace(/\s+/g, " ").trim();

  let notas = `Cta •••${cuentaOrigen.slice(-4)} → ${nombreDestino || bancoDestino} •••${cuentaDestino.slice(-4)}`;
  if (confirmacion) notas += ` | Conf. ${confirmacion}`;

  return {
    fecha: fechaCorreo,
    hora: Utilities.formatDate(fechaCorreo, "GMT-6", "HH:mm"),
    monto: monto,
    moneda: moneda,
    comercio: comercio,
    cuentaOrigen: cuentaOrigen,
    cuentaDestino: cuentaDestino,
    nombreDestino: nombreDestino,
    bancoDestino: bancoDestino,
    descripcion: descripcion,
    confirmacion: confirmacion,
    notas: notas
  };
}

/**
 * Lee la hoja "Cuentas Propias" (números de cuenta tuyos y de tu esposa).
 * Todo movimiento hacia estas cuentas se considera interno y NO es gasto.
 */
function leerCuentasPropias(ss) {
  const hoja = ss.getSheetByName(CONFIG_TRF.HOJA_CUENTAS_PROPIAS);
  if (!hoja || hoja.getLastRow() < 2) return [];
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues()
    .map(r => r[0].toString().trim())
    .filter(c => c.length > 0);
}

function registrarTransferenciaInterna(ss, trf, persona) {
  let hoja = ss.getSheetByName(CONFIG_TRF.HOJA_TRANSF_INTERNAS);
  if (!hoja) {
    hoja = ss.insertSheet(CONFIG_TRF.HOJA_TRANSF_INTERNAS);
    hoja.appendRow(["Fecha", "Hora", "Monto", "Moneda", "Origen", "Destino", "Descripción", "Confirmación", "Persona"]);
    hoja.getRange(1, 1, 1, 9).setFontWeight("bold");
  }
  hoja.appendRow([
    trf.fecha, trf.hora, trf.monto, trf.moneda,
    "•••" + trf.cuentaOrigen.slice(-4),
    `${trf.nombreDestino} •••${trf.cuentaDestino.slice(-4)}`,
    trf.descripcion, trf.confirmacion, persona
  ]);
}

/**
 * Prueba del parser con los dos correos reales (GTQ y USD).
 * Ejecutar desde el editor para verificar sin tocar Gmail.
 */
function probarParserTransferencia() {
  const ejemploGTQ = `Estimado Usuario:

Banco G&T Continental te informa: tu transferencia interbancaria ACH ha sido enviada con la siguiente información:

Origen de transferencia ACH

Banco: BANCO G&T CONTINENTAL Q.
Cuenta: 05800102741
Nombre: JOSE ROBERTO CASTA?E
Destino de transferencia ACH

Banco: BANCO INDUSTRIAL Q.
Cuenta: 1850066539
Tipo Cuenta: MONETARIA
Nombre: BANCO INDUSTRIAL Q.
Descripción: Carton enviado
Monto: 90.00
Moneda: GTQ
No de confirmación: 11274965
Fecha: 06/07/2026
Si no realizaste esta acción comunícate al 1718`;

  const r = parseEmailTransferenciaGyT(ejemploGTQ, new Date());
  console.log(JSON.stringify(r, null, 2));
  console.log("Monto = 90:", r.monto === 90);
  console.log("Moneda = GTQ:", r.moneda === "GTQ");
  console.log("Origen = 05800102741:", r.cuentaOrigen === "05800102741");
  console.log("Destino = 1850066539:", r.cuentaDestino === "1850066539");
  console.log("Comercio = Carton enviado:", r.comercio === "Carton enviado");
}

// ╔══════════════════════════════════════════════════════════════╗
// ║ MetodosPago                                                        
// ╚══════════════════════════════════════════════════════════════╝

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

// ╔══════════════════════════════════════════════════════════════╗
// ║ ControlMaestro                                                        
// ╚══════════════════════════════════════════════════════════════╝

/**
 * MÓDULO v3 — CONTROL MAESTRO (GASTOS FIJOS)
 * Archivo NUEVO: no modifica nada del sistema v2.
 *
 * Trabaja sobre el spreadsheet "Flujo de Caja Familiar - Control Maestro 2026"
 * (separado del sheet de seguimiento). Crea dos hojas nuevas SIN tocar las
 * existentes:
 *
 *   "Fijos"          → catálogo de ingresos/egresos fijos (editable).
 *                      Se siembra una sola vez desde la hoja "Rubros".
 *   "Control Pagos"  → checklist mensual: una fila por fijo por mes,
 *                      con checkbox de Pagado/Cobrado y fecha.
 *
 * Los fijos van por MES CALENDARIO (la renta se paga "en julio"), a
 * diferencia de los variables que van por período de corte 10→9.
 *
 * Trigger requerido (setupV3 lo crea automáticamente):
 *   generarPagosDelMes | Por tiempo | Mensual | Día 1
 */

const CONFIG_CM = {
  SPREADSHEET_ID: "1Sd-GU6buTU_zBL7cqX_6nkfr23ifcnhRqIXvvNH5nZo",
  HOJA_RUBROS: "Rubros",       // hoja existente de la que se siembra el catálogo
  HOJA_FIJOS: "Fijos",
  HOJA_PAGOS: "Control Pagos"
};

function abrirControlMaestro() {
  return SpreadsheetApp.openById(CONFIG_CM.SPREADSHEET_ID);
}

/**
 * Crea el catálogo "Fijos" sembrado desde la hoja "Rubros" del Control
 * Maestro (solo filas con Naturaleza = "Fijo"). Si la hoja ya existe,
 * no hace nada — es seguro ejecutarlo varias veces.
 */
function crearCatalogoFijos() {
  const cm = abrirControlMaestro();
  if (cm.getSheetByName(CONFIG_CM.HOJA_FIJOS)) {
    console.log("La hoja 'Fijos' ya existe; no se modifica.");
    return;
  }

  const rubros = cm.getSheetByName(CONFIG_CM.HOJA_RUBROS);
  if (!rubros) throw new Error("No se encontró la hoja 'Rubros' en el Control Maestro.");

  // Estructura de Rubros: #, Tipo flujo, Naturaleza, Categoría/Rubro, Detalle, Monto mensual GTQ, Notas
  const datos = rubros.getDataRange().getValues();
  const fijos = [];
  for (const r of datos) {
    const id = parseInt(r[0]);
    if (isNaN(id)) continue; // encabezados u otras filas
    const tipo = (r[1] || "").toString().trim();       // Ingreso / Egreso
    const naturaleza = (r[2] || "").toString().trim(); // Fijo / Variable
    if (naturaleza.toLowerCase() !== "fijo") continue;
    fijos.push([
      id,
      tipo,
      (r[3] || "").toString().trim(),  // Nombre (Categoría/Rubro)
      (r[4] || "").toString().trim(),  // Detalle
      typeof r[5] === "number" ? r[5] : 0, // Monto mensual GTQ
      "",                              // Día del mes (opcional, editable)
      "Sí",                            // Activo
      (r[6] || "").toString().trim()   // Notas
    ]);
  }

  const hoja = cm.insertSheet(CONFIG_CM.HOJA_FIJOS);
  const encabezado = ["ID", "Tipo", "Nombre", "Detalle", "Monto GTQ", "Día de pago", "Activo", "Notas"];
  hoja.getRange(1, 1, 1, encabezado.length).setValues([encabezado]).setFontWeight("bold").setBackground("#f0f0f0");
  if (fijos.length > 0) {
    hoja.getRange(2, 1, fijos.length, encabezado.length).setValues(fijos);
    hoja.getRange(2, 5, fijos.length, 1).setNumberFormat("Q#,##0.00");
  }
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, encabezado.length);
  console.log(`Catálogo 'Fijos' creado con ${fijos.length} rubros fijos.`);
}

/**
 * Genera las filas del checklist del mes actual en "Control Pagos".
 * Idempotente: si las filas del mes ya existen, no duplica.
 * Se ejecuta automáticamente el día 1 de cada mes (y desde setupV3).
 */
function generarPagosDelMes() {
  const cm = abrirControlMaestro();
  const hojaFijos = cm.getSheetByName(CONFIG_CM.HOJA_FIJOS);
  if (!hojaFijos) throw new Error("Ejecuta primero crearCatalogoFijos().");

  let hoja = cm.getSheetByName(CONFIG_CM.HOJA_PAGOS);
  if (!hoja) {
    hoja = cm.insertSheet(CONFIG_CM.HOJA_PAGOS);
    const enc = ["Año", "Mes", "ID", "Tipo", "Nombre", "Monto GTQ", "Pagado", "Fecha pago", "Notas"];
    hoja.getRange(1, 1, 1, enc.length).setValues([enc]).setFontWeight("bold").setBackground("#f0f0f0");
    hoja.setFrozenRows(1);
  }

  const ahora = new Date();
  const anio = ahora.getFullYear();
  const mes = ahora.getMonth() + 1; // 1-12

  // Filas existentes del mes (para no duplicar).
  const existentes = new Set();
  if (hoja.getLastRow() > 1) {
    const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 3).getValues();
    for (const [a, m, id] of datos) {
      if (a === anio && m === mes) existentes.add(id.toString());
    }
  }

  // Fijos activos del catálogo.
  const fijos = hojaFijos.getRange(2, 1, Math.max(hojaFijos.getLastRow() - 1, 1), 8).getValues();
  const nuevas = [];
  for (const f of fijos) {
    const [id, tipo, nombre, , monto, , activo] = f;
    if (!id || (activo || "").toString().toLowerCase().startsWith("n")) continue;
    if (existentes.has(id.toString())) continue;
    nuevas.push([anio, mes, id, tipo, nombre, monto, false, "", ""]);
  }

  if (nuevas.length > 0) {
    const filaInicio = hoja.getLastRow() + 1;
    hoja.getRange(filaInicio, 1, nuevas.length, 9).setValues(nuevas);
    hoja.getRange(filaInicio, 7, nuevas.length, 1).insertCheckboxes();
    hoja.getRange(filaInicio, 6, nuevas.length, 1).setNumberFormat("Q#,##0.00");
  }
  console.log(`Control Pagos ${mes}/${anio}: ${nuevas.length} filas nuevas.`);
}

/**
 * Marca (o desmarca) un pago del checklist. Usado por la API de la app.
 * Devuelve true si encontró y actualizó la fila.
 */
function marcarPagoFijo(anio, mes, id, pagado) {
  const cm = abrirControlMaestro();
  const hoja = cm.getSheetByName(CONFIG_CM.HOJA_PAGOS);
  if (!hoja || hoja.getLastRow() < 2) return false;

  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 3).getValues();
  for (let i = 0; i < datos.length; i++) {
    if (datos[i][0] === anio && datos[i][1] === mes && datos[i][2].toString() === id.toString()) {
      const fila = i + 2;
      hoja.getRange(fila, 7).setValue(pagado === true);
      hoja.getRange(fila, 8).setValue(pagado === true ? new Date() : "");
      return true;
    }
  }
  return false;
}

/**
 * Lee el checklist de un mes para la API.
 */
function leerPagosDelMes(anio, mes) {
  const cm = abrirControlMaestro();
  const hoja = cm.getSheetByName(CONFIG_CM.HOJA_PAGOS);
  if (!hoja || hoja.getLastRow() < 2) return [];

  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 9).getValues();
  return datos
    .filter(r => r[0] === anio && r[1] === mes)
    .map(r => ({
      id: r[2],
      tipo: r[3],
      nombre: r[4],
      monto: typeof r[5] === "number" ? r[5] : 0,
      pagado: r[6] === true,
      fechaPago: r[7] instanceof Date ? Utilities.formatDate(r[7], "GMT-6", "yyyy-MM-dd") : ""
    }));
}

// ╔══════════════════════════════════════════════════════════════╗
// ║ API                                                        
// ╚══════════════════════════════════════════════════════════════╝

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

// ╔══════════════════════════════════════════════════════════════╗
// ║ Setup                                                        
// ╚══════════════════════════════════════════════════════════════╝

/**
 * MÓDULO v3 — SETUP DE UN SOLO CLIC
 * Archivo NUEVO: no modifica nada del sistema v2.
 *
 * Ejecutar setupV3() UNA VEZ desde el editor. Hace todo lo necesario:
 *   1. Crea la hoja "Cuentas Propias" (con tus dos cuentas GyT sembradas).
 *   2. Crea la hoja "Transferencias Internas".
 *   3. Crea el catálogo "Fijos" y el checklist "Control Pagos" en el
 *      Control Maestro, sembrados desde tu hoja "Rubros".
 *   4. Genera el checklist del mes actual.
 *   5. Construye la hoja "Métodos de Pago".
 *   6. Genera el token de la API (queda en los logs — cópialo).
 *   7. Crea los triggers nuevos (sin tocar los del sistema v2):
 *        procesarTransferencias  → cada 15 minutos
 *        actualizarMetodosPago   → cada hora
 *        generarPagosDelMes      → día 1 de cada mes
 *
 * Es idempotente: se puede volver a ejecutar sin duplicar nada.
 */

function setupV3() {
  crearHojaCuentasPropias();
  crearCatalogoFijos();
  generarPagosDelMes();
  actualizarMetodosPago();
  const token = obtenerOCrearToken();
  crearTriggersV3();

  console.log("========================================");
  console.log("SETUP v3 COMPLETO ✓");
  console.log("TOKEN DE LA API (cópialo para la app):");
  console.log(token);
  console.log("Siguiente paso: Implementar > Nueva implementación > Aplicación web");
  console.log("  Ejecutar como: Yo | Acceso: Cualquier usuario");
  console.log("========================================");
}

function crearHojaCuentasPropias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(CONFIG_TRF.HOJA_CUENTAS_PROPIAS)) return;

  const hoja = ss.insertSheet(CONFIG_TRF.HOJA_CUENTAS_PROPIAS);
  hoja.getRange(1, 1, 1, 2).setValues([["Número de cuenta", "Descripción"]])
    .setFontWeight("bold").setBackground("#f0f0f0");
  hoja.getRange(2, 1, 4, 2).setValues([
    ["05800102741", "GyT Monetaria Q (JR)"],
    ["05858062721", "GyT Monetaria $ (JR)"],
    ["", "← agregar aquí cuentas BI, Banrural, de tu esposa, etc."],
    ["", "Toda transferencia HACIA estas cuentas se ignora como gasto"]
  ]);
  // Formato de texto para que no se coman los ceros iniciales.
  hoja.getRange("A:A").setNumberFormat("@");
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, 2);
  console.log("Hoja 'Cuentas Propias' creada. REVISA y agrega las que falten.");
}

function obtenerOCrearToken() {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty("API_TOKEN");
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, "");
    props.setProperty("API_TOKEN", token);
  }
  return token;
}

/** Muestra el token en los logs sin regenerarlo (por si lo perdiste). */
function mostrarToken() {
  console.log("TOKEN: " + (PropertiesService.getScriptProperties().getProperty("API_TOKEN") || "(no generado aún — ejecuta setupV3)"));
}

function crearTriggersV3() {
  const deseados = {
    "procesarTransferencias": t => t.timeBased().everyMinutes(15),
    "actualizarMetodosPago": t => t.timeBased().everyHours(1),
    "generarPagosDelMes": t => t.timeBased().onMonthDay(1).atHour(6)
  };

  const existentes = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  for (const [fn, builder] of Object.entries(deseados)) {
    if (existentes.includes(fn)) continue;
    builder(ScriptApp.newTrigger(fn)).create();
    console.log("Trigger creado: " + fn);
  }
}
