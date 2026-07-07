/**
 * SISTEMA DE FINANZAS FAMILIARES — v2
 * Procesa automáticamente correos de GyT (G&T Continental) y Banrural
 * y registra los gastos en Google Sheets.
 *
 * AJUSTES v2 (mayo 2026):
 * - Parser de GyT calibrado al formato real del banco (remitente bpi@gtc.com.gt)
 * - Extrae cuenta (últimos 4 dígitos), número de autorización, monto, moneda y localidad
 * - Limpia el sufijo "GT" del nombre del comercio
 * - Detecta crédito vs débito desde el cuerpo del correo
 *
 * INSTRUCCIONES DE USO:
 * 1. Abrir Google Sheets > Extensiones > Apps Script
 * 2. Borrar el código de ejemplo y pegar TODO este archivo
 * 3. Guardar (Ctrl+S) con el nombre "Finanzas Familiares"
 * 4. Ir a "Activadores" (icono de reloj) y crear:
 *    - Función: procesarCorreos | Evento: Por tiempo | Cada 15 minutos
 *    - Función: enviarResumenSemanal | Evento: Por tiempo | Semanal | Lunes 7am
 * 5. Ejecutar manualmente "probarSistema" la primera vez para autorizar permisos
 */

// ============================================================
// CONFIGURACIÓN
// ============================================================
const CONFIG = {
  // Remitentes de los correos de los bancos (CONFIRMADOS).
  REMITENTE_GYT: "bpi@gtc.com.gt",
  REMITENTE_BANRURAL: "notificaciones@banrural.com.gt", // <-- AJUSTAR cuando confirmemos

  // Etiqueta de Gmail que se asigna a los correos ya procesados.
  ETIQUETA_PROCESADO: "Finanzas/Procesado",

  // Nombre de las hojas en el Google Sheet.
  HOJA_TRANSACCIONES: "Transacciones",
  HOJA_REGLAS: "Reglas",
  HOJA_PRESUPUESTO: "Presupuesto",
  HOJA_DASHBOARD: "Dashboard Mensual",
  HOJA_ATIPICOS: "Gastos atípicos",
  HOJA_CONFIG: "Configuración",

  // Si un correo viene reenviado, se asume que es de tu esposa.
  PREFIJO_REENVIO: ["Fwd:", "Fw:", "RV:", "Reenviar:"],

  // Email destinatario para alertas y resumen semanal.
  EMAIL_RESUMEN: "",
  EMAIL_RESUMEN_ESPOSA: "",

  // Procesar correos de los últimos N días (para evitar buscar todo el historial).
  DIAS_BUSQUEDA: 3,

  // Día del mes en que inicia tu "mes financiero".
  // Tu ciclo es del 10 al 9 del mes siguiente (corte de tarjeta el 9).
  DIA_INICIO_MES: 10
};

// ============================================================
// FUNCIÓN PRINCIPAL: SE EJECUTA CADA 15 MINUTOS
// ============================================================
function procesarCorreos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTx = ss.getSheetByName(CONFIG.HOJA_TRANSACCIONES);
  const reglas = leerReglas(ss);
  const etiqueta = obtenerOCrearEtiqueta(CONFIG.ETIQUETA_PROCESADO);

  // Usar PropertiesService para rastrear IDs de mensajes ya procesados.
  const props = PropertiesService.getDocumentProperties();

  // Buscar correos no procesados de los bancos.
  const fechaDesde = new Date();
  fechaDesde.setDate(fechaDesde.getDate() - CONFIG.DIAS_BUSQUEDA);
  const fechaStr = Utilities.formatDate(fechaDesde, "GMT-6", "yyyy/MM/dd");

  const query = `(from:${CONFIG.REMITENTE_GYT} OR from:${CONFIG.REMITENTE_BANRURAL}) ` +
                `after:${fechaStr} -label:${CONFIG.ETIQUETA_PROCESADO}`;

  const hilos = GmailApp.search(query, 0, 100);
  let procesados = 0;

  for (const hilo of hilos) {
    const mensajes = hilo.getMessages();
    for (const msg of mensajes) {
      try {
        // Verificar si este mensaje específico ya fue procesado por su ID único.
        const msgId = msg.getId();
        const yaProcessado = props.getProperty("msg_" + msgId);
        if (yaProcessado) continue;

        const tx = parsearCorreo(msg);
        if (tx) {
          tx.rubro = clasificar(tx.comercio, reglas);
          escribirTransaccion(hojaTx, tx);
          verificarAtipico(ss, tx);
          verificarAlertaPresupuesto(ss, tx);
          // Marcar este mensaje individual como procesado.
          props.setProperty("msg_" + msgId, "1");
          procesados++;
        }
      } catch (e) {
        console.error("Error procesando mensaje: " + e.message);
      }
    }
    hilo.addLabel(etiqueta);
  }

  console.log(`Procesados: ${procesados} mensajes`);
}

// ============================================================
// PARSEO DE CORREOS
// ============================================================
function parsearCorreo(msg) {
  const remitente = msg.getFrom().toLowerCase();
  const asunto = msg.getSubject();
  const cuerpo = msg.getPlainBody();

  // Determinar persona (Jr o Ms) por si el correo viene reenviado.
  const esReenvio = CONFIG.PREFIJO_REENVIO.some(p => asunto.startsWith(p));
  const persona = esReenvio ? "Ms" : "Jr";

  let datos = null;
  if (remitente.includes("gtc.com.gt") || remitente.includes("bpi@gtc")) {
    datos = parseEmailGyT(cuerpo, asunto, msg.getDate());
    if (datos) datos.banco = "GyT";
  } else if (remitente.includes("banrural")) {
    datos = parseEmailBanrural(cuerpo, asunto, msg.getDate());
    if (datos) datos.banco = "Banrural";
  }

  if (!datos) return null;

  return {
    fecha: datos.fecha,
    hora: datos.hora,
    monto: datos.monto,
    moneda: datos.moneda,
    comercio: datos.comercio,
    metodo: datos.metodo,
    banco: datos.banco,
    persona: persona,
    notas: datos.notas || ""
  };
}

/**
 * Parser de GyT — calibrado al formato real (mayo 2026).
 *
 * Formato observado:
 *   Estimado Usuario,
 *   Banco G&T Continental, te informa:
 *   Consumo tarjeta crédito con la cuenta 4447 No. autorización: 00056410 Monto: Q. 30.00
 *   Localidad: PARQUEOS DE ANTIGUA GUATEMALA GT
 */
function parseEmailGyT(cuerpo, asunto, fechaCorreo) {
  // Solo procesa correos que sean realmente de consumo (no avisos generales).
  if (!cuerpo.match(/Consumo\s+tarjeta/i)) {
    console.log("Correo de GyT ignorado (no es notificación de consumo)");
    return null;
  }

  // 1. Detectar método: "Consumo tarjeta crédito" o "Consumo tarjeta débito".
  let metodo = "TC";
  const metodoMatch = cuerpo.match(/Consumo\s+tarjeta\s+(cr[ée]dito|d[ée]bito)/i);
  if (metodoMatch) {
    metodo = metodoMatch[1].toLowerCase().startsWith("d") ? "Débito" : "TC";
  }

  // 2. Extraer últimos 4 dígitos de la cuenta.
  let cuentaUlt4 = "";
  const cuentaMatch = cuerpo.match(/cuenta\s+(\d{4})/i);
  if (cuentaMatch) cuentaUlt4 = cuentaMatch[1];

  // 3. Extraer número de autorización.
  let autorizacion = "";
  const autMatch = cuerpo.match(/autorizaci[óo]n[:\s]+(\d+)/i);
  if (autMatch) autorizacion = autMatch[1];

  // 4. Extraer monto y moneda.
  // Formatos observados en correos reales de GyT:
  //   "Monto: Q. 30.00"  "Monto: Q 30.00"  "Monto: Q30.00"
  //   "Monto: $. 10.00"  "Monto: $ 10.00"  "Monto: $10.00"  "Monto: USD 10.00"
  // Nota: el cuerpo puede llegar en una sola línea larga (texto plano de HTML).
  let monto = null;
  let moneda = "GTQ";

  const montoMatch = cuerpo.match(
    /Monto[:\s]+(Q\.?|USD|\$\.?|GTQ)\.?\s*([\d,]+\.?\d*)/i
  );
  if (montoMatch) {
    const simbolo = montoMatch[1].replace(".", "").toUpperCase();
    monto = parseFloat(montoMatch[2].replace(/,/g, ""));
    if (simbolo === "USD" || simbolo === "$") moneda = "USD";
    else moneda = "GTQ";
  }

  if (!monto || isNaN(monto)) {
    console.log("No se pudo extraer monto de correo GyT. Primeros 300 chars:", cuerpo.substring(0, 300));
    return null;
  }

  // 5. Extraer comercio (campo "Localidad:") y limpiar código de país al final.
  // El cuerpo puede llegar en una sola línea — buscar hasta doble espacio o fin de línea.
  let comercio = "GyT - Desconocido";
  const comercioMatch = cuerpo.match(/Localidad[:\s]+(.+?)(?:\s{2,}|Sitio web|$)/i);
  if (comercioMatch) {
    comercio = comercioMatch[1].trim();
    // Limpiar código de país de 2 letras al final (GT, US, MX, etc.)
    comercio = comercio.replace(/\s+[A-Z]{2}\s*$/i, "");
    // Limpiar espacios múltiples.
    comercio = comercio.replace(/\s+/g, " ").trim();
  }

  // 6. Construir nota con detalles útiles para auditoría.
  let notas = "";
  if (cuentaUlt4) notas += `Tarjeta •••${cuentaUlt4}`;
  if (autorizacion) notas += (notas ? " | " : "") + `Aut. ${autorizacion}`;

  return {
    monto: monto,
    moneda: moneda,
    comercio: comercio,
    fecha: fechaCorreo,
    hora: Utilities.formatDate(fechaCorreo, "GMT-6", "HH:mm"),
    metodo: metodo,
    notas: notas
  };
}

/**
 * Parser de Banrural — placeholder (ajustar cuando confirmemos formato real).
 */
function parseEmailBanrural(cuerpo, asunto, fechaCorreo) {
  const montoMatch = cuerpo.match(/(?:Monto|MONTO|Total)[:\s]+(?:Q|GTQ|USD|\$)?\s*([\d,]+\.?\d*)/i);
  if (!montoMatch) return null;

  const monedaMatch = cuerpo.match(/(USD|GTQ|Q\.?|\$)/);
  const comercioMatch = cuerpo.match(/(?:Comercio|Establecimiento|Lugar|Localidad|En)[:\s]+([^\n\r]+)/i);

  const monto = parseFloat(montoMatch[1].replace(/,/g, ""));
  let moneda = "GTQ";
  if (monedaMatch) {
    const m = monedaMatch[1].toUpperCase();
    if (m === "USD" || m === "$") moneda = "USD";
  }

  let comercio = comercioMatch ? comercioMatch[1].trim() : "Banrural - Desconocido";
  comercio = comercio.replace(/\s+GT\s*$/i, "").replace(/\s+/g, " ").trim();

  return {
    monto: monto,
    moneda: moneda,
    comercio: comercio,
    fecha: fechaCorreo,
    hora: Utilities.formatDate(fechaCorreo, "GMT-6", "HH:mm"),
    metodo: cuerpo.toLowerCase().includes("debito") || cuerpo.toLowerCase().includes("débito") ? "Débito" : "TC",
    notas: ""
  };
}

// ============================================================
// CLASIFICACIÓN
// ============================================================
function leerReglas(ss) {
  const hoja = ss.getSheetByName(CONFIG.HOJA_REGLAS);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 5) return [];

  const datos = hoja.getRange(5, 1, ultimaFila - 4, 2).getValues();
  return datos
    .filter(r => r[0] && r[1])
    .map(r => ({ clave: r[0].toString().toUpperCase(), rubro: r[1].toString() }));
}

function clasificar(comercio, reglas) {
  const c = comercio.toUpperCase();
  for (const regla of reglas) {
    if (c.includes(regla.clave)) return regla.rubro;
  }
  return "Otros";
}

// ============================================================
// ESCRITURA EN HOJA
// ============================================================
function escribirTransaccion(hoja, tx) {
  const fila = [
    tx.fecha,
    tx.hora,
    tx.monto,
    tx.moneda,
    tx.comercio,
    tx.metodo,
    tx.banco,
    tx.persona,
    tx.rubro,
    tx.notas
  ];
  hoja.appendRow(fila);
}

// ============================================================
// DETECTOR DE GASTOS ATÍPICOS
// ============================================================
function verificarAtipico(ss, tx) {
  const hojaTx = ss.getSheetByName(CONFIG.HOJA_TRANSACCIONES);
  const hojaConfig = ss.getSheetByName(CONFIG.HOJA_CONFIG);
  const umbral = hojaConfig.getRange("B11").getValue();

  const ultimaFila = hojaTx.getLastRow();
  if (ultimaFila < 5) return;

  const datos = hojaTx.getRange(5, 1, ultimaFila - 4, 10).getValues();
  const mismoRubro = datos.filter(r => r[8] === tx.rubro && r[2] > 0);
  if (mismoRubro.length < 3) return;

  const tc = obtenerTipoCambio(ss);
  const montos = mismoRubro.map(r => r[3] === "USD" ? r[2] * tc : r[2]);
  const promedio = montos.reduce((a, b) => a + b, 0) / montos.length;
  const montoGTQ = tx.moneda === "USD" ? tx.monto * tc : tx.monto;

  if (montoGTQ > promedio * umbral) {
    const hojaAtip = ss.getSheetByName(CONFIG.HOJA_ATIPICOS);
    hojaAtip.appendRow([
      tx.fecha,
      tx.comercio,
      tx.rubro,
      montoGTQ,
      promedio,
      (montoGTQ / promedio).toFixed(2) + "x"
    ]);

    enviarAlertaAtipico(tx, montoGTQ, promedio);
  }
}

function enviarAlertaAtipico(tx, monto, promedio) {
  const destinatario = CONFIG.EMAIL_RESUMEN || Session.getActiveUser().getEmail();
  const asunto = `⚠ Gasto atípico: ${tx.comercio} (${tx.rubro})`;
  const cuerpo = `
Se detectó un gasto inusualmente alto:

  Comercio: ${tx.comercio}
  Rubro: ${tx.rubro}
  Monto: Q${monto.toFixed(2)}
  Promedio histórico del rubro: Q${promedio.toFixed(2)}
  Veces sobre el promedio: ${(monto / promedio).toFixed(2)}x

Si este gasto es correcto, ignora este correo. Si no, revisa la transacción
en el Google Sheet de finanzas familiares.
`;
  GmailApp.sendEmail(destinatario, asunto, cuerpo);
}

function obtenerTipoCambio(ss) {
  return ss.getSheetByName(CONFIG.HOJA_CONFIG).getRange("B3").getValue();
}

/**
 * Calcula el inicio y fin del período financiero actual.
 * El mes financiero va del día DIA_INICIO_MES al día (DIA_INICIO_MES - 1) del mes siguiente.
 * Ejemplo con DIA_INICIO_MES = 10:
 *   Si hoy es 7 de mayo  → período: 10 abril – 9 mayo
 *   Si hoy es 15 de mayo → período: 10 mayo  – 9 junio
 */
function obtenerPeriodoFinanciero() {
  const hoy = new Date();
  const dia = hoy.getDate();
  const mes = hoy.getMonth();
  const anio = hoy.getFullYear();
  const inicio_dia = CONFIG.DIA_INICIO_MES;

  let inicioMes, finMes;

  if (dia >= inicio_dia) {
    // Estamos en la segunda mitad del período: inició este mes
    inicioMes = new Date(anio, mes, inicio_dia, 0, 0, 0);
    finMes    = new Date(anio, mes + 1, inicio_dia - 1, 23, 59, 59);
  } else {
    // Estamos en la primera mitad: el período inició el mes pasado
    inicioMes = new Date(anio, mes - 1, inicio_dia, 0, 0, 0);
    finMes    = new Date(anio, mes, inicio_dia - 1, 23, 59, 59);
  }

  return { inicio: inicioMes, fin: finMes };
}

// ============================================================
// VERIFICACIÓN DE ALERTAS DE PRESUPUESTO
// ============================================================
function verificarAlertaPresupuesto(ss, tx) {
  const hojaConfig = ss.getSheetByName(CONFIG.HOJA_CONFIG);
  const umbral = hojaConfig.getRange("B10").getValue();

  const hojaTx = ss.getSheetByName(CONFIG.HOJA_TRANSACCIONES);
  const ultimaFila = hojaTx.getLastRow();
  if (ultimaFila < 5) return;

  const datos = hojaTx.getRange(5, 1, ultimaFila - 4, 10).getValues();
  const tc = obtenerTipoCambio(ss);
  const periodo = obtenerPeriodoFinanciero();

  const gastoMesEnRubro = datos
    .filter(r => {
      if (!r[0] || !(r[0] instanceof Date)) return false;
      return r[8] === tx.rubro &&
             r[0] >= periodo.inicio &&
             r[0] <= periodo.fin;
    })
    .reduce((sum, r) => sum + (r[3] === "USD" ? r[2] * tc : r[2]), 0);

  const hojaPres = ss.getSheetByName(CONFIG.HOJA_PRESUPUESTO);
  const presupuestos = hojaPres.getRange(4, 1, hojaPres.getLastRow() - 3, 4).getValues();
  const filaPres = presupuestos.find(r => r[1] === tx.rubro);
  if (!filaPres) return;

  const presupuestoGTQ = filaPres[2] === "USD" ? filaPres[3] * tc : filaPres[3];
  if (presupuestoGTQ === 0) return;

  const porcentaje = gastoMesEnRubro / presupuestoGTQ;

  const propiedades = PropertiesService.getDocumentProperties();
  const llaveAlerta = `alerta_${tx.rubro}_${Utilities.formatDate(periodo.inicio, "GMT-6", "yyyyMM")}`;
  const yaAlerto = propiedades.getProperty(llaveAlerta);

  if (porcentaje >= 1 && yaAlerto !== "excedido") {
    enviarAlertaPresupuesto(tx.rubro, "EXCEDIDO", gastoMesEnRubro, presupuestoGTQ);
    propiedades.setProperty(llaveAlerta, "excedido");
  } else if (porcentaje >= umbral && !yaAlerto) {
    enviarAlertaPresupuesto(tx.rubro, "CERCA DEL LÍMITE", gastoMesEnRubro, presupuestoGTQ);
    propiedades.setProperty(llaveAlerta, "umbral");
  }
}

function enviarAlertaPresupuesto(rubro, tipo, gastado, presupuesto) {
  const destinatarios = [CONFIG.EMAIL_RESUMEN || Session.getActiveUser().getEmail()];
  if (CONFIG.EMAIL_RESUMEN_ESPOSA) destinatarios.push(CONFIG.EMAIL_RESUMEN_ESPOSA);

  const asunto = tipo === "EXCEDIDO"
    ? `🚨 ${rubro} excedió el presupuesto`
    : `⚠ ${rubro} cerca del límite`;

  const cuerpo = `
Estado del rubro "${rubro}":

  Gastado este mes: Q${gastado.toFixed(2)}
  Presupuesto: Q${presupuesto.toFixed(2)}
  Porcentaje usado: ${((gastado / presupuesto) * 100).toFixed(1)}%

${tipo === "EXCEDIDO"
  ? "Ya superaste el presupuesto de este rubro este mes."
  : "Estás cerca del límite. Considera revisar antes del próximo gasto."}

Ver dashboard: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}
`;
  GmailApp.sendEmail(destinatarios.join(","), asunto, cuerpo);
}

// ============================================================
// RESUMEN SEMANAL (LUNES POR LA MAÑANA)
// ============================================================
function enviarResumenSemanal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTx = ss.getSheetByName(CONFIG.HOJA_TRANSACCIONES);
  const tc = obtenerTipoCambio(ss);
  const ultimaFila = hojaTx.getLastRow();
  const datos = ultimaFila >= 5
    ? hojaTx.getRange(5, 1, ultimaFila - 4, 10).getValues()
    : [];

  const ahora = new Date();
  const haceUnaSemana = new Date(ahora.getTime() - 7 * 24 * 60 * 60 * 1000);
  const periodo = obtenerPeriodoFinanciero();

  const txSemana = datos.filter(r => r[0] instanceof Date && r[0] >= haceUnaSemana);

  const porRubro = {};
  for (const r of txSemana) {
    const monto = r[3] === "USD" ? r[2] * tc : r[2];
    if (!porRubro[r[8]]) porRubro[r[8]] = { total: 0, count: 0 };
    porRubro[r[8]].total += monto;
    porRubro[r[8]].count++;
  }

  const totalSemana = Object.values(porRubro).reduce((s, v) => s + v.total, 0);

  // Total del período financiero actual (del 10 al 9)
  const txMes = datos.filter(r => r[0] instanceof Date && r[0] >= periodo.inicio && r[0] <= periodo.fin);
  const totalMes = txMes.reduce((s, r) => s + (r[3] === "USD" ? r[2] * tc : r[2]), 0);

  const periodoStr = `${Utilities.formatDate(periodo.inicio, "GMT-6", "d MMM")} – ${Utilities.formatDate(periodo.fin, "GMT-6", "d MMM")}`;

  const ordenado = Object.entries(porRubro).sort((a, b) => b[1].total - a[1].total);
  let detalle = "";
  for (const [rubro, info] of ordenado) {
    detalle += `  ${rubro}: Q${info.total.toFixed(2)} (${info.count} transacciones)\n`;
  }

  if (!detalle) detalle = "  (Sin transacciones esta semana)\n";

  const asunto = `Resumen semanal de finanzas — ${Utilities.formatDate(ahora, "GMT-6", "d 'de' MMMM")}`;
  const cuerpo = `
Buenos días. Aquí está el resumen de la última semana:

GASTOS DE LA SEMANA: Q${totalSemana.toFixed(2)}

Por rubro:
${detalle}

GASTOS DEL PERÍODO (${periodoStr}): Q${totalMes.toFixed(2)}

Ver dashboard completo: ${ss.getUrl()}

— Sistema de finanzas familiares
`;

  const destinatarios = [CONFIG.EMAIL_RESUMEN || Session.getActiveUser().getEmail()];
  if (CONFIG.EMAIL_RESUMEN_ESPOSA) destinatarios.push(CONFIG.EMAIL_RESUMEN_ESPOSA);

  GmailApp.sendEmail(destinatarios.join(","), asunto, cuerpo);
}

// ============================================================
// UTILIDADES
// ============================================================
function obtenerOCrearEtiqueta(nombre) {
  let etiqueta = GmailApp.getUserLabelByName(nombre);
  if (!etiqueta) etiqueta = GmailApp.createLabel(nombre);
  return etiqueta;
}

// ============================================================
// FUNCIONES DE PRUEBA
// ============================================================

/**
 * Función de prueba general: inserta una transacción ficticia y verifica
 * que el sistema funciona end-to-end (sin tocar correos reales).
 */
function probarSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTx = ss.getSheetByName(CONFIG.HOJA_TRANSACCIONES);
  const reglas = leerReglas(ss);

  console.log("Reglas cargadas:", reglas.length);
  console.log("Tipo de cambio:", obtenerTipoCambio(ss));

  const txPrueba = {
    fecha: new Date(),
    hora: "12:00",
    monto: 100,
    moneda: "GTQ",
    comercio: "PRUEBA WALMART",
    metodo: "TC",
    banco: "Prueba",
    persona: "Jr",
    rubro: clasificar("PRUEBA WALMART", reglas),
    notas: "Transacción de prueba — borrar después"
  };

  console.log("Rubro asignado:", txPrueba.rubro);
  escribirTransaccion(hojaTx, txPrueba);
  console.log("Transacción de prueba escrita correctamente");
}

/**
 * Prueba específica del parser de GyT con el ejemplo real recibido.
 * Útil para verificar que el regex funciona antes de procesar correos en vivo.
 */
function probarParserGyT() {
  const cuerpoEjemplo = `Estimado Usuario,

Banco G&T Continental, te informa:

Consumo tarjeta crédito con la cuenta 4447 No. autorización: 00056410 Monto: Q. 30.00
Localidad: PARQUEOS DE ANTIGUA GUATEMALA GT`;

  const resultado = parseEmailGyT(cuerpoEjemplo, "MENSAJE DE ALERTA", new Date());
  console.log("Resultado del parser GyT:");
  console.log(JSON.stringify(resultado, null, 2));

  // Verificaciones esperadas
  console.log("\nVerificaciones:");
  console.log("Monto = 30:", resultado.monto === 30);
  console.log("Moneda = GTQ:", resultado.moneda === "GTQ");
  console.log("Comercio limpio (sin 'GT'):", resultado.comercio === "PARQUEOS DE ANTIGUA GUATEMALA");
  console.log("Método = TC:", resultado.metodo === "TC");
  console.log("Notas tienen tarjeta y autorización:", resultado.notas.includes("4447") && resultado.notas.includes("00056410"));
}

/**
 * Procesa el último correo de GyT para diagnóstico.
 * Útil cuando algo falla y queremos ver qué está pasando.
 */
function diagnosticarUltimoCorreoGyT() {
  const hilos = GmailApp.search(`from:${CONFIG.REMITENTE_GYT}`, 0, 1);
  if (hilos.length === 0) {
    console.log("No hay correos de GyT en Gmail.");
    return;
  }

  const msg = hilos[0].getMessages()[0];
  console.log("Asunto:", msg.getSubject());
  console.log("Remitente:", msg.getFrom());
  console.log("Fecha:", msg.getDate());
  console.log("\n--- CUERPO DEL CORREO ---\n");
  console.log(msg.getPlainBody());
  console.log("\n--- RESULTADO DEL PARSER ---\n");

  const resultado = parseEmailGyT(msg.getPlainBody(), msg.getSubject(), msg.getDate());
  console.log(JSON.stringify(resultado, null, 2));
}