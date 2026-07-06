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
