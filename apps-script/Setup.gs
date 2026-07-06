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
