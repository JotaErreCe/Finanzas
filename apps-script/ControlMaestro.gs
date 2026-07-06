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
