const RESPONSE_STATUS_HEADERS_ = ["VALIDACION_ESTADO", "VALIDACION_DETALLE", "SINCRONIZACION_ESTADO", "SINCRONIZACION_DETALLE", "SINCRONIZACION_FECHA"];

/**
 * Inserta o actualiza el inventario tecnico de un equipo.
 * @param {Object} payload Datos del equipo y metadatos de la solicitud.
 * @returns {GoogleAppsScript.Content.TextOutput} Resultado JSON de la operacion.
 */
function saveTechnicalRecord(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const spreadsheetId = payload.spreadsheet_id || Config.spreadsheetId();
    // Asegura el trigger antes de modificar la hoja tecnica.
    try {
      ensureFormSubmitTrigger_(spreadsheetId);
    } catch (error) {
      // La falta de permisos para crear el trigger no impide guardar el registro tecnico.
    }
    const sheetName = payload.sheet_name || Config.technicalSheetName();
    const sheet = getOrCreateTechnicalSheet_(spreadsheetId, sheetName);
    const technical = payload.technical || {};
    // Valida el identificador antes de consultar o modificar la hoja.
    const recordId = String(payload.record_id || technical.id_equipo || "").trim();
    if (!/^HV-[0-9]{8}-[0-9]{6}-[A-Z0-9]{6}$/.test(recordId)) {
      return jsonResponse({ ok: false, error: "ID_EQUIPO invalido.", received: recordId });
    }

    // Busca coincidencias por ID y por serial para actualizar en lugar de duplicar.
    const serialBios = String(technical.serial_bios || "").trim();
    const existing = findTechnicalRowByIdOrSerial_(sheet, recordId, serialBios);
    if (existing.conflict) {
      return jsonResponse({ ok: false, error: "ID_EQUIPO y SERIAL_BIOS pertenecen a filas diferentes." });
    }

    // Construye la fila con el orden de columnas de TECNICA_EQUIPOS.
    const row = [
      payload.generated_at || new Date(), recordId, technical.link_hoja_vida || "",
      technical.nombre_equipo || "", serialBios, technical.mac_principal || "",
      technical.modelo_equipo || "", technical.discos_particiones || "",
      technical.red_tipo_conexion || "", technical.red_ipv4 || "", technical.red_velocidad || "",
      technical.ram_resumen || "", technical.os_name || "", technical.os_version || "",
      technical.last_boot || "", technical.cpu || "", technical.ram_gb || "",
    ];

    // Actualiza la fila existente o agrega una nueva si no hay coincidencias.
    if (existing.row > 0) {
      sheet.getRange(existing.row, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    return jsonResponse({ ok: true, saved: true, updated: existing.row > 0, sheet: sheetName, recordId });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Obtiene o crea la hoja tecnica y sus encabezados.
 * @param {string} spreadsheetId Identificador del libro.
 * @param {string} sheetName Nombre de la hoja.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Hoja lista para operar.
 */
function getOrCreateTechnicalSheet_(spreadsheetId, sheetName) {
  const spreadsheet = Config.spreadsheetById(spreadsheetId);
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["FECHA_REGISTRO", "ID_EQUIPO", "LINK_A_HOJA_DE_VIDA", "NOMBRE_EQUIPO", "SERIAL_BIOS", "MAC_PRINCIPAL", "MODELO_EQUIPO", "DISCOS_PARTICIONES", "RED_TIPO_CONEXION", "RED_IPV4", "RED_VELOCIDAD", "RAM_RESUMEN", "OS_NAME", "OS_VERSION", "LAST_BOOT", "CPU", "RAM_GB"]);
  }
  return sheet;
}

/**
 * Busca la fila de un equipo por ID_EQUIPO o SERIAL_BIOS.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Hoja tecnica.
 * @param {string} recordId Identificador del equipo.
 * @param {string} serialBios Serial del BIOS.
 * @returns {{row: number, conflict: boolean}} Fila encontrada y estado de conflicto.
 */
function findTechnicalRowByIdOrSerial_(sheet, recordId, serialBios) {
  if (sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return { row: -1, conflict: false };
  const headers = getSheetHeaders_(sheet);
  const idIndex = findIdColumnIndex_(headers);
  const serialIndex = headers.findIndex((header) => normalizeHeader_(header) === "SERIAL_BIOS");
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  let idRow = -1;
  let serialRow = -1;
  data.forEach((row, index) => {
    const rowNumber = index + 2;
    if (idIndex >= 0 && String(row[idIndex] || "").trim() === recordId) idRow = rowNumber;
    if (serialBios && serialIndex >= 0 && String(row[serialIndex] || "").trim() === serialBios) serialRow = rowNumber;
  });
  return idRow > 0 && serialRow > 0 && idRow !== serialRow
    ? { row: -1, conflict: true }
    : { row: idRow > 0 ? idRow : serialRow, conflict: false };
}

/**
 * Localiza la columna identificadora de la hoja tecnica.
 * @param {Array<*>} headers Encabezados de la hoja.
 * @returns {number} Indice de la columna o -1.
 */
function findIdColumnIndex_(headers) {
  const idIndex = headers.findIndex((header) => normalizeHeader_(header) === "ID_EQUIPO");
  if (idIndex >= 0) return idIndex;
  return headers.findIndex((header) => normalizeHeader_(header).includes("SERIAL"));
}

/**
 * Crea el trigger instalable de respuestas si aun no existe.
 * @param {string} spreadsheetId Identificador del libro de respuestas.
 * @returns {void}
 */
function ensureFormSubmitTrigger_(spreadsheetId) {
  if (!spreadsheetId) return;
  const exists = ScriptApp.getProjectTriggers().some((trigger) => {
    if (trigger.getHandlerFunction() !== "onFormSubmit" || trigger.getEventType() !== ScriptApp.EventType.ON_FORM_SUBMIT) return false;
    try { return !trigger.getTriggerSourceId() || trigger.getTriggerSourceId() === spreadsheetId; } catch (error) { return false; }
  });
  if (!exists) ScriptApp.newTrigger("onFormSubmit").forSpreadsheet(spreadsheetId).onFormSubmit().create();
}

/**
 * Crea la hoja de incidencias con sus encabezados base.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet Libro de trabajo.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Hoja de incidencias.
 */
function ensureIncidenciasSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(Config.incidenceSheetName());
  if (!sheet) sheet = spreadsheet.insertSheet(Config.incidenceSheetName());
  if (sheet.getLastRow() === 0) sheet.appendRow(["FECHA_INCIDENTE", "MOTIVO", "ID_EQUIPO", "HOJA_ORIGEN", "FILA_ORIGEN", "RESPUESTA_JSON"]);
  return sheet;
}

/**
 * Registra una respuesta descartada junto con su contexto original.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet Libro de trabajo.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} responseSheet Hoja de respuestas.
 * @param {number} rowNumber Fila afectada.
 * @param {string} reason Motivo del registro.
 * @param {string} idValue Identificador involucrado.
 * @param {Array<*>} headers Encabezados originales.
 * @param {Array<*>} rowValues Valores originales.
 * @returns {void}
 */
function appendIncidencia_(spreadsheet, responseSheet, rowNumber, reason, idValue, headers, rowValues) {
  const payload = {};
  headers.forEach((header, index) => { payload[header] = rowValues[index]; });
  ensureIncidenciasSheet_(spreadsheet).appendRow([nowString_(), reason, idValue, responseSheet.getName(), rowNumber, JSON.stringify(payload)]);
}

/**
 * Rechaza un envio reciente cuando su ID_EQUIPO ya existe.
 * @param {Object} event Evento de formulario.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet Libro de respuestas.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} responseSheet Hoja de respuestas.
 * @returns {Object} Decision de rechazo y metadatos.
 */
function rejectDuplicateSubmissionIfNeeded_(event, spreadsheet, responseSheet) {
  if (!responseSheet || responseSheet.getLastRow() < 2) return { rejected: false };
  const headers = getSheetHeaders_(responseSheet);
  const idIndex = findIdColumnIndex_(headers);
  if (idIndex < 0) return { rejected: false, message: "No existe columna ID_EQUIPO" };
  const submittedRow = event && event.range && event.range.getSheet().getName() === responseSheet.getName() ? event.range.getRow() : responseSheet.getLastRow();
  const submittedValues = responseSheet.getRange(submittedRow, 1, 1, headers.length).getValues()[0];
  const submittedId = String(submittedValues[idIndex] || "").trim();
  if (!submittedId) return { rejected: false, message: "Envio sin ID_EQUIPO" };
  const ids = responseSheet.getRange(2, idIndex + 1, responseSheet.getLastRow() - 1, 1).getValues();
  const duplicateIndex = ids.findIndex((row, index) => index + 2 !== submittedRow && String(row[0] || "").trim() === submittedId);
  if (duplicateIndex < 0) return { rejected: false, message: "Sin duplicado" };
  const existingRow = duplicateIndex + 2;
  const reason = `ENVIO_DUPLICADO_FORM: ID_EQUIPO ya existe en fila ${existingRow}`;
  appendIncidencia_(spreadsheet, responseSheet, submittedRow, reason, submittedId, headers, submittedValues);
  responseSheet.deleteRow(submittedRow);
  return { rejected: true, message: reason, recordId: submittedId, existingRow };
}

function detectResponseSheet_(spreadsheet, technicalSheetName) {
  const configured = Config.get("FORM_RESPONSES_SHEET", false);
  if (configured && spreadsheet.getSheetByName(configured)) return spreadsheet.getSheetByName(configured);
  if (spreadsheet.getSheetByName(Config.responseSheetName())) return spreadsheet.getSheetByName(Config.responseSheetName());
  return spreadsheet.getSheets().filter((sheet) => sheet.getName() !== technicalSheetName && sheet.getName() !== Config.incidenceSheetName()).find((sheet) => findIdColumnIndex_(getSheetHeaders_(sheet)) >= 0) || null;
}

/**
 * Elimina duplicados historicos y conserva la primera aparicion de cada equipo.
 * @returns {string} Resumen de la depuracion.
 */
function depurarDuplicadosHistoricos() {
  const spreadsheet = Config.spreadsheet();
  const responseSheet = detectResponseSheet_(spreadsheet, Config.technicalSheetName());
  if (!responseSheet || responseSheet.getLastRow() < 2) return "Sin respuestas para depurar";
  const headers = getSheetHeaders_(responseSheet);
  const idIndex = findIdColumnIndex_(headers);
  if (idIndex < 0) throw new Error("No existe columna ID_EQUIPO en respuestas");
  const seen = new Set();
  const rowsToDelete = [];
  const data = responseSheet.getRange(2, 1, responseSheet.getLastRow() - 1, headers.length).getValues();
  data.forEach((row, index) => {
    const id = String(row[idIndex] || "").trim();
    if (!id || !seen.has(id)) { if (id) seen.add(id); return; }
    const rowNumber = index + 2;
    appendIncidencia_(spreadsheet, responseSheet, rowNumber, `DUPLICADO_HISTORICO: ID_EQUIPO repetido`, id, headers, row);
    rowsToDelete.push(rowNumber);
  });
  rowsToDelete.sort((a, b) => b - a).forEach((rowNumber) => responseSheet.deleteRow(rowNumber));
  return `Depuracion completada. Duplicados movidos: ${rowsToDelete.length}`;
}

/**
 * Elimina columnas auxiliares de validacion y sincronizacion.
 * @returns {string} Resultado de la limpieza.
 */
function limpiarColumnasAuxiliaresRespuestas() {
  const sheet = Config.spreadsheet().getSheetByName(Config.responseSheetName());
  if (!sheet) throw new Error(`No se encontro hoja ${Config.responseSheetName()}`);
  const headers = getSheetHeaders_(sheet);
  const auxiliary = new Set(RESPONSE_STATUS_HEADERS_.map(normalizeHeader_));
  headers.forEach((header, index) => { if (auxiliary.has(normalizeHeader_(header))) sheet.deleteColumn(index + 1); });
  return "Columnas auxiliares eliminadas.";
}

/**
 * Ordena las respuestas por la marca temporal ascendente.
 * @returns {string} Resultado de la ordenacion.
 */
function ordenarRespuestasPorFechaAscendente() {
  const sheet = Config.spreadsheet().getSheetByName(Config.responseSheetName());
  if (!sheet) throw new Error(`No se encontro hoja ${Config.responseSheetName()}`);
  if (sheet.getLastRow() < 3) return "No hay suficientes respuestas para ordenar.";
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort({ column: 1, ascending: true });
  return "Respuestas ordenadas por Marca temporal.";
}

/**
 * Procesa una respuesta de formulario y elimina duplicados en caliente.
 * @param {Object} event Evento instalable onFormSubmit.
 * @returns {void}
 */
function onFormSubmit(event) {
  try {
    const spreadsheet = Config.spreadsheet();
    const responseSheet = detectResponseSheet_(spreadsheet, Config.technicalSheetName());
    const decision = rejectDuplicateSubmissionIfNeeded_(event, spreadsheet, responseSheet);
    if (decision.rejected) return;
  } catch (error) {
    // Un fallo del trigger no debe interrumpir otras automatizaciones del libro.
    return;
  }
}
