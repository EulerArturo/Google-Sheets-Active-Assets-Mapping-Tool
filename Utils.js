/**
 * Construye una respuesta JSON para los endpoints del Web App.
 * @param {Object} obj Objeto que se serializara.
 * @returns {GoogleAppsScript.Content.TextOutput} Respuesta JSON.
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Normaliza un encabezado para comparaciones independientes de tildes y formato.
 * @param {*} text Encabezado original.
 * @returns {string} Encabezado normalizado.
 */
function normalizeHeader_(text) {
  let value = String(text || "").trim().toUpperCase();
  value = value
    .replace(/\u00c1/g, "A")
    .replace(/\u00c9/g, "E")
    .replace(/\u00cd/g, "I")
    .replace(/\u00d3/g, "O")
    .replace(/\u00da/g, "U")
    .replace(/\u00dc/g, "U")
    .replace(/\u00d1/g, "N");
  return value.replace(/[^A-Z0-9]+/g, "_").replace(/^_+|_+$/g, "");
}

/** @returns {string} Fecha y hora actual con la zona horaria del proyecto. */
function nowString_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

/**
 * Limpia un nombre para usarlo como carpeta de Drive.
 * @param {*} value Nombre original.
 * @returns {string} Nombre seguro y limitado a 80 caracteres.
 */
function normalizarNombreCarpeta_(value) {
  return String(value)
    .trim()
    .replace(/[\\/:*?"<>|]+/g, "-")
    .replace(/\s+/g, " ")
    .substring(0, 80)
    .trim();
}

/**
 * Recorta texto y protege valores que Sheets podria interpretar como formulas.
 * @param {*} value Valor original.
 * @param {number} maxLength Longitud maxima.
 * @returns {string} Texto normalizado.
 */
function normalizarTexto_(value, maxLength) {
  const text = String(value || "").trim();
  const limitado = maxLength ? text.slice(0, maxLength) : text;
  return protegerFormula_(limitado);
}

/**
 * Evita la evaluacion automatica de formulas enviadas a una hoja.
 * @param {string} value Texto a proteger.
 * @returns {string} Texto protegido.
 */
function protegerFormula_(value) {
  return /^[=+\-@]/.test(value) ? `'${value}` : value;
}

/**
 * Convierte un valor vacio en un marcador legible para evidencias.
 * @param {*} value Valor original.
 * @returns {string} Valor formateado.
 */
function formatearValor_(value) {
  return value === null || value === undefined || value === "" ? "-" : String(value);
}

/**
 * Escapa caracteres especiales antes de insertar texto en HTML.
 * @param {*} value Valor original.
 * @returns {string} Texto HTML seguro.
 */
function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Lee la fila de encabezados de una hoja.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Hoja consultada.
 * @returns {Array<*>} Encabezados encontrados.
 */
function getSheetHeaders_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) {
    return [];
  }
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/**
 * Crea un indice de columnas por encabezado normalizado.
 * @param {Array<*>} headers Encabezados de la hoja.
 * @returns {Object<string, number>} Mapa de encabezados a indices.
 */
function buildHeaderIndexMap_(headers) {
  const map = {};
  headers.forEach((header, index) => {
    map[normalizeHeader_(header)] = index;
  });
  return map;
}

/**
 * Valida el formato minimo de un identificador de archivo de Drive.
 * @param {*} value Identificador recibido.
 * @returns {string} Identificador valido o cadena vacia.
 */
function safeFileId_(value) {
  const fileId = String(value || "").trim();
  return /^[A-Za-z0-9_-]{10,}$/.test(fileId) ? fileId : "";
}
