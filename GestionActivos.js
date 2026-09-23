const NOMBRE_HOJA_ORIGEN_ACTIVOS = "REEMPLAZAR_CON_NOMBRE_HOJA_ORIGEN";
const NOMBRE_HOJA_DESTINO_ACTIVOS = "REEMPLAZAR_CON_NOMBRE_HOJA_DESTINO";
const ID_CARPETA_PDF_ACTIVOS = "REEMPLAZAR_CON_FOLDER_ID";

/**
 * Agrega al libro el menu de operaciones de activos.
 * @returns {void}
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Formulario AF")
    .addItem("1. Mapear Fila Seleccionada", "mapearMovimientoActivo")
    .addSeparator()
    .addItem("2. Adjuntar PDF Firmado", "abrirVentanaSubirPdf")
    .addToUi();
}

/**
 * Copia los responsables y hasta tres activos de una fila al formato AF.
 * @returns {void}
 */
function mapearMovimientoActivo() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // Obtiene las hojas del libro vinculado y valida su existencia.
  const sourceSheet = spreadsheet.getSheetByName(NOMBRE_HOJA_ORIGEN_ACTIVOS);
  const targetSheet = spreadsheet.getSheetByName(NOMBRE_HOJA_DESTINO_ACTIVOS);
  if (!sourceSheet || !targetSheet) {
    SpreadsheetApp.getUi().alert(
      `Verifica que existan las hojas '${NOMBRE_HOJA_ORIGEN_ACTIVOS}' y '${NOMBRE_HOJA_DESTINO_ACTIVOS}'.`
    );
    return;
  }
  // Confirma que la seleccion se encuentre en la hoja de origen y no en el encabezado.
  const activeSheet = spreadsheet.getActiveSheet();
  const rowNumber = activeSheet.getActiveCell().getRow();
  if (activeSheet.getName() !== NOMBRE_HOJA_ORIGEN_ACTIVOS || rowNumber < 2) {
    SpreadsheetApp.getUi().alert(
      `Selecciona una fila valida en '${NOMBRE_HOJA_ORIGEN_ACTIVOS}'.`
    );
    return;
  }

  // Lee desde la columna D los 28 campos definidos para el movimiento.
  const values = sourceSheet.getRange(rowNumber, 4, 1, 28).getValues()[0];

  // Asigna los datos de los colaboradores a las posiciones del formato.
  [["G12", values[0]], ["G13", values[1]], ["G14", values[2]], ["G15", values[3]], ["G16", values[4]],
    ["U12", values[5]], ["U13", values[6]], ["U14", values[7]], ["U15", values[8]], ["U16", values[9]]]
    .forEach(([cell, value]) => targetSheet.getRange(cell).setValue(value));
  // Asigna los seis campos de cada uno de los tres activos.
  for (let asset = 0; asset < 3; asset += 1) {
    const offset = 10 + asset * 6;
    const targetRow = 20 + asset;
    [[`C${targetRow}`, values[offset]], [`J${targetRow}`, values[offset + 1]], [`N${targetRow}`, values[offset + 2]], [`Q${targetRow}`, values[offset + 3]], [`T${targetRow}`, values[offset + 4]], [`W${targetRow}`, values[offset + 5]]]
      .forEach(([cell, value]) => targetSheet.getRange(cell).setValue(value));
  }
  // Cambia a la hoja de destino para que el usuario revise el resultado.
  spreadsheet.setActiveSheet(targetSheet);
  SpreadsheetApp.getUi().alert("Mapeo realizado con exito.");
}

/**
 * Abre la ventana modal para adjuntar un PDF firmado a la fila seleccionada.
 * @returns {void}
 */
function abrirVentanaSubirPdf() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  if (activeSheet.getName() !== NOMBRE_HOJA_ORIGEN_ACTIVOS || activeSheet.getActiveCell().getRow() < 2) {
    SpreadsheetApp.getUi().alert(`Selecciona una fila valida en '${NOMBRE_HOJA_ORIGEN_ACTIVOS}'.`);
    return;
  }
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile("SubirArchivoHTML").setWidth(450).setHeight(200),
    "Adjuntar PDF Firmado",
  );
}

/**
 * Guarda un PDF en Drive y escribe su enlace en la columna AF.
 * @param {string} dataBase64 Contenido Base64 del archivo.
 * @param {string} nombreArchivo Nombre original del archivo.
 * @returns {string} Mensaje de confirmacion.
 */
function guardarPdfEnDriveYEnlazar(dataBase64, nombreArchivo) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(NOMBRE_HOJA_ORIGEN_ACTIVOS);
  if (!sheet) {
    throw new Error(`No se encontro la hoja '${NOMBRE_HOJA_ORIGEN_ACTIVOS}'.`);
  }
  const rowNumber = sheet.getActiveCell().getRow();
  const folder = DriveApp.getFolderById(ID_CARPETA_PDF_ACTIVOS);
  const file = folder.createFile(Utilities.newBlob(
    Utilities.base64Decode(dataBase64), "application/pdf", normalizarNombreCarpeta_(nombreArchivo),
  ));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  sheet.getRange(rowNumber, 32).setValue(`=HYPERLINK("${file.getUrl()}", "Ver PDF Firmado")`);
  return "Archivo subido y enlazado con exito.";
}
