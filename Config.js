const CONFIG_DEFAULTS_ = Object.freeze({
  EVIDENCE_ROOT_FOLDER_ID: "REEMPLAZAR_CON_FOLDER_ID",
  EVIDENCE_HTML_FOLDER_ID: "REEMPLAZAR_CON_FOLDER_ID",
  PDF_FOLDER_ID: "REEMPLAZAR_CON_PDF_FOLDER_ID",
  MANTENIMIENTOS_SHEET: "MANTENIMIENTOS",
  TECHNICAL_SHEET: "TECNICA_EQUIPOS",
  FORM_RESPONSES_SHEET: "RESPONSABLES DE EQUIPOS",
  INCIDENCIAS_SHEET: "INCIDENCIAS_FORM",
  ASSETS_SOURCE_SHEET: "MOVIMIENTO_DE_ACTIVOS",
  ASSETS_TARGET_SHEET: "FORMATO MOVIMIENTO AF",
});

const Config = Object.freeze({
  /**
   * Obtiene una propiedad del proyecto y aplica un valor predeterminado opcional.
   * @param {string} key Nombre de la propiedad.
   * @param {boolean} required Indica si la propiedad debe existir.
   * @returns {string} Valor configurado.
   */
  get(key, required) {
    const value = PropertiesService.getScriptProperties().getProperty(key)
      || CONFIG_DEFAULTS_[key]
      || "";
    if (required && !value) {
      throw new Error(`Falta ${key} en Script Properties.`);
    }
    return value;
  },
  /**
   * Guarda una propiedad del proyecto como texto.
   * @param {string} key Nombre de la propiedad.
   * @param {*} value Valor que se almacenara.
   * @returns {void}
   */
  set(key, value) {
    PropertiesService.getScriptProperties().setProperty(key, String(value));
  },
  /** @returns {string} Identificador del Spreadsheet principal. */
  spreadsheetId() {
    return this.get("SPREADSHEET_ID", true);
  },
  /** @returns {string} Token de autenticacion de la API. */
  appToken() {
    return this.get("APP_TOKEN", true);
  },
  /** @returns {string} Identificador de la carpeta de archivos publicados. */
  driveFolderId() {
    return this.get("DRIVE_FOLDER_ID", true);
  },
  /** @returns {string} Identificador de la carpeta raiz de evidencias. */
  evidenceRootFolderId() {
    return this.get("EVIDENCE_ROOT_FOLDER_ID", true);
  },
  /** @returns {string} Identificador de la carpeta de evidencias HTML. */
  evidenceHtmlFolderId() {
    return this.get("EVIDENCE_HTML_FOLDER_ID", true);
  },
  /** @returns {string} Identificador de la carpeta de PDF de activos. */
  pdfFolderId() {
    return this.get("PDF_FOLDER_ID", true);
  },
  sheetName(key, fallbackKey) {
    return this.get(key, false) || this.get(fallbackKey, true);
  },
  /** @returns {string} Nombre de la hoja de mantenimiento. */
  maintenanceSheetName() {
    return this.get("MAINTENANCE_SHEET_NAME", false) || this.get("MANTENIMIENTOS_SHEET", true);
  },
  /** @returns {string} Nombre de la hoja tecnica. */
  technicalSheetName() {
    return this.get("TECHNICAL_SHEET_NAME", false) || this.get("TECHNICAL_SHEET", true);
  },
  /** @returns {string} Nombre de la hoja de respuestas. */
  responseSheetName() {
    return this.get("FORM_RESPONSES_SHEET", true);
  },
  /** @returns {string} Nombre de la hoja de incidencias. */
  incidenceSheetName() {
    return this.get("INCIDENCIAS_SHEET", true);
  },
  /** @returns {string} Nombre de la hoja origen de activos. */
  assetsSourceSheetName() {
    return this.get("ASSETS_SOURCE_SHEET", true);
  },
  /** @returns {string} Nombre de la hoja destino de activos. */
  assetsTargetSheetName() {
    return this.get("ASSETS_TARGET_SHEET", true);
  },
  /** @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} Spreadsheet principal. */
  spreadsheet() {
    return SpreadsheetApp.openById(this.spreadsheetId());
  },
  /**
   * Abre un Spreadsheet por identificador.
   * @param {string} spreadsheetId Identificador del Spreadsheet.
   * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} Libro solicitado.
   */
  spreadsheetById(spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId || this.spreadsheetId());
  },
  /**
   * Obtiene una carpeta de Drive.
   * @param {string} folderId Identificador de la carpeta.
   * @returns {GoogleAppsScript.Drive.Folder} Carpeta solicitada.
   */
  folder(folderId) {
    return DriveApp.getFolderById(folderId);
  },
  /** @returns {GoogleAppsScript.Drive.Folder} Carpeta raiz de evidencias. */
  evidenceRootFolder() {
    return this.folder(this.evidenceRootFolderId());
  },
  /** @returns {GoogleAppsScript.Drive.Folder} Carpeta de evidencias HTML. */
  evidenceHtmlFolder() {
    return this.folder(this.evidenceHtmlFolderId());
  },
  /** @returns {GoogleAppsScript.Drive.Folder} Carpeta de PDF de activos. */
  pdfFolder() {
    return this.folder(this.pdfFolderId());
  },
});
