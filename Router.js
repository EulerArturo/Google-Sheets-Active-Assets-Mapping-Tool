/**
 * Atiende solicitudes GET del Web App y sirve el formulario o un archivo HTML autorizado.
 * @param {GoogleAppsScript.Events.DoGet} e Evento HTTP con parametros de consulta.
 * @returns {GoogleAppsScript.HTML.HtmlOutput} Contenido HTML de la respuesta.
 */
function doGet(e) {
  const parameters = (e && e.parameter) || {};
  const evidenciaId = safeFileId_(parameters.evidenciaId);
  const fileId = safeFileId_(parameters.fileId);
  const requestedFileId = evidenciaId || fileId;

  // Sin identificador se muestra el formulario principal.
  if (!requestedFileId) {
    return HtmlService.createTemplateFromFile("Index").evaluate()
      .setTitle("Formulario de Soporte")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Se rechazan identificadores que no tengan el formato esperado.
  try {
    const file = DriveApp.getFileById(requestedFileId);
    // La evidencia y los archivos publicados tienen validaciones de carpeta diferentes.
    const allowed = evidenciaId
      ? esEvidenciaHtmlValida_(file)
      : esArchivoHtmlPublicadoValido_(file);
    if (!allowed) {
      return HtmlService.createHtmlOutput("Acceso denegado.");
    }
    return HtmlService.createHtmlOutput(file.getBlob().getDataAsString("UTF-8"))
      .setTitle(file.getName())
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput("Archivo no encontrado.");
  }
}

/**
 * Atiende solicitudes POST autenticadas para cargar HTML o guardar datos tecnicos.
 * @param {GoogleAppsScript.Events.DoPost} e Evento HTTP con cuerpo JSON.
 * @returns {GoogleAppsScript.Content.TextOutput} Respuesta JSON serializada.
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e && e.postData && e.postData.contents || "{}");
    // El token debe coincidir con la propiedad protegida del proyecto.
    if (payload.token !== Config.appToken()) {
      return jsonResponse({ ok: false, status: 401, error: "Token invalido" });
    }

    if (String(payload.action || "upload_html") === "save_technical") {
      return saveTechnicalRecord(payload);
    }
    return uploadHtmlFile_(payload);
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error) });
  }
}

/**
 * Guarda un archivo recibido en Base64 dentro de la carpeta autorizada.
 * @param {Object} payload Datos de la carga.
 * @returns {GoogleAppsScript.Content.TextOutput} Resultado de la operacion.
 */
function uploadHtmlFile_(payload) {
  const folderId = payload.folder_id || Config.driveFolderId();
  if (!payload.file_name || !payload.content_base64) {
    return jsonResponse({ ok: false, error: "Faltan file_name o content_base64" });
  }

  const bytes = Utilities.base64Decode(payload.content_base64);
  const blob = Utilities.newBlob(bytes, payload.mime_type || "text/html", payload.file_name);
  const file = Config.folder(folderId).createFile(blob);
  const baseWebAppUrl = ScriptApp.getService().getUrl();
  const renderUrl = baseWebAppUrl
    ? `${baseWebAppUrl}?fileId=${encodeURIComponent(file.getId())}`
    : "";

  return jsonResponse({
    ok: true,
    fileId: file.getId(),
    url: file.getUrl(),
    renderUrl,
    recordId: payload.record_id || "",
  });
}
