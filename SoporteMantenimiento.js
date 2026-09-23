const REGLAS_FORMULARIO_ = {
  required: [
    "fechaAtencion", "sede", "tipoEquipo", "tipoMantenimiento",
    "herramientasUtilizadas", "servidorPublico", "servidorCedula",
    "areaEntidad", "tecnicoNombre", "tecnicoCedula", "descripcion",
    "proteccionDatosAceptada", "equipo", "servidorCargo", "diagnostico",
    "resolvioNecesidad",
  ],
  maxLengths: {
    idSolicitud: 20, fechaAtencion: 10, sede: 80, tipoEquipo: 80, equipo: 60,
    tipoMantenimiento: 20, herramientasUtilizadas: 500, servidorPublico: 80,
    servidorCedula: 15, servidorCargo: 80, areaEntidad: 80, tecnicoNombre: 80,
    tecnicoCedula: 15, descripcion: 500, diagnostico: 500, soporteExterno: 3,
    soporteExternoObservaciones: 500, resolvioNecesidad: 3,
    proteccionDatosAceptada: 3, estadoCaso: 12, fechaCierre: 10,
    expectativas: 300, clasificacionTexto: 600, clasificacionCeldas: 200,
    hardware: 500, software: 500, sistemas: 500,
  },
  maxFotos: 5,
  maxFotoBytes: 3 * 1024 * 1024,
  allowedMimeTypes: ["image/jpeg", "image/png", "image/webp"],
};

const CAMPOS_LABEL_ = {
  fechaAtencion: "Fecha Atencion", sede: "Sede o sucursal", tipoEquipo: "Tipo de equipo",
  tipoMantenimiento: "Tipo de mantenimiento", herramientasUtilizadas: "Herramientas utilizadas",
  servidorPublico: "Profesional / Usuario", servidorCedula: "Cedula del Profesional / Usuario",
  areaEntidad: "Area o Entidad", tecnicoNombre: "Nombre del Tecnico",
  tecnicoCedula: "Cedula del Tecnico", descripcion: "Descripcion del Requerimiento",
};

/**
 * Genera el siguiente identificador consecutivo de solicitud bajo bloqueo.
 * @returns {string} Identificador con formato CS-0000.
 */
function obtenerSiguienteId() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return obtenerSiguienteId_();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Calcula y persiste el siguiente identificador sin adquirir un bloqueo adicional.
 * @returns {string} Identificador formateado.
 */
function obtenerSiguienteId_() {
  // Lee el contador persistido y usa la hoja como respaldo inicial.
  let ultimoNumero = parseInt(Config.get("ultimoIdSolicitud", false), 10);
  if (!ultimoNumero) {
    ultimoNumero = obtenerUltimoIdDesdeHoja_();
  }
  // Persiste el nuevo contador antes de devolver el identificador.
  const siguienteNumero = (ultimoNumero || 0) + 1;
  Config.set("ultimoIdSolicitud", siguienteNumero);
  return formatearIdSolicitud_(siguienteNumero);
}

/**
 * Valida, registra y genera la evidencia de una solicitud de mantenimiento.
 * @param {Object} formData Datos enviados por el formulario web.
 * @returns {string} Identificador asignado a la solicitud.
 */
function procesarFormularioWeb(formData) {
  if (!formData) {
    throw new Error("No se recibieron datos del formulario.");
  }

  // Normaliza y valida antes de adquirir recursos de Drive o Sheets.
  const datos = normalizarFormulario_(formData);
  validarFormulario_(datos);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // El bloqueo cubre la comprobacion de ID, la evidencia y la escritura de la fila.
    const registroSheet = Config.spreadsheet().getSheetByName(Config.maintenanceSheetName());
    if (!registroSheet) {
      throw new Error(`No se encontro la hoja '${Config.maintenanceSheetName()}'.`);
    }

    // Usa el ID recibido solo si tiene formato valido y no esta repetido.
    const solicitudId = String(datos.idSolicitud || "").trim();
    const idValido = /^CS-\d{4}$/.test(solicitudId);
    datos.idSolicitud = !idValido || idYaExiste_(registroSheet, solicitudId)
      ? obtenerSiguienteId_()
      : solicitudId;

    // Crea la evidencia antes de guardar las URLs junto con la solicitud.
    const evidencia = crearEvidenciaEnDrive_(datos);
    // Inserta una fila con el orden definido por MANTENIMIENTOS.
    registroSheet.appendRow([
      datos.idSolicitud || "", datos.fechaAtencion || "", datos.sede || "",
      datos.tipoEquipo || "", datos.equipo || "", datos.tipoMantenimiento || "",
      datos.herramientasUtilizadas || "", datos.servidorPublico || "",
      datos.servidorCedula || "", datos.servidorCargo || "", datos.areaEntidad || "",
      datos.tecnicoNombre || "", datos.tecnicoCedula || "", datos.descripcion || "",
      datos.diagnostico || "", datos.soporteExterno || "",
      datos.soporteExternoObservaciones || "", datos.resolvioNecesidad || "",
      datos.proteccionDatosAceptada || "", datos.estadoCaso || "", datos.fechaCierre || "",
      datos.expectativas || "", datos.hardware || "No aplica", datos.software || "No aplica",
      datos.sistemas || "No aplica", evidencia.url, evidencia.folderUrl, new Date(),
    ]);
    return datos.idSolicitud || "";
  } finally {
    lock.releaseLock();
  }
}

/**
 * Reinicia la hoja de mantenimiento y crea sus encabezados estandar.
 * @returns {void}
 */
function resetRegistroSolicitudes() {
  // Obtiene la hoja existente o la crea dentro del libro configurado.
  const sheet = Config.spreadsheet().getSheetByName(Config.maintenanceSheetName())
    || Config.spreadsheet().insertSheet(Config.maintenanceSheetName());
  sheet.clearContents();
  // Define el contrato de columnas consumido por el resto del sistema.
  sheet.appendRow([
    "ID Solicitud", "Fecha Atención", "Sede", "Tipo Equipo", "Equipo",
    "Tipo Mantenimiento", "Herramientas Utilizadas", "Servidor Público",
    "Servidor Cédula", "Servidor Cargo", "Área / Entidad", "Técnico Nombre",
    "Técnico Cédula", "Descripción", "Diagnóstico", "Soporte Externo",
    "Soporte Externo Observaciones", "Resolvió Necesidad", "Protección Datos Aceptada",
    "Estado Caso", "Fecha Cierre", "Expectativas", "Hardware", "Software", "Sistemas",
    "Evidencia HTML", "Carpeta Evidencia", "Fecha de Registro",
  ]);
  sheet.setFrozenRows(1);
}

/**
 * Crea la carpeta, fotografias y documento HTML de una solicitud.
 * @param {Object} formData Datos normalizados de la solicitud.
 * @returns {{url: string, folderUrl: string, fotosUrls: Array<string>}} Enlaces generados.
 */
function crearEvidenciaEnDrive_(formData) {
  // Cada solicitud tiene una carpeta propia para sus archivos.
  const rootFolder = Config.evidenceRootFolder();
  const solicitudFolder = obtenerCarpetaSolicitud_(rootFolder, formData);
  // Guarda las fotos y genera el documento con los datos resultantes.
  const fotos = guardarFotosEnCarpeta_(solicitudFolder, formData.fotosEvidencia);
  const file = Config.evidenceHtmlFolder().createFile(
    `Evidencia_${formData.idSolicitud}.html`,
    crearHtmlEvidencia_(formData, fotos),
    MimeType.HTML,
  );
  const baseUrl = ScriptApp.getService().getUrl();
  return {
    url: baseUrl ? `${baseUrl}?evidenciaId=${file.getId()}` : file.getUrl(),
    folderUrl: solicitudFolder.getUrl(),
    fotosUrls: fotos.map((foto) => foto.url),
  };
}

/**
 * Obtiene o crea la carpeta individual de una solicitud.
 * @param {GoogleAppsScript.Drive.Folder} rootFolder Carpeta raiz.
 * @param {Object} formData Datos que identifican la solicitud.
 * @returns {GoogleAppsScript.Drive.Folder} Carpeta de la solicitud.
 */
function obtenerCarpetaSolicitud_(rootFolder, formData) {
  const safeName = normalizarNombreCarpeta_(
    `${formData.idSolicitud || "Solicitud"} - ${formData.equipo || "Sin_equipo"}`,
  );
  const iterator = rootFolder.getFoldersByName(safeName);
  return iterator.hasNext() ? iterator.next() : rootFolder.createFolder(safeName);
}

/**
 * Guarda fotografias Base64 en una carpeta de Drive.
 * @param {GoogleAppsScript.Drive.Folder} folder Carpeta destino.
 * @param {Array<Object>} fotos Fotografias recibidas.
 * @returns {Array<Object>} Metadatos y representaciones de las fotos.
 */
function guardarFotosEnCarpeta_(folder, fotos) {
  if (!Array.isArray(fotos) || fotos.length === 0) {
    return [];
  }
  // Valida cantidad, MIME y tamaño antes de crear archivos.
  validarFotos_(fotos);
  return fotos.map((foto, index) => {
    const fileName = `Foto_${String(index + 1).padStart(2, "0")}_${normalizarNombreCarpeta_(foto.name || `foto_${index + 1}`)}`;
    const blob = Utilities.newBlob(
      Utilities.base64Decode(foto.data), foto.mimeType || "image/jpeg", fileName,
    );
    const file = folder.createFile(blob);
    const dataUrl = `data:${foto.mimeType || "image/jpeg"};base64,${Utilities.base64Encode(blob.getBytes())}`;
    return {
      name: fileName,
      url: file.getUrl(),
      viewUrl: `https://drive.google.com/uc?export=view&id=${file.getId()}`,
      dataUrl,
    };
  });
}

/**
 * Construye el HTML de evidencia con datos y fotografias embebidas.
 * @param {Object} formData Datos de la solicitud.
 * @param {Array<Object>} fotos Fotografias guardadas.
 * @returns {string} Documento HTML completo.
 */
function crearHtmlEvidencia_(formData, fotos) {
  const filas = [
    ["ID Solicitud", formData.idSolicitud], ["Fecha Atencion", formData.fechaAtencion],
    ["Sede o sucursal", formData.sede], ["Tipo de equipo", formData.tipoEquipo],
    ["Equipo", formData.equipo], ["Tipo de mantenimiento", formData.tipoMantenimiento],
    ["Herramientas utilizadas", formData.herramientasUtilizadas],
    ["Servidor Publico / Usuario", formData.servidorPublico],
    ["Cedula Servidor Publico / Usuario", formData.servidorCedula],
    ["Cargo Servidor Publico / Usuario", formData.servidorCargo], ["Area o Entidad", formData.areaEntidad],
    ["Nombre del Tecnico", formData.tecnicoNombre], ["Cedula del Tecnico", formData.tecnicoCedula],
    ["Descripcion del Requerimiento", formData.descripcion], ["Diagnostico y Justificacion", formData.diagnostico],
    ["Soporte Externo", formData.soporteExterno], ["Observaciones Soporte Externo", formData.soporteExternoObservaciones],
    ["Resolvio Necesidad", formData.resolvioNecesidad], ["Proteccion de datos", formData.proteccionDatosAceptada],
    ["Estado del caso", formData.estadoCaso], ["Fecha de cierre", formData.fechaCierre],
    ["Expectativas", formData.expectativas], ["Clasificacion", formData.clasificacionTexto],
    ["Fecha de registro", Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")],
  ];
  const rowsHtml = filas.map(([label, value]) => `<tr><th>${escapeHtml_(label)}</th><td>${escapeHtml_(formatearValor_(value))}</td></tr>`).join("");
  const photosHtml = (fotos || []).map((foto) => `
    <figure class="photo">
      <img src="${escapeHtml_(foto.dataUrl || foto.viewUrl)}" alt="${escapeHtml_(foto.name)}" loading="lazy">
      <figcaption>${escapeHtml_(foto.name)}</figcaption>
    </figure>`).join("");
  return `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Evidencia ${escapeHtml_(formatearValor_(formData.idSolicitud))}</title><style>body{font-family:Arial,sans-serif;background:#f4f6fb;color:#1e2a44;padding:24px}.card{background:#fff;border-radius:12px;padding:24px;box-shadow:0 12px 30px rgba(15,23,42,.08)}table{width:100%;border-collapse:collapse}th,td{text-align:left;padding:8px 10px;border-bottom:1px solid #e6ecf5;vertical-align:top}th{width:280px;color:#5a6b88}.photos{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:16px;margin-top:16px}.photo{margin:0}.photo img{display:block;width:100%;max-height:360px;object-fit:contain;background:#f4f6fb;border-radius:8px}.photo figcaption{font-size:12px;margin-top:6px;word-break:break-word}</style></head><body><div class="card"><h1>Formato de Atencion a Usuarios - Proceso Informatico</h1><table>${rowsHtml}</table>${photosHtml ? `<h2>Fotografias</h2><div class="photos">${photosHtml}</div>` : ""}</div></body></html>`;
}

/** @param {GoogleAppsScript.Drive.File} file Archivo candidato. @returns {boolean} Si es una evidencia autorizada. */
function esEvidenciaHtmlValida_(file) {
  if (!file || file.getMimeType() !== MimeType.HTML || !/^Evidencia_/.test(file.getName())) {
    return false;
  }
  const parents = file.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === Config.evidenceHtmlFolderId()) {
      return true;
    }
  }
  return false;
}

/** @param {GoogleAppsScript.Drive.File} file Archivo candidato. @returns {boolean} Si pertenece a la carpeta publicada. */
function esArchivoHtmlPublicadoValido_(file) {
  if (!file || file.getMimeType() !== MimeType.HTML) {
    return false;
  }
  const parents = file.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === Config.driveFolderId()) {
      return true;
    }
  }
  return false;
}

/**
 * Normaliza los valores recibidos del formulario al modelo interno.
 * @param {Object} formData Datos sin procesar.
 * @returns {Object} Datos normalizados.
 */
function normalizarFormulario_(formData) {
  const categoria = (value, maxLength) => {
    const text = Array.isArray(value) ? value.join(", ") : String(value || "").trim();
    return text ? normalizarTexto_(text, maxLength) : "No aplica";
  };
  const estadoCaso = normalizarTexto_(formData.estadoCaso, 12) || "Abierto";
  return {
    idSolicitud: normalizarTexto_(formData.idSolicitud, 20),
    fechaAtencion: normalizarTexto_(formData.fechaAtencion, 10), sede: normalizarTexto_(formData.sede, 80),
    tipoEquipo: normalizarTexto_(formData.tipoEquipo, 80), equipo: normalizarTexto_(formData.equipo, 60),
    tipoMantenimiento: normalizarTexto_(formData.tipoMantenimiento, 20),
    herramientasUtilizadas: normalizarTexto_(formData.herramientasUtilizadas, 500),
    servidorPublico: normalizarTexto_(formData.servidorPublico, 80), servidorCedula: normalizarTexto_(formData.servidorCedula, 15),
    servidorCargo: normalizarTexto_(formData.servidorCargo, 80), areaEntidad: normalizarTexto_(formData.areaEntidad, 80),
    tecnicoNombre: normalizarTexto_(formData.tecnicoNombre, 80), tecnicoCedula: normalizarTexto_(formData.tecnicoCedula, 15),
    descripcion: normalizarTexto_(formData.descripcion, 500), diagnostico: normalizarTexto_(formData.diagnostico, 500),
    soporteExterno: normalizarTexto_(formData.soporteExterno, 3), soporteExternoObservaciones: normalizarTexto_(formData.soporteExternoObservaciones, 500),
    resolvioNecesidad: normalizarTexto_(formData.resolvioNecesidad, 3), proteccionDatosAceptada: formData.proteccionDatosAceptada === true || formData.proteccionDatosAceptada === "Sí",
    estadoCaso, fechaCierre: estadoCaso === "Cerrado" ? normalizarTexto_(formData.fechaCierre, 10) : "",
    expectativas: normalizarTexto_(formData.expectativas, 300), hardware: categoria(formData.hardware, 500),
    software: categoria(formData.software, 500), sistemas: categoria(formData.sistemas, 500),
    clasificacionTexto: normalizarTexto_(formData.clasificacionTexto, 600), clasificacionCeldas: normalizarTexto_(formData.clasificacionCeldas, 200),
    fotosEvidencia: Array.isArray(formData.fotosEvidencia) ? formData.fotosEvidencia : [],
  };
}

/**
 * Valida campos, formatos, clasificaciones y fotografias de una solicitud.
 * @param {Object} formData Datos normalizados.
 * @returns {void}
 */
function validarFormulario_(formData) {
  // Comprueba presencia de campos obligatorios antes de validar formatos.
  REGLAS_FORMULARIO_.required.forEach((field) => {
      // Valida los formatos y catalogos aceptados por el proceso.
    if (!formData[field]) throw new Error(`Falta el campo obligatorio: ${CAMPOS_LABEL_[field] || field}.`);
  });
  if (!/^\d{4}-\d{2}-\d{2}$/.test(formData.fechaAtencion)) throw new Error("Fecha de atencion no valida.");
  if (!["Preventivo", "Correctivo", "Seguimiento"].includes(formData.tipoMantenimiento)) throw new Error("Tipo de mantenimiento no valido.");
  if (!/^\d{5,15}$/.test(formData.servidorCedula) || !/^\d{5,15}$/.test(formData.tecnicoCedula)) throw new Error("Cedula no valida.");
  if (formData.soporteExterno === "Sí" && !formData.soporteExternoObservaciones) throw new Error("Indique observaciones de soporte externo.");
  if (formData.proteccionDatosAceptada !== true) throw new Error("Debe aceptar la protección de datos personales.");
  validarEstadoCaso_(formData.estadoCaso, formData.fechaCierre);
  validarClasificacion_(formData.clasificacionCeldas);
  validarFotos_(formData.fotosEvidencia);
}

/** @param {string} estadoCaso Estado seleccionado. @param {string} fechaCierre Fecha de cierre. @returns {void} */
function validarEstadoCaso_(estadoCaso, fechaCierre) {
  if (!["Abierto", "En proceso", "Cerrado"].includes(estadoCaso)) throw new Error("Estado del caso no valido.");
  if (estadoCaso === "Cerrado" && (!fechaCierre || !/^\d{4}-\d{2}-\d{2}$/.test(fechaCierre))) throw new Error("Fecha de cierre no valida.");
}

/** @param {string} value Codigos de clasificacion separados por coma. @returns {void} */
function validarClasificacion_(value) {
  const codigos = String(value || "").split(",").map((item) => item.trim()).filter(Boolean);
  const regex = /^(C(1[4-9]|2\d|3[0-6])|F(1[4-9]|2\d|3[0-4]))$/;
  codigos.forEach((codigo) => { if (!regex.test(codigo)) throw new Error("Clasificacion no valida."); });
}

/** @param {Array<Object>} fotos Fotografias recibidas. @returns {void} */
function validarFotos_(fotos) {
  if (!Array.isArray(fotos)) return;
  if (fotos.length > REGLAS_FORMULARIO_.maxFotos) throw new Error(`Maximo ${REGLAS_FORMULARIO_.maxFotos} fotos permitidas.`);
  fotos.forEach((foto, index) => {
    if (!foto || !foto.data) throw new Error(`Foto ${index + 1} no valida.`);
    if (!REGLAS_FORMULARIO_.allowedMimeTypes.includes(foto.mimeType)) throw new Error(`Formato no permitido en foto ${index + 1}.`);
    if (Math.floor(String(foto.data).length * 3 / 4) > REGLAS_FORMULARIO_.maxFotoBytes) throw new Error(`Foto ${index + 1} supera 3 MB.`);
  });
}

/**
 * Busca el ultimo identificador almacenado en MANTENIMIENTOS.
 * @returns {number} Ultimo consecutivo encontrado o cero.
 */
function obtenerUltimoIdDesdeHoja_() {
  const sheet = Config.spreadsheet().getSheetByName(Config.maintenanceSheetName());
  if (!sheet || sheet.getLastRow() < 1) return 0;
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().flat();
  for (let index = values.length - 1; index >= 0; index -= 1) {
    const number = extraerNumeroId_(values[index]);
    if (number) return number;
  }
  return 0;
}

/** @param {*} value Valor de la celda. @returns {number} Parte numerica del ID. */
function extraerNumeroId_(value) {
  const match = String(value || "").trim().match(/^CS-(\d{4})$/);
  return match ? parseInt(match[1], 10) : 0;
}

/** @param {number} numero Consecutivo. @returns {string} ID con formato CS-0000. */
function formatearIdSolicitud_(numero) {
  return `CS-${String(numero).padStart(4, "0")}`;
}

/** @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Hoja de mantenimiento. @param {string} id ID buscado. @returns {boolean} Si el ID ya existe. */
function idYaExiste_(sheet, id) {
  if (sheet.getLastRow() < 1) return false;
  return Boolean(sheet.getRange(1, 1, sheet.getLastRow(), 1).createTextFinder(id).matchEntireCell(true).findNext());
}

/**
 * Regenera las evidencias desde las filas existentes y sincroniza sus URLs.
 * @returns {void}
 */
function actualizarEvidenciasDesdeHoja() {
  const sheet = Config.spreadsheet().getSheetByName(Config.maintenanceSheetName());
  if (!sheet || sheet.getLastRow() < 2) return;
  const htmlFolder = Config.evidenceHtmlFolder();
  // Lee todas las filas para regenerar cada evidencia de forma consistente.
  const data = sheet.getDataRange().getDisplayValues();
  data.slice(1).forEach((row, index) => {
    if (!row[0]) return;
    // Reconstruye el modelo desde las posiciones definidas en MANTENIMIENTOS.
    const formData = {
      idSolicitud: row[0], fechaAtencion: row[1], sede: row[2], tipoEquipo: row[3], equipo: row[4],
      tipoMantenimiento: row[5], herramientasUtilizadas: row[6], servidorPublico: row[7], servidorCedula: row[8],
      servidorCargo: row[9], areaEntidad: row[10], tecnicoNombre: row[11], tecnicoCedula: row[12], descripcion: row[13],
      diagnostico: row[14], soporteExterno: row[15], soporteExternoObservaciones: row[16], resolvioNecesidad: row[17],
      proteccionDatosAceptada: row[18], estadoCaso: row[19], fechaCierre: row[20], expectativas: row[21],
      hardware: row[22], software: row[23], sistemas: row[24], fechaRegistro: row[27] || row[1],
      clasificacionTexto: [row[22] !== "No aplica" ? `Hardware - ${row[22]}` : "", row[23] !== "No aplica" ? `Software - ${row[23]}` : "", row[24] !== "No aplica" ? `Sistemas - ${row[24]}` : ""].filter(Boolean).join(" | "),
    };
    // Actualiza o crea el archivo HTML y escribe su URL en la columna de evidencia.
    const fileName = `Evidencia_${row[0]}.html`;
    const files = htmlFolder.getFilesByName(fileName);
    const evidenceFile = files.hasNext()
      ? files.next()
      : htmlFolder.createFile(fileName, crearHtmlEvidencia_(formData, []), MimeType.HTML);
    evidenceFile.setContent(crearHtmlEvidencia_(formData, []));
    const baseUrl = ScriptApp.getService().getUrl();
    const evidenceUrl = baseUrl
      ? `${baseUrl}?evidenciaId=${encodeURIComponent(evidenceFile.getId())}`
      : evidenceFile.getUrl();
    sheet.getRange(index + 2, 26).setValue(evidenceUrl);
  });
}
