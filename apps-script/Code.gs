// IMPORTANTE: en el editor de Apps Script, antes de implementar, ve a
// Servicios (icono +) y agrega "Drive API" (Advanced Google Services).
// Sin eso, la conversion de archivos .xls/.xlsx no va a funcionar.

function doGet(e) {
  try {
    if (e.parameter.obras === "1") return jsonOutput_(listarObras_());
    if (e.parameter.presupuestos === "1") return jsonOutput_(listarPresupuestos_());
    if (e.parameter.previsualizar) return jsonOutput_(parsearPresupuesto_(e.parameter.previsualizar));
    if (e.parameter.obra) return jsonOutput_(leerObra_(e.parameter.obra));
    if (e.parameter.informeSupervision) return jsonOutput_(obtenerDatosInforme_(e.parameter.informeSupervision));
    if (e.parameter.historialInformes) return jsonOutput_(listarHistorialInformes_(e.parameter.historialInformes));
    return jsonOutput_({ error: "Parametro no reconocido" });
  } catch (err) {
    return jsonOutput_({ error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.accion === "crear_obra") return jsonOutput_(crearObra_(body));
    if (body.accion === "guardar_medida") return jsonOutput_(guardarMedida_(body));
    if (body.accion === "copiar_medidas") return jsonOutput_(copiarMedidas_(body));
    if (body.accion === "editar_medida") return jsonOutput_(editarMedida_(body));
    if (body.accion === "borrar_medida") return jsonOutput_(borrarMedida_(body));
    if (body.accion === "borrar_obra") return jsonOutput_(borrarObra_(body));
    if (body.accion === "editar_item") return jsonOutput_(editarItemPresupuesto_(body));
    if (body.accion === "agregar_item") return jsonOutput_(agregarItemPresupuesto_(body));
    if (body.accion === "borrar_item") return jsonOutput_(borrarItemPresupuesto_(body));
    if (body.accion === "actualizar_reportes") return jsonOutput_(actualizarReportesObra_(body));
    if (body.accion === "guardar_datos_informe") return jsonOutput_(guardarDatosInforme_(body));
    if (body.accion === "subir_documento_informe") return jsonOutput_(subirDocumentoInforme_(body));
    if (body.accion === "calcular_resumen_informe") return jsonOutput_(calcularResumenInforme_(body));
    if (body.accion === "registrar_informe_generado") return jsonOutput_(registrarInformeGenerado_(body));
    if (body.accion === "subir_informe_generado") return jsonOutput_(subirInformeGeneradoDocx_(body));
    return jsonOutput_({ ok: false, error: "Accion no reconocida" });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err) });
  }
}

// ---------- Carpetas de Drive (se crean solas la primera vez) ----------

function obtenerCarpetaId_(clave, nombre, idPadre) {
  var props = PropertiesService.getScriptProperties();
  var cacheado = props.getProperty(clave);
  if (cacheado) {
    try {
      DriveApp.getFolderById(cacheado);
      return cacheado;
    } catch (e) {
      // la carpeta ya no existe, se vuelve a crear mas abajo
    }
  }
  var padre = idPadre ? DriveApp.getFolderById(idPadre) : DriveApp.getRootFolder();
  var it = padre.getFoldersByName(nombre);
  var carpeta = it.hasNext() ? it.next() : padre.createFolder(nombre);
  props.setProperty(clave, carpeta.getId());
  return carpeta.getId();
}

function carpetaRaiz_() { return obtenerCarpetaId_("carpetaRaiz", "Memorias JACBEL", null); }
function carpetaPresupuestos_() { return obtenerCarpetaId_("carpetaPresupuestos", "Presupuestos", carpetaRaiz_()); }
function carpetaFotos_() { return obtenerCarpetaId_("carpetaFotos", "Fotos", carpetaRaiz_()); }
function carpetaObras_() { return obtenerCarpetaId_("carpetaObras", "Obras (hojas)", carpetaRaiz_()); }
// A diferencia de las fotos de medicion (todas juntas en "Fotos"), los
// adjuntos del Informe de Supervision (Acta de Inicio, polizas, contrato,
// comprobante de anticipo) son documentos legales que hay que poder ubicar
// por obra -- por eso van en una subcarpeta PROPIA de cada obra, dentro de
// una carpeta raiz aparte.
function carpetaInformesRaiz_() { return obtenerCarpetaId_("carpetaInformesRaiz", "Informes Supervision", carpetaRaiz_()); }
function carpetaInformesObra_(obraId) { return obtenerCarpetaId_("carpetaInformeObra_" + obraId, obraId, carpetaInformesRaiz_()); }

// ---------- Listar y previsualizar presupuestos desde Drive ----------

function listarPresupuestos_() {
  var folder = DriveApp.getFolderById(carpetaPresupuestos_());
  var files = folder.getFiles();
  var lista = [];
  while (files.hasNext()) {
    var f = files.next();
    var nombre = f.getName();
    var ext = nombre.split(".").pop().toLowerCase();
    if (ext === "xls" || ext === "xlsx") {
      lista.push({ id: f.getId(), nombre: nombre, fecha: f.getLastUpdated().toISOString() });
    }
  }
  lista.sort(function (a, b) { return new Date(b.fecha) - new Date(a.fecha); });
  return lista;
}

function convertirYLeer_(fileId) {
  var fileBlob = DriveApp.getFileById(fileId).getBlob();
  var resource = {
    name: "temp_conversion_" + fileId,
    mimeType: MimeType.GOOGLE_SHEETS,
  };
  var convertido = Drive.Files.create(resource, fileBlob);
  try {
    var ss = SpreadsheetApp.openById(convertido.id);
    var hoja = ss.getSheets()[0];
    return hoja.getDataRange().getValues();
  } finally {
    Drive.Files.remove(convertido.id);
  }
}

// ---------- Parser del presupuesto ----------
//
// Los presupuestos de obra de ACUAVALLE (y similares) tienen: unas filas de
// encabezado con los datos del contrato (numero, objeto, contratista,
// supervisor, fechas), y luego UNA O VARIAS tablas de items -una por cada
// "direccion" o frente de obra a intervenir-, cada una empezando con una
// fila "ITEM | DESCRIPCION | UND | CANT. | VR. UNITARIO | VR. PARCIAL"
// seguida de una fila con el nombre de la direccion.
//
// Dentro de cada tabla, las filas pueden ser:
//  - Encabezado de capitulo (ej "1 PRELIMINARES"): tiene ITEM y DESCRIPCION,
//    pero NUNCA tiene VR. UNITARIO (ni CANT. real, aunque a veces trae un 0
//    residual). Estas filas solo sirven para agrupar visualmente.
//  - Item real (ej "1.1 Suministro e instalacion de valla..."): tiene ITEM,
//    DESCRIPCION, y SIEMPRE un VR. UNITARIO numerico (aunque su cantidad
//    presupuestada sea 0). Estos son los que se pueden medir en obra.
// Por eso la regla para distinguir un item real de un capitulo es mirar si
// CANT. y VR. UNITARIO son ambos numeros, no si UND viene lleno (algunos
// presupuestos tienen filas de item real con la columna UND vacia por error
// de captura; esos quedan con unidad "" y hay que completarla a mano en la
// app la primera vez que se usan).

function normalizarTexto_(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

// La columna FotoURL guarda una o varias URLs de Drive separadas por "|"
// (las URLs de Drive nunca traen ese caracter, asi que es un separador
// seguro sin necesidad de JSON ni una hoja aparte).
function separarFotos_(fotoUrlCombinada) {
  var t = normalizarTexto_(fotoUrlCombinada);
  if (!t) return [];
  return t.split("|").map(function (u) { return u.trim(); }).filter(function (u) { return u; });
}

// El error de Apps Script "The string did not match the expected pattern."
// al llamar Utilities.base64Decode() es casi siempre un base64 mal formado:
// longitud que no es multiplo de 4, caracteres fuera del alfabeto base64
// (por ejemplo si en el camino quedaron saltos de linea o espacios), o el
// payload llego incompleto porque la conexion del celular se corto a medio
// subir la foto. Antes de decodificar, se limpia y se valida el formato; si
// no pasa, esa foto en particular se salta (no se pierde la medida completa
// por una sola foto con problemas) y se informa en "fallos" para que el
// front la reporte al usuario.
function base64Valido_(s) {
  var limpio = String(s || "").replace(/\s+/g, "");
  if (!limpio) return null;
  if (limpio.length % 4 !== 0) return null;
  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(limpio)) return null;
  return limpio;
}

function subirFotos_(fotos) {
  var resultado = { urls: [], fallos: [] };
  if (!fotos || !fotos.length) return resultado;
  var carpetaId = carpetaFotos_();
  fotos.forEach(function (f, idx) {
    if (!f || !f.base64) return;
    try {
      var b64 = base64Valido_(f.base64);
      if (!b64) throw new Error("base64 invalido o incompleto");
      var blob = Utilities.newBlob(
        Utilities.base64Decode(b64),
        f.tipo || "image/jpeg",
        (f.nombre || "foto_" + idx) + ".jpg"
      );
      var archivo = DriveApp.getFolderById(carpetaId).createFile(blob);
      archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      resultado.urls.push(archivo.getUrl());
    } catch (err) {
      Logger.log("Foto " + idx + " no se pudo subir: " + err);
      resultado.fallos.push({ nombre: (f && f.nombre) || ("foto_" + idx), error: String(err) });
    }
  });
  return resultado;
}

function normalizarEncabezadoTabla_(v) {
  var t = normalizarTexto_(v).toUpperCase();
  t = t.replace(/[ÁÀÄ]/g, "A").replace(/[ÉÈË]/g, "E").replace(/[ÍÌÏ]/g, "I")
    .replace(/[ÓÒÖ]/g, "O").replace(/[ÚÙÜ]/g, "U");
  return t.replace(/[^A-Z]/g, "");
}

function esNumero_(v) {
  return typeof v === "number" && !isNaN(v);
}

function extraerValorConDosPuntos_(texto) {
  var idx = texto.indexOf(":");
  if (idx === -1) return texto;
  return texto.substring(idx + 1).trim();
}

function parsearPresupuesto_(fileId) {
  var filas = convertirYLeer_(fileId);

  var meta = {
    numeroContrato: "", objeto: "", contratista: "", supervisor: "",
    fecha: "", fechaInicio: "", fechaTerminacion: "",
  };
  for (var f = 0; f < Math.min(10, filas.length); f++) {
    for (var c = 0; c < filas[f].length; c++) {
      var t = normalizarTexto_(filas[f][c]);
      if (!t) continue;
      var up = t.toUpperCase();
      if (up.indexOf("CONTRATO DE OBRA") === 0) {
        var m = /No\.?\s*(.+)/i.exec(t);
        meta.numeroContrato = m ? m[1].trim() : t;
      } else if (up.indexOf("OBJETO") === 0) {
        meta.objeto = extraerValorConDosPuntos_(t);
      } else if (up.indexOf("CONTRATISTA") === 0) {
        meta.contratista = extraerValorConDosPuntos_(t);
      } else if (up.indexOf("SUPERVISOR") === 0) {
        meta.supervisor = extraerValorConDosPuntos_(t);
      } else if (up.indexOf("FECHA DE INICIO") === 0) {
        meta.fechaInicio = extraerValorConDosPuntos_(t);
      } else if (up.indexOf("FECHA DE TERMINACION") === 0 || up.indexOf("FECHA DE TERMINACIÓN") === 0) {
        meta.fechaTerminacion = extraerValorConDosPuntos_(t);
      } else if (up.indexOf("FECHA") === 0 && !meta.fecha) {
        meta.fecha = extraerValorConDosPuntos_(t);
      }
    }
  }

  var encabezados = [];
  for (var i = 0; i < filas.length; i++) {
    var b = normalizarEncabezadoTabla_(filas[i][1]);
    var c2 = normalizarEncabezadoTabla_(filas[i][2]);
    var d = normalizarEncabezadoTabla_(filas[i][3]);
    if (b === "ITEM" && c2 === "DESCRIPCION" && d === "UND") encabezados.push(i);
  }

  var direcciones = [];
  encabezados.forEach(function (hi, idx) {
    var fin = idx + 1 < encabezados.length ? encabezados[idx + 1] : filas.length;
    var nombreDir = hi + 1 < fin ? normalizarTexto_(filas[hi + 1][2]) : ("Dirección " + (idx + 1));
    if (!nombreDir) nombreDir = "Dirección " + (idx + 1);
    var items = [];
    var capituloActual = "";
    for (var i2 = hi + 2; i2 < fin; i2++) {
      var fila = filas[i2] || [];
      var itemVal = normalizarTexto_(fila[1]);
      var descVal = normalizarTexto_(fila[2]);
      var undVal = normalizarTexto_(fila[3]);
      var cantVal = fila[4];
      var vrUnitVal = fila[5];
      if (!itemVal) continue;
      var esItemReal = esNumero_(cantVal) && esNumero_(vrUnitVal);
      if (esItemReal) {
        items.push({
          item: itemVal,
          descripcion: descVal,
          unidad: undVal,
          cantidadPresupuestada: cantVal,
          capitulo: capituloActual,
          // Fila real (1-based) dentro del archivo original donde esta este
          // item, para poder editarlo despues sin ambiguedad (ver
          // actualizarFilaPresupuestoOriginal_). Corresponde a la misma
          // fila en el archivo .xls/.xlsx tal como Drive lo convierte.
          filaOrigen: i2 + 1,
        });
      } else {
        capituloActual = descVal;
      }
    }
    direcciones.push({ nombre: nombreDir, items: items });
  });

  return { meta: meta, direcciones: direcciones };
}

// ---------- Indice de obras (hoja "Obras" de este mismo spreadsheet) ----------

function obtenerHojaIndice_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Obras");
  if (!hoja) {
    hoja = ss.insertSheet("Obras");
    hoja.appendRow(["ObraId", "Nombre", "SpreadsheetId", "NumeroContrato", "Objeto", "Contratista", "Supervisor", "FechaCreacion", "FileIdPresupuesto"]);
  } else if (!normalizarTexto_(hoja.getRange(1, 9).getValue())) {
    // Compatibilidad con hojas "Obras" creadas antes de agregar esta
    // columna: se le pone el encabezado sin tocar las filas existentes (esas
    // obras quedan con FileIdPresupuesto vacio hasta que se vuelvan a crear).
    hoja.getRange(1, 9).setValue("FileIdPresupuesto");
  }
  return hoja;
}

function listarObras_() {
  var hoja = obtenerHojaIndice_();
  var lastRow = hoja.getLastRow();
  if (lastRow < 2) return [];
  var valores = hoja.getRange(2, 1, lastRow - 1, 9).getValues();
  return valores.filter(function (r) { return r[0]; }).map(function (r) {
    return {
      obraId: r[0], nombre: r[1], spreadsheetId: r[2], numeroContrato: r[3],
      objeto: r[4], contratista: r[5], supervisor: r[6], fechaCreacion: r[7],
      fileIdPresupuesto: r[8] || "",
    };
  });
}

function buscarObra_(obraId) {
  var obras = listarObras_();
  for (var i = 0; i < obras.length; i++) {
    if (obras[i].obraId === obraId) return obras[i];
  }
  return null;
}

// ---------- Reportes bajo demanda ----------
//
// "Memoria de Calculo", "Registro Fotografico" y "Ejecucion Real" son solo
// para cuando alguien abre el Google Sheet directamente (imprimir/compartir
// el informe) -- la app nunca lee esas hojas, calcula todo en memoria
// directo desde "Memoria" cada vez (ver leerObra_). Por eso no hace falta
// que esten al dia al instante: en vez de reconstruirlas en el momento del
// guardado (lo que hacia sentir lenta la app), cada accion que cambia datos
// solo marca la obra como "pendiente" (PropertiesService), y esas hojas se
// reconstruyen cuando el usuario presiona "Actualizar reportes" en la app
// (ver actualizarReportesObra_) -- rapido para seguir cargando mediciones,
// y el Sheet se pone al dia justo antes de imprimir/compartir.
//
// (La idea original era un disparador por tiempo 100% automatico
// -actualizarReportesPendientes_ mas abajo- pero choco con un permiso de
// Google que este proyecto no tiene autorizado y que no se pudo conceder
// por un problema de la interfaz del editor de Apps Script. La funcion
// queda por si en el futuro se resuelve eso y se quiere retomar; hoy no la
// llama nada.)
function marcarObraPendiente_(obraId) {
  if (!obraId) return;
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty("obrasPendientes");
  var lista = raw ? JSON.parse(raw) : [];
  if (lista.indexOf(obraId) === -1) {
    lista.push(obraId);
    props.setProperty("obrasPendientes", JSON.stringify(lista));
  }
}

function quitarObraPendiente_(obraId) {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty("obrasPendientes");
  if (!raw) return;
  var lista = JSON.parse(raw).filter(function (id) { return id !== obraId; });
  props.setProperty("obrasPendientes", JSON.stringify(lista));
}

function obraTieneReportesPendientes_(obraId) {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty("obrasPendientes");
  if (!raw) return false;
  return JSON.parse(raw).indexOf(obraId) !== -1;
}

// Boton "🔄 Actualizar reportes" de la app: reconstruye las hojas de
// reporte de ESTA obra puntual (a diferencia de actualizarReportesPendientes_,
// que recorreria todas las pendientes) y limpia su marca de pendiente.
function actualizarReportesObra_(body) {
  var obra = buscarObra_(body.obraId);
  if (!obra) return { ok: false, error: "Obra no encontrada" };
  var ss = SpreadsheetApp.openById(obra.spreadsheetId);
  regenerarMemoriaCalculo_(ss);
  quitarObraPendiente_(body.obraId);
  return { ok: true };
}

// Pensada para un disparador por tiempo (ver nota de arriba: hoy no esta
// conectada a ninguno). Recorre las obras marcadas como pendientes y
// reconstruye sus hojas de reporte. Se quita la marca ANTES de procesar
// (no despues) para que, si algo vuelve a cambiar esa obra mientras esta
// funcion esta corriendo, quede marcada de nuevo para la siguiente pasada
// en vez de perderse.
function actualizarReportesPendientes_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty("obrasPendientes");
  if (!raw) return;
  var lista = JSON.parse(raw);
  if (!lista.length) return;

  props.deleteProperty("obrasPendientes");

  lista.forEach(function (obraId) {
    try {
      var obra = buscarObra_(obraId);
      if (!obra) return;
      var ss = SpreadsheetApp.openById(obra.spreadsheetId);
      regenerarMemoriaCalculo_(ss);
    } catch (err) {
      Logger.log("No se pudo actualizar reportes de la obra " + obraId + ": " + err);
      marcarObraPendiente_(obraId); // reintenta en la siguiente pasada del disparador
    }
  });
}

// ---------- Informe de Supervision (AP2-FO-024): datos manuales + adjuntos ----------
//
// A diferencia del resto de la obra (que se lee del presupuesto/las
// mediciones), estos datos NO existen en ningun lado dentro de la app: son
// los que normalmente salen del Acta de Inicio, las polizas y el contrato.
// Se ingresan una sola vez a mano y quedan guardados en la propia hoja de
// la obra (hoja "InformeSupervision", clave/valor, y hoja "Amparos", tabla)
// para reusarlos en cada informe parcial futuro sin volver a digitarlos.
// Los archivos adjuntos (Acta, polizas, contrato, comprobante de anticipo)
// son solo de referencia/respaldo -- no se leen ni procesan (sin OCR), solo
// se guardan en Drive y su URL queda como un campo mas del informe.

var CAMPOS_INFORME_SUPERVISION_ = [
  "numeroContrato", "objeto", "contratista", "valorInicial", "plazo",
  "fechaActaInicio", "fechaTerminacionInicial", "nuevaFechaTerminacion",
  "polizaCumplimiento", "polizaResponsabilidadCivil", "companiaAseguradora",
  "porcentajeAnticipo", "valorAnticipo",
  "urlActaInicio", "urlPolizas", "urlContrato", "urlComprobanteAnticipo",
];

function hojaInformeSupervision_(ss) {
  var hoja = ss.getSheetByName("InformeSupervision");
  if (!hoja) {
    hoja = ss.insertSheet("InformeSupervision");
    hoja.getRange(1, 1, 1, 2).setValues([["Campo", "Valor"]]).setFontWeight("bold");
    hoja.setColumnWidth(1, 220);
    hoja.setColumnWidth(2, 420);
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function hojaAmparos_(ss) {
  var hoja = ss.getSheetByName("Amparos");
  if (!hoja) {
    hoja = ss.insertSheet("Amparos");
    hoja.getRange(1, 1, 1, 5).setValues([["Tipo", "Porcentaje", "Valor Asegurado", "Vigencia Desde", "Vigencia Hasta"]]).setFontWeight("bold");
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function leerCamposInforme_(ss) {
  var hoja = hojaInformeSupervision_(ss);
  var lastRow = hoja.getLastRow();
  var datos = {};
  if (lastRow > 1) {
    hoja.getRange(2, 1, lastRow - 1, 2).getValues().forEach(function (r) {
      if (r[0]) datos[r[0]] = valorPlano_(r[1]);
    });
  }
  return datos;
}

function leerAmparos_(ss) {
  var hoja = hojaAmparos_(ss);
  var lastRow = hoja.getLastRow();
  if (lastRow < 2) return [];
  return hoja.getRange(2, 1, lastRow - 1, 5).getValues()
    .filter(function (r) { return r.some(function (v) { return v !== ""; }); })
    .map(function (r) {
      return {
        tipo: valorPlano_(r[0]), porcentaje: valorPlano_(r[1]), valorAsegurado: valorPlano_(r[2]),
        vigenciaDesde: valorPlano_(r[3]), vigenciaHasta: valorPlano_(r[4]),
      };
    });
}

// Si Sheets interpreto un texto como "2026-07-18" y lo convirtio solo a un
// objeto Date (pasa facil con columnas donde TODA la fila que se escribe de
// una sola vez parece fecha), lo devuelve como texto plano "YYYY-MM-DD" de
// nuevo -- si no, lo deja tal cual. Sin esto, el campo le llega al
// <input type="date"> del formulario con hora/zona incluida y no lo
// reconoce, asi que el formulario se ve vacio aunque el dato si este
// guardado.
function valorPlano_(v) {
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return v;
}

// Actualiza SOLO los campos que vienen en "camposParciales" (los demas
// quedan tal cual estaban) -- asi subirDocumentoInforme_ puede actualizar
// nada mas su propia URL sin tener que reenviar todo el formulario.
function upsertCamposInforme_(ss, camposParciales) {
  var actuales = leerCamposInforme_(ss);
  Object.keys(camposParciales || {}).forEach(function (k) {
    actuales[k] = camposParciales[k];
  });
  var hoja = hojaInformeSupervision_(ss);
  var lastRow = hoja.getLastRow();
  if (lastRow > 1) hoja.getRange(2, 1, lastRow - 1, 2).clearContent();
  // Se escribe siempre en el mismo orden (CAMPOS_INFORME_SUPERVISION_) para
  // que la hoja quede prolija si alguien la abre directo en Sheets;
  // cualquier clave vieja/desconocida se agrega al final para no perderla.
  var claves = CAMPOS_INFORME_SUPERVISION_.slice();
  Object.keys(actuales).forEach(function (k) { if (claves.indexOf(k) === -1) claves.push(k); });
  var filas = claves
    .filter(function (k) { return actuales[k] !== undefined && actuales[k] !== ""; })
    .map(function (k) { return [k, actuales[k]]; });
  if (filas.length) {
    // Texto plano ANTES de escribir: la columna "Valor" mezcla numeros,
    // fechas y texto libre segun la fila, y sin esto Sheets a veces
    // convierte solo de una las que parecen fecha a un objeto Date real
    // (ver valorPlano_ arriba, que ademas cubre los datos que hayan
    // quedado asi guardados de antes de este arreglo).
    hoja.getRange(2, 1, filas.length, 2).setNumberFormat("@");
    hoja.getRange(2, 1, filas.length, 2).setValues(filas);
  }
  return actuales;
}

function guardarAmparos_(ss, amparos) {
  var hoja = hojaAmparos_(ss);
  var lastRow = hoja.getLastRow();
  if (lastRow > 1) hoja.getRange(2, 1, lastRow - 1, 5).clearContent();
  var filas = (amparos || [])
    .filter(function (a) { return a && (a.tipo || a.valorAsegurado); })
    .map(function (a) { return [a.tipo || "", a.porcentaje || "", a.valorAsegurado || "", a.vigenciaDesde || "", a.vigenciaHasta || ""]; });
  if (filas.length) {
    // Mismo motivo que en upsertCamposInforme_: forzar texto plano para
    // que las columnas de vigencia no queden como fecha real de Sheets.
    hoja.getRange(2, 1, filas.length, 5).setNumberFormat("@");
    hoja.getRange(2, 1, filas.length, 5).setValues(filas);
  }
}

function guardarDatosInforme_(body) {
  var obra = buscarObra_(body.obraId);
  if (!obra) return { ok: false, error: "Obra no encontrada" };
  var ss = SpreadsheetApp.openById(obra.spreadsheetId);
  var datos = upsertCamposInforme_(ss, body.datos || {});
  if (body.amparos) guardarAmparos_(ss, body.amparos);
  return { ok: true, datos: datos, amparos: leerAmparos_(ss) };
}

function obtenerDatosInforme_(obraId) {
  var obra = buscarObra_(obraId);
  if (!obra) return { ok: false, error: "Obra no encontrada" };
  var ss = SpreadsheetApp.openById(obra.spreadsheetId);
  return { ok: true, datos: leerCamposInforme_(ss), amparos: leerAmparos_(ss) };
}

// A que campo de InformeSupervision corresponde la URL de cada tipo de
// adjunto, para que subirDocumentoInforme_ la guarde sola sin tener que
// reenviar el resto del formulario.
var TIPO_A_CAMPO_DOCUMENTO_ = {
  acta_inicio: "urlActaInicio",
  polizas: "urlPolizas",
  contrato: "urlContrato",
  comprobante_anticipo: "urlComprobanteAnticipo",
};

function subirDocumentoInforme_(body) {
  var obra = buscarObra_(body.obraId);
  if (!obra) return { ok: false, error: "Obra no encontrada" };
  var campo = TIPO_A_CAMPO_DOCUMENTO_[body.tipo];
  if (!campo) return { ok: false, error: "Tipo de documento no reconocido: " + body.tipo };
  var b64 = base64Valido_(body.base64);
  if (!b64) return { ok: false, error: "El archivo llegó incompleto o dañado, intenta adjuntarlo de nuevo" };
  try {
    var esPdf = (body.mime || "").indexOf("pdf") !== -1;
    var blob = Utilities.newBlob(
      Utilities.base64Decode(b64),
      body.mime || "application/pdf",
      body.nombre || (body.tipo + (esPdf ? ".pdf" : ".jpg"))
    );
    var carpeta = DriveApp.getFolderById(carpetaInformesObra_(body.obraId));
    var archivo = carpeta.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = archivo.getUrl();
    var ss = SpreadsheetApp.openById(obra.spreadsheetId);
    var camposActualizados = {};
    camposActualizados[campo] = url;
    upsertCamposInforme_(ss, camposActualizados);
    return { ok: true, url: url, campo: campo };
  } catch (err) {
    return { ok: false, error: "No se pudo subir el archivo: " + err };
  }
}

// ---------- Historial de informes generados ----------

function hojaHistorialInformes_(ss) {
  var hoja = ss.getSheetByName("HistorialInformes");
  if (!hoja) {
    hoja = ss.insertSheet("HistorialInformes");
    hoja.getRange(1, 1, 1, 6).setValues([["FechaGeneracion", "Tipo", "NumeroParcial", "FechaDesde", "FechaHasta", "UrlDocx"]]).setFontWeight("bold");
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function leerHistorialInformes_(ss) {
  var hoja = hojaHistorialInformes_(ss);
  var lastRow = hoja.getLastRow();
  if (lastRow < 2) return [];
  return hoja.getRange(2, 1, lastRow - 1, 6).getValues()
    .filter(function (r) { return r.some(function (v) { return v !== ""; }); })
    .map(function (r) {
      return {
        fechaGeneracion: valorPlano_(r[0]), tipo: r[1], numeroParcial: r[2],
        fechaDesde: valorPlano_(r[3]), fechaHasta: valorPlano_(r[4]), urlDocx: r[5],
      };
    });
}

function registrarInformeGenerado_(body) {
  var obra = buscarObra_(body.obraId);
  if (!obra) return { ok: false, error: "Obra no encontrada" };
  var ss = SpreadsheetApp.openById(obra.spreadsheetId);
  var hoja = hojaHistorialInformes_(ss);
  var fechaGeneracion = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var fila = [fechaGeneracion, body.tipo || "", body.numeroParcial || "", body.fechaDesde || "", body.fechaHasta || "", body.urlDocx || ""];
  hoja.getRange(hoja.getLastRow() + 1, 1, 1, 6).setNumberFormat("@").setValues([fila]);
  return { ok: true, historial: leerHistorialInformes_(ss) };
}

function listarHistorialInformes_(obraId) {
  var obra = buscarObra_(obraId);
  if (!obra) return { ok: false, error: "Obra no encontrada" };
  var ss = SpreadsheetApp.openById(obra.spreadsheetId);
  var historial = leerHistorialInformes_(ss);
  // Sugerencia de numero de parcial: 1 + cuantos "Parcial" ya se generaron
  // (el usuario siempre puede escribir otro numero distinto a mano).
  var siguienteParcial = 1 + historial.filter(function (h) { return h.tipo === "Parcial"; }).length;
  return { ok: true, historial: historial, siguienteParcialSugerido: siguienteParcial };
}

// ---------- Balance financiero y resumen de actividades por periodo de corte ----------
//
// A diferencia de "Ejecucion Real" (siempre acumulado desde el inicio de la
// obra, se regenera completa cada vez), esto calcula DOS cortes a la vez
// para un rango de fechas puntual (el periodo que se va a cobrar en este
// informe): lo ejecutado SOLO en ese periodo (para "Valor Acta N"), y lo
// ejecutado acumulado hasta la fecha final del periodo (para "Saldo por
// ejecutar") -- ninguno de los dos existia antes porque hasta ahora todo se
// leia siempre acumulado desde el inicio, sin filtrar por fecha.
function calcularResumenInforme_(body) {
  var obra = buscarObra_(body.obraId);
  if (!obra) return { ok: false, error: "Obra no encontrada" };
  var fechaDesde = normalizarTexto_(body.fechaDesde);
  var fechaHasta = normalizarTexto_(body.fechaHasta);
  if (!fechaDesde || !fechaHasta) return { ok: false, error: "Falta la fecha de inicio o de fin del periodo" };

  var ss = SpreadsheetApp.openById(obra.spreadsheetId);

  var hojaPres = ss.getSheetByName("Presupuesto");
  var ultimaFilaPres = hojaPres.getLastRow();
  var filasPres = ultimaFilaPres > 1 ? hojaPres.getRange(2, 1, ultimaFilaPres - 1, 9).getValues() : [];

  var hojaMemoria = ss.getSheetByName("Memoria");
  var ultimaFilaMemoria = hojaMemoria.getLastRow();
  var filasMemoria = ultimaFilaMemoria > 1 ? hojaMemoria.getRange(2, 1, ultimaFilaMemoria - 1, 14).getValues() : [];

  // Cantidad ejecutada por item, separada en "de este periodo" (fecha entre
  // fechaDesde y fechaHasta, ambas inclusive) y "acumulada a la fecha"
  // (cualquier fecha hasta fechaHasta inclusive, sin importar cuando empezo).
  var ejecutadoPeriodo = {}, ejecutadoAcumulado = {};
  filasMemoria.forEach(function (r) {
    var fechaMedida = String(valorPlano_(r[1]) || "").slice(0, 10);
    if (!fechaMedida) return;
    var tipo = tipoUnidad_(r[5]);
    var m = { longitud: r[6], ancho: r[7], alto: r[8], volumen: r[9], distanciaKm: r[10], cantidad: r[11] };
    var cantidadParcial = calcularCantidadParcial_(tipo, m);
    var clave = normalizarTexto_(r[2]) + "||" + normalizarTexto_(r[3]);
    if (fechaMedida <= fechaHasta) ejecutadoAcumulado[clave] = (ejecutadoAcumulado[clave] || 0) + cantidadParcial;
    if (fechaMedida >= fechaDesde && fechaMedida <= fechaHasta) ejecutadoPeriodo[clave] = (ejecutadoPeriodo[clave] || 0) + cantidadParcial;
  });

  // Precios: los items del presupuesto original vienen del archivo oficial
  // aparte (igual que en regenerarEjecucionReal_); los agregados desde la
  // app (Origen="APP") ya traen su propio VR Unitario en la misma hoja
  // Presupuesto.
  var preciosOficiales = {};
  var fileId = buscarPresupuestoOficialPorContrato_(obra.numeroContrato);
  if (fileId) {
    var direccionesConocidas = {};
    filasPres.forEach(function (r) { if (r[0]) direccionesConocidas[normalizarTexto_(r[0])] = true; });
    leerPresupuestoOficialConPrecios_(fileId, direccionesConocidas).forEach(function (f) {
      if (f.nivel === 3) preciosOficiales[f.direccion + "||" + f.item] = f.vrUnit;
    });
  }

  var items = [];
  var totalContratadoVr = 0, totalEjecutadoPeriodoVr = 0, totalEjecutadoAcumuladoVr = 0;
  filasPres.forEach(function (r) {
    var direccion = normalizarTexto_(r[0]), item = normalizarTexto_(r[2]);
    if (!direccion || !item) return;
    var clave = direccion + "||" + item;
    var esApp = r[8] === "APP";
    var vrUnitario = esApp ? (Number(r[7]) || 0) : (Number(preciosOficiales[clave]) || 0);
    var cantidadContratada = Number(r[5]) || 0;
    var cantPeriodo = ejecutadoPeriodo[clave] || 0;
    var cantAcumulado = ejecutadoAcumulado[clave] || 0;
    var vrPeriodo = cantPeriodo * vrUnitario;
    var vrAcumulado = cantAcumulado * vrUnitario;
    totalContratadoVr += cantidadContratada * vrUnitario;
    totalEjecutadoPeriodoVr += vrPeriodo;
    totalEjecutadoAcumuladoVr += vrAcumulado;
    if (cantidadContratada || cantPeriodo || cantAcumulado) {
      items.push({
        direccion: r[0], item: item, descripcion: normalizarTexto_(r[3]), unidad: normalizarTexto_(r[4]),
        cantidadContratada: cantidadContratada, vrUnitario: vrUnitario,
        cantidadEjecutadaPeriodo: cantPeriodo, vrEjecutadoPeriodo: vrPeriodo,
        cantidadEjecutadaAcumulada: cantAcumulado, vrEjecutadoAcumulado: vrAcumulado,
      });
    }
  });

  // Fotos de las mediciones tomadas DENTRO del periodo (las de este corte),
  // agrupadas luego por direccion/item en el generador del Word.
  var fotos = [];
  filasMemoria.forEach(function (r) {
    var fechaMedida = String(valorPlano_(r[1]) || "").slice(0, 10);
    if (fechaMedida < fechaDesde || fechaMedida > fechaHasta) return;
    separarFotos_(r[12]).forEach(function (url) {
      fotos.push({ direccion: r[2], item: r[3], descripcion: r[4], fotoUrl: url, fecha: fechaMedida });
    });
  });

  // Los "items" (y su VR Unitario) representan el COSTO DIRECTO de cada
  // actividad -- lo mismo que "TOTAL COSTOS DIRECTOS TCD" en la hoja
  // "Ejecucion Real" (ver regenerarEjecucionReal_). Lo que realmente se
  // cobra en un acta es el COSTO TOTAL (TCD + Administracion 30,4% +
  // Imprevistos 1% + Utilidad 6%, los mismos porcentajes que ya usa
  // "Ejecucion Real"), no el TCD solo -- por eso se aplica el mismo
  // factor antes de devolver los totales.
  var FACTOR_AIU_ = 1 + 0.304 + 0.01 + 0.06;

  return {
    ok: true, items: items, fotos: fotos,
    totalContratadoVr: totalContratadoVr * FACTOR_AIU_,
    totalEjecutadoPeriodoVr: totalEjecutadoPeriodoVr * FACTOR_AIU_,
    totalEjecutadoAcumuladoVr: totalEjecutadoAcumuladoVr * FACTOR_AIU_,
    sinArchivoOficial: !fileId,
  };
}

// Sube el .docx YA GENERADO (por el backend Node, con la libreria docx) a
// la misma carpeta de Drive de esta obra -- a diferencia de
// subirDocumentoInforme_, este no actualiza ningun campo de
// InformeSupervision (un informe generado no es un dato del contrato, es un
// documento de salida; queda registrado aparte en HistorialInformes via
// registrarInformeGenerado_).
function subirInformeGeneradoDocx_(body) {
  var obra = buscarObra_(body.obraId);
  if (!obra) return { ok: false, error: "Obra no encontrada" };
  var b64 = base64Valido_(body.base64);
  if (!b64) return { ok: false, error: "El archivo generado llegó incompleto, intenta de nuevo" };
  try {
    var blob = Utilities.newBlob(
      Utilities.base64Decode(b64),
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      body.nombre || "Informe.docx"
    );
    var carpeta = DriveApp.getFolderById(carpetaInformesObra_(body.obraId));
    var archivo = carpeta.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { ok: true, url: archivo.getUrl() };
  } catch (err) {
    return { ok: false, error: "No se pudo guardar el informe generado: " + err };
  }
}

// ---------- Crear una obra nueva a partir de un presupuesto ----------

function crearObra_(body) {
  var fileId = body.fileId;
  var nombreObra = normalizarTexto_(body.nombreObra) || "Obra sin nombre";
  var datos = parsearPresupuesto_(fileId);

  var ss = SpreadsheetApp.create(nombreObra + " - Memoria de Cálculo");
  var ssId = ss.getId();

  var archivo = DriveApp.getFileById(ssId);
  DriveApp.getFolderById(carpetaObras_()).addFile(archivo);
  DriveApp.getRootFolder().removeFile(archivo);

  var hojaConfig = ss.getSheets()[0];
  hojaConfig.setName("Config");
  hojaConfig.getRange(1, 1, 8, 2).setValues([
    ["Nombre Obra", nombreObra],
    ["Número Contrato", datos.meta.numeroContrato],
    ["Objeto", datos.meta.objeto],
    ["Contratista", datos.meta.contratista],
    ["Supervisor", datos.meta.supervisor],
    ["Fecha", datos.meta.fecha],
    ["Fecha Inicio", datos.meta.fechaInicio],
    ["Fecha Terminación", datos.meta.fechaTerminacion],
  ]);
  hojaConfig.getRange(1, 1, 8, 1).setFontWeight("bold");
  hojaConfig.setColumnWidth(1, 160);
  hojaConfig.setColumnWidth(2, 420);

  var hojaPres = ss.insertSheet("Presupuesto");
  // La columna "Fila Origen" (numero de fila dentro del archivo original en
  // Drive) es lo que permite despues editar un item desde la app y que el
  // cambio se escriba tambien en ese archivo original, sin tener que
  // adivinar cual fila es por texto (ver editarItemPresupuesto_). "VR
  // Unitario" normalmente queda vacia para los items del presupuesto
  // original (su precio vive en el archivo oficial aparte, ver
  // regenerarEjecucionReal_); solo se usa para items agregados manualmente
  // desde la app, que no estan en ese archivo oficial y por eso necesitan
  // su propio precio si se quiere ver su valor en "Ejecucion Real".
  // "Origen" (columna 9) es el marcador confiable de "este item lo agrego
  // el usuario a mano desde el boton + Agregar item nuevo" ("APP") vs
  // "vino del presupuesto original al crear la obra" (vacio). No basta con
  // mirar si "Fila Origen" esta vacia para eso: en obras creadas ANTES de
  // que existiera ese tracking, TODOS los items originales tienen Fila
  // Origen vacia tambien, y sin esta columna aparte terminaban
  // confundiendose con items agregados manualmente (ver
  // regenerarEjecucionReal_ y borrarItemPresupuesto_).
  var filasPres = [["Dirección", "Capítulo", "Item", "Descripción", "Unidad", "Cantidad Presupuestada", "Fila Origen", "VR Unitario", "Origen"]];
  datos.direcciones.forEach(function (d) {
    d.items.forEach(function (it) {
      filasPres.push([d.nombre, it.capitulo, it.item, it.descripcion, it.unidad, it.cantidadPresupuestada, it.filaOrigen || "", "", ""]);
    });
  });
  hojaPres.getRange(1, 1, filasPres.length, 9).setValues(filasPres);
  hojaPres.getRange(1, 1, 1, 9).setFontWeight("bold");
  hojaPres.setFrozenRows(1);

  var hojaMemoria = ss.insertSheet("Memoria");
  hojaMemoria.appendRow(["ID", "FechaHora", "Dirección", "Item", "Descripción", "Unidad",
    "Longitud", "Ancho", "Alto", "Volumen", "DistanciaKm", "Cantidad", "FotoURL", "Observación"]);
  hojaMemoria.getRange(1, 1, 1, 14).setFontWeight("bold");
  hojaMemoria.setFrozenRows(1);

  regenerarMemoriaCalculo_(ss);

  var hojaIndice = obtenerHojaIndice_();
  var obraId = Utilities.getUuid();
  hojaIndice.appendRow([obraId, nombreObra, ssId, datos.meta.numeroContrato, datos.meta.objeto,
    datos.meta.contratista, datos.meta.supervisor, new Date().toISOString(), fileId]);

  return { ok: true, obraId: obraId, totalDirecciones: datos.direcciones.length };
}

function borrarObra_(body) {
  var hoja = obtenerHojaIndice_();
  var lastRow = hoja.getLastRow();
  if (lastRow < 2) return { ok: false, error: "No hay obras" };
  var ids = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === body.obraId) {
      hoja.deleteRow(i + 2);
      return { ok: true };
    }
  }
  return { ok: false, error: "Obra no encontrada" };
}

// ---------- Leer una obra (config + presupuesto con totales + medidas) ----------

// Nota sobre rendimiento: esta funcion YA NO reconstruye "Memoria de
// Cálculo" / "Registro Fotográfico" / "Ejecución Real" cada vez que se abre
// la obra (antes lo hacia, y era la causa de que guardar una medida se
// sintiera lento: guardar + recargar terminaba reconstruyendo esas 3 hojas
// completas dos veces). Esos calculos (cantidadEjecutada, totales) se hacen
// aqui mismo en memoria a partir de la hoja "Memoria" (la fuente real de
// datos), asi que la app nunca dependio de que esas hojas estuvieran
// reconstruidas para mostrar informacion correcta -- son solo para cuando
// alguien abre el Google Sheet directamente (para imprimir/compartir el
// informe). Por eso, en vez de reconstruirlas aqui (bloqueando al usuario),
// se marca la obra como "pendiente" (ver marcarObraPendiente_) y el usuario
// las pone al dia cuando quiera con el boton "Actualizar reportes" de la
// app (accion actualizar_reportes -> actualizarReportesObra_).
function leerObra_(obraId) {
  var obra = buscarObra_(obraId);
  if (!obra) throw new Error("Obra no encontrada");

  var ss = SpreadsheetApp.openById(obra.spreadsheetId);

  var hojaPres = ss.getSheetByName("Presupuesto");
  var ultimaFilaPres = hojaPres.getLastRow();
  var filasPres = ultimaFilaPres > 1 ? hojaPres.getRange(2, 1, ultimaFilaPres - 1, 9).getValues() : [];

  var hojaMemoria = ss.getSheetByName("Memoria");
  var ultimaFilaMemoria = hojaMemoria.getLastRow();
  var filasMemoria = ultimaFilaMemoria > 1 ? hojaMemoria.getRange(2, 1, ultimaFilaMemoria - 1, 14).getValues() : [];

  var totales = {};
  var medidas = filasMemoria.map(function (r) {
    var tipo = tipoUnidad_(r[5]);
    var m = { longitud: r[6], ancho: r[7], alto: r[8], volumen: r[9], distanciaKm: r[10], cantidad: r[11] };
    var cantidadParcial = calcularCantidadParcial_(tipo, m);
    var clave = r[2] + "||" + r[3];
    totales[clave] = (totales[clave] || 0) + cantidadParcial;
    return {
      id: r[0], fechaHora: r[1], direccion: r[2], item: r[3], descripcion: r[4], unidad: r[5],
      longitud: r[6], ancho: r[7], alto: r[8], volumen: r[9], distanciaKm: r[10],
      cantidad: r[11], cantidadParcial: cantidadParcial, fotoUrl: r[12], fotosUrls: separarFotos_(r[12]), observacion: r[13],
    };
  });

  var direccionesMap = {};
  var ordenDirecciones = [];
  filasPres.forEach(function (r) {
    var direccion = r[0];
    if (!direccionesMap[direccion]) {
      direccionesMap[direccion] = [];
      ordenDirecciones.push(direccion);
    }
    var clave = r[0] + "||" + r[2];
    direccionesMap[direccion].push({
      capitulo: r[1], item: r[2], descripcion: r[3], unidad: r[4],
      cantidadPresupuestada: r[5], cantidadEjecutada: totales[clave] || 0,
      filaOrigen: r[6] || "", vrUnitario: r[7] || "", origen: r[8] || "",
    });
  });
  var direcciones = ordenDirecciones.map(function (nombre) {
    return { nombre: nombre, items: direccionesMap[nombre] };
  });

  return {
    obra: obra, direcciones: direcciones, medidas: medidas,
    reportesDesactualizados: obraTieneReportesPendientes_(obraId),
  };
}

// ---------- Medidas (memoria de calculo) ----------

function guardarMedida_(body) {
  var obra = buscarObra_(body.obraId);
  if (!obra) return { ok: false, error: "Obra no encontrada" };

  var ss = SpreadsheetApp.openById(obra.spreadsheetId);
  var hoja = ss.getSheetByName("Memoria");

  var subida = subirFotos_(body.fotos);

  var id = Utilities.getUuid();
  var ahora = new Date();
  var fechaHora = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  var fila = [
    id, fechaHora, body.direccion || "", body.item || "", body.descripcion || "", body.unidad || "",
    body.longitud || "", body.ancho || "", body.alto || "", body.volumen || "", body.distanciaKm || "",
    Number(body.cantidad) || 0, subida.urls.join("|"), body.observacion || "",
  ];
  hoja.appendRow(fila);
  // No se reconstruyen aqui "Memoria de Calculo" / "Registro Fotografico" /
  // "Ejecucion Real" (eso es lo que hacia sentir lento el guardado). En vez
  // de eso, se marca la obra como pendiente y el usuario las actualiza
  // cuando quiera con el boton "Actualizar reportes" de la app.
  marcarObraPendiente_(body.obraId);

  return { ok: true, id: id, fechaHora: fechaHora, fotosUrls: subida.urls, fotosFallidas: subida.fallos };
}

// Crea varias medidas nuevas de una sola vez en el item destino (direccion +
// item + descripcion + unidad vienen del item ACTUAL, no del origen), a
// partir de una lista ya preparada por el frontend (ver mapearCamposEntreTipos
// en index.html): cada entrada trae solo los campos que tienen sentido para
// la unidad del item destino (por ejemplo, si el origen era ML y el destino
// es M2, "ancho" llega vacio a proposito para que el usuario lo complete
// despues editando cada medida). Se usa para el flujo de "copiar
// mediciones de otro item" (mismas domiciliarias repetidas en varios
// items). regenerarMemoriaCalculo_ se llama una sola vez al final, no por
// cada fila, para que copiar 20 domiciliarias no sea 20 veces mas lento que
// guardar una sola medida.
function copiarMedidas_(body) {
  var obra = buscarObra_(body.obraId);
  if (!obra) return { ok: false, error: "Obra no encontrada" };
  var medidas = body.medidas || [];
  if (!medidas.length) return { ok: false, error: "No hay mediciones para copiar" };

  var ss = SpreadsheetApp.openById(obra.spreadsheetId);
  var hoja = ss.getSheetByName("Memoria");
  var ahora = new Date();
  var fechaHora = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  var idsCreadas = [];
  var filas = medidas.map(function (m) {
    var id = Utilities.getUuid();
    idsCreadas.push(id);
    return [
      id, fechaHora, body.direccion || "", body.item || "", body.descripcion || "", body.unidad || "",
      m.longitud || "", m.ancho || "", m.alto || "", m.volumen || "", m.distanciaKm || "",
      Number(m.cantidad) || 0, "", m.observacion || "",
    ];
  });
  hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, 14).setValues(filas);
  marcarObraPendiente_(body.obraId); // ver nota en guardarMedida_

  return { ok: true, ids: idsCreadas, cantidad: idsCreadas.length };
}

function editarMedida_(body) {
  var obra = buscarObra_(body.obraId);
  if (!obra) return { ok: false, error: "Obra no encontrada" };

  var ss = SpreadsheetApp.openById(obra.spreadsheetId);
  var hoja = ss.getSheetByName("Memoria");
  var lastRow = hoja.getLastRow();
  if (lastRow < 2) return { ok: false, error: "No hay medidas" };

  var ids = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === body.medidaId) {
      var fila = i + 2;
      hoja.getRange(fila, 3).setValue(body.direccion || "");
      hoja.getRange(fila, 4).setValue(body.item || "");
      hoja.getRange(fila, 5).setValue(body.descripcion || "");
      hoja.getRange(fila, 6).setValue(body.unidad || "");
      hoja.getRange(fila, 7).setValue(body.longitud || "");
      hoja.getRange(fila, 8).setValue(body.ancho || "");
      hoja.getRange(fila, 9).setValue(body.alto || "");
      hoja.getRange(fila, 10).setValue(body.volumen || "");
      hoja.getRange(fila, 11).setValue(body.distanciaKm || "");
      hoja.getRange(fila, 12).setValue(Number(body.cantidad) || 0);
      var subida = subirFotos_(body.fotos);
      if (subida.urls.length) {
        var fotosActuales = separarFotos_(hoja.getRange(fila, 13).getValue());
        hoja.getRange(fila, 13).setValue(fotosActuales.concat(subida.urls).join("|"));
      }
      hoja.getRange(fila, 14).setValue(body.observacion || "");
      marcarObraPendiente_(body.obraId); // ver nota en guardarMedida_
      return { ok: true, fotosFallidas: subida.fallos };
    }
  }
  return { ok: false, error: "Medida no encontrada" };
}

function borrarMedida_(body) {
  var obra = buscarObra_(body.obraId);
  if (!obra) return { ok: false, error: "Obra no encontrada" };

  var ss = SpreadsheetApp.openById(obra.spreadsheetId);
  var hoja = ss.getSheetByName("Memoria");
  var lastRow = hoja.getLastRow();
  if (lastRow < 2) return { ok: false, error: "No hay medidas" };

  var ids = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === body.medidaId) {
      hoja.deleteRow(i + 2);
      marcarObraPendiente_(body.obraId); // ver nota en guardarMedida_
      return { ok: true };
    }
  }
  return { ok: false, error: "Medida no encontrada" };
}

// ---------- Agregar un item nuevo al presupuesto de una obra ----------
//
// Para cuando el presupuesto original de Drive se queda corto (un item que
// nunca estuvo ahi, o que se agrego despues en una adicion al contrato).
// Esto NO toca el archivo original en Drive -- a diferencia de editar_item,
// no hay "fila origen" que actualizar porque el item nunca vino de ahi. Solo
// queda agregado en la copia de Presupuesto de esta obra (FilaOrigen vacia),
// asi que si mas adelante se corrige desde el lapiz ✏️, el aviso va a decir
// que no se pudo actualizar el archivo original -- es lo esperado.
function agregarItemPresupuesto_(body) {
  var obra = buscarObra_(body.obraId);
  if (!obra) return { ok: false, error: "Obra no encontrada" };

  var direccion = normalizarTexto_(body.direccion);
  var itemNuevo = normalizarTexto_(body.item);
  if (!direccion || !itemNuevo) return { ok: false, error: "Falta dirección o número de item" };

  var descripcion = normalizarTexto_(body.descripcion);
  var unidad = normalizarTexto_(body.unidad);
  var capitulo = normalizarTexto_(body.capitulo) || "Agregado manualmente";
  var cantidadPresupuestada = (body.cantidadPresupuestada !== "" && body.cantidadPresupuestada != null)
    ? (Number(body.cantidadPresupuestada) || 0) : 0;
  // Opcional: si se da un precio unitario, este item puede aparecer con
  // valor en "Ejecucion Real" ademas de solo la cantidad (ver
  // regenerarEjecucionReal_). Si se deja vacio, igual aparece ahi, solo que
  // sin columnas de valor.
  var vrUnitario = (body.vrUnitario !== "" && body.vrUnitario != null) ? (Number(body.vrUnitario) || "") : "";

  var ss = SpreadsheetApp.openById(obra.spreadsheetId);
  var hojaPres = ss.getSheetByName("Presupuesto");
  var lastRow = hojaPres.getLastRow();

  // Se busca la ULTIMA fila que ya pertenece a esta direccion, para insertar
  // el item nuevo justo despues (dentro de su propio bloque), en vez de
  // pegarlo al final de toda la hoja -- si no, Memoria de Calculo terminaria
  // mostrando el nombre de la direccion duplicado (su bloque original mas
  // arriba, y otro aparte solo con el item nuevo al fondo de la hoja).
  var filaInsercion = -1; // 1-based; -1 = no se encontro esa direccion
  if (lastRow > 1) {
    var valores = hojaPres.getRange(2, 1, lastRow - 1, 3).getValues();
    for (var i = 0; i < valores.length; i++) {
      if (normalizarTexto_(valores[i][0]) === direccion) {
        if (normalizarTexto_(valores[i][2]) === itemNuevo) {
          return { ok: false, error: "Ya existe un item \"" + itemNuevo + "\" en esa dirección. Usa el lápiz ✏️ para editarlo en vez de crear uno nuevo." };
        }
        filaInsercion = i + 2; // ultima fila vista de esta direccion (se sigue actualizando al recorrer)
      }
    }
  }

  // "APP" en la columna Origen es lo que despues distingue este item de uno
  // que vino del presupuesto original (ver nota en crearObra_).
  var filaNueva = [direccion, capitulo, itemNuevo, descripcion, unidad, cantidadPresupuestada, "", vrUnitario, "APP"];
  var filaDestino;
  if (filaInsercion === -1) {
    // No habia ningun item de esta direccion todavia (caso raro): se agrega
    // al final de toda la hoja, igual que antes.
    hojaPres.appendRow(filaNueva);
    filaDestino = hojaPres.getLastRow();
  } else {
    hojaPres.insertRowAfter(filaInsercion);
    filaDestino = filaInsercion + 1;
    hojaPres.getRange(filaDestino, 1, 1, 9).setValues([filaNueva]);
  }
  // Texto plano en la columna Item para que numeros como "9.10" no pierdan
  // el cero final (Sheets los convertiria en el numero 9.1, chocando con
  // un item "9.1" real -- el mismo problema que resuelve editar_item).
  hojaPres.getRange(filaDestino, 3).setNumberFormat("@").setValue(itemNuevo);
  marcarObraPendiente_(body.obraId);

  return { ok: true };
}

// Borra un item de la hoja "Presupuesto" de la obra. Por seguridad SOLO
// permite borrar items marcados como Origen="APP" (los agregados
// manualmente desde el boton + Agregar item nuevo) -- un item que si vino
// del presupuesto original no se puede borrar aqui, para no perder de
// vista algo que en realidad esta en el contrato. Ojo: NO basta con mirar
// si FilaOrigen esta vacia para eso -- en obras creadas antes de que
// existiera esa columna, TODOS los items originales la tienen vacia
// tambien, y por eso el chequeo real es la columna Origen (ver nota en
// crearObra_/agregarItemPresupuesto_).
// Las mediciones ya guardadas contra ese item en la hoja "Memoria" NO se
// borran (evita perder datos por accidente); solo dejan de aparecer en la
// app y en los reportes porque ya no hay item con el que cruzarlas.
function borrarItemPresupuesto_(body) {
  var obra = buscarObra_(body.obraId);
  if (!obra) return { ok: false, error: "Obra no encontrada" };

  var direccion = normalizarTexto_(body.direccion);
  var item = normalizarTexto_(body.item);
  if (!direccion || !item) return { ok: false, error: "Falta dirección o item" };
  var descripcionActual = normalizarTexto_(body.descripcionActual);
  var cantidadActualRaw = body.cantidadPresupuestadaActual;
  var tieneCantidadActual = cantidadActualRaw !== undefined && cantidadActualRaw !== null && cantidadActualRaw !== "";
  var cantidadActual = tieneCantidadActual ? Number(cantidadActualRaw) : null;

  var ss = SpreadsheetApp.openById(obra.spreadsheetId);
  var hojaPres = ss.getSheetByName("Presupuesto");
  var lastRow = hojaPres.getLastRow();
  if (lastRow < 2) return { ok: false, error: "Este presupuesto no tiene items" };

  var valores = hojaPres.getRange(2, 1, lastRow - 1, 9).getValues();
  var candidatos = [];
  for (var i = 0; i < valores.length; i++) {
    if (normalizarTexto_(valores[i][0]) === direccion && normalizarTexto_(valores[i][2]) === item) candidatos.push(i);
  }

  // Misma cascada de desempate que editarItemPresupuesto_ (ver ese
  // comentario): descripcion actual primero, cantidad presupuestada actual
  // si la descripcion tambien quedo identica.
  var indiceElegido = -1;
  if (candidatos.length === 1) {
    indiceElegido = candidatos[0];
  } else if (candidatos.length > 1) {
    var porDescripcion = candidatos.filter(function (i) { return normalizarTexto_(valores[i][3]) === descripcionActual; });
    var restantes = porDescripcion.length ? porDescripcion : candidatos;
    if (restantes.length === 1) {
      indiceElegido = restantes[0];
    } else if (tieneCantidadActual) {
      var porCantidad = restantes.filter(function (i) { return Number(valores[i][5]) === cantidadActual; });
      if (porCantidad.length === 1) indiceElegido = porCantidad[0];
    }
  }

  if (indiceElegido === -1) {
    if (candidatos.length > 1) {
      return { ok: false, error: "Hay más de un item \"" + item + "\" en esta dirección y no se pudo identificar cuál exactamente. Recarga la app e inténtalo de nuevo." };
    }
    return { ok: false, error: "No se encontró ese item en el presupuesto de la obra" };
  }

  var origen = valores[indiceElegido][8];
  if (origen !== "APP") {
    return { ok: false, error: "Este item viene del presupuesto original, así que no se puede eliminar desde la app (solo los agregados manualmente). Usa el lápiz ✏️ si necesitas corregirlo." };
  }

  hojaPres.deleteRow(indiceElegido + 2);
  marcarObraPendiente_(body.obraId);

  return { ok: true };
}

// ---------- Editar un item del presupuesto (unidad/descripcion/numero) ----------
//
// Permite corregir desde la app un item que quedo mal capturado al leer el
// presupuesto original (tipico: la columna UND vino vacia por error de
// captura, o el numero de item quedo mal escrito). Actualiza tres lugares
// para que todo quede consistente:
//  1. La hoja "Presupuesto" de la obra (la copia que usa la app).
//  2. Las filas de "Memoria" que ya tenian mediciones cargadas contra el
//     item viejo, para que no queden huerfanas (el cruce con el
//     presupuesto se hace por direccion+item; si no se migran, esas
//     mediciones dejarian de sumar contra el item corregido).
//  3. El archivo original en Drive (el .xls/.xlsx que se uso para crear la
//     obra), en la fila exacta de donde salio ese item -- solo si se
//     conoce esa fila (obras creadas despues de agregar este seguimiento;
//     ver parsearPresupuesto_ / crearObra_) y si el archivo todavia existe.
// El paso 3 es el unico que puede fallar sin perder el resto del cambio
// (por eso va envuelto en try/catch): si falla, el item queda igual
// corregido en la app, y se devuelve un aviso de que el archivo original no
// se pudo actualizar.
function editarItemPresupuesto_(body) {
  var obra = buscarObra_(body.obraId);
  if (!obra) return { ok: false, error: "Obra no encontrada" };

  var direccion = normalizarTexto_(body.direccion);
  var itemViejo = normalizarTexto_(body.itemViejo);
  if (!direccion || !itemViejo) return { ok: false, error: "Falta dirección o item a editar" };

  var itemNuevo = normalizarTexto_(body.item) || itemViejo;
  var descripcionNueva = body.descripcion != null && normalizarTexto_(body.descripcion) ? normalizarTexto_(body.descripcion) : null;
  var unidadNueva = body.unidad != null ? normalizarTexto_(body.unidad) : null;
  var cantidadNueva = (body.cantidadPresupuestada !== undefined && body.cantidadPresupuestada !== null && body.cantidadPresupuestada !== "")
    ? (Number(body.cantidadPresupuestada) || 0) : null;
  var vrUnitarioNuevo = (body.vrUnitario !== undefined && body.vrUnitario !== null && body.vrUnitario !== "")
    ? (Number(body.vrUnitario) || 0) : null;
  // Solo para corregir items viejos agregados desde la app ANTES de que
  // existiera la columna Origen (por eso quedaron sin el marcador "APP" y
  // no se reconocen en regenerarEjecucionReal_/borrarItemPresupuesto_):
  // permite ponerles el marcador ahora, sin tener que borrarlos y
  // volverlos a crear. Solo acepta "APP" a proposito (no sirve para
  // ocultar un item real como si fuera agregado).
  var origenNuevo = (body.origen === "APP") ? "APP" : null;
  // La descripcion y cantidad presupuestada ACTUALES del item que el
  // usuario abrio en la app (no las nuevas) -- sirven para desempatar
  // cuando hay mas de una fila con el mismo numero de item en la misma
  // direccion (ver comentario mas abajo).
  var descripcionActual = normalizarTexto_(body.descripcionActual);
  var cantidadActualRaw = body.cantidadPresupuestadaActual;
  var tieneCantidadActual = cantidadActualRaw !== undefined && cantidadActualRaw !== null && cantidadActualRaw !== "";
  var cantidadActual = tieneCantidadActual ? Number(cantidadActualRaw) : null;

  var ss = SpreadsheetApp.openById(obra.spreadsheetId);
  var hojaPres = ss.getSheetByName("Presupuesto");
  var lastRow = hojaPres.getLastRow();
  if (lastRow < 2) return { ok: false, error: "Este presupuesto no tiene items" };

  var valores = hojaPres.getRange(2, 1, lastRow - 1, 9).getValues();

  // Puede haber mas de una fila con el mismo TEXTO de item en la misma
  // direccion sin ser un error de captura real: es tipico que un numero
  // como "7.10" en el presupuesto original quede guardado por Google
  // Sheets como el numero 7.1 (le recorta el cero final), chocando
  // visualmente con un "7.1" que si es otro item distinto. Si hay mas de
  // un candidato, se desempata en cascada: primero por la descripcion
  // ACTUAL (el frontend ya sabe cual abrio el usuario, no solo el numero);
  // si las descripciones tambien quedaron identicas por el mismo error de
  // captura (pasa: a veces hasta la descripcion se copio mal en el
  // presupuesto original), se desempata por la cantidad presupuestada
  // actual, que casi siempre SI es distinta entre los dos.
  var candidatos = [];
  for (var i = 0; i < valores.length; i++) {
    if (normalizarTexto_(valores[i][0]) === direccion && normalizarTexto_(valores[i][2]) === itemViejo) {
      candidatos.push(i);
    }
  }

  var indiceElegido = -1;
  if (candidatos.length === 1) {
    indiceElegido = candidatos[0];
  } else if (candidatos.length > 1) {
    var porDescripcion = candidatos.filter(function (i) { return normalizarTexto_(valores[i][3]) === descripcionActual; });
    var restantes = porDescripcion.length ? porDescripcion : candidatos;
    if (restantes.length === 1) {
      indiceElegido = restantes[0];
    } else if (tieneCantidadActual) {
      var porCantidad = restantes.filter(function (i) { return Number(valores[i][5]) === cantidadActual; });
      if (porCantidad.length === 1) indiceElegido = porCantidad[0];
    }
  }

  if (indiceElegido === -1) {
    if (candidatos.length > 1) {
      return {
        ok: false,
        error: "Hay " + candidatos.length + " items \"" + itemViejo + "\" en esta dirección y no se pudo identificar cuál exactamente " +
          "(posiblemente uno de ellos en realidad es otro número, como \"" + itemViejo + "0\", que Google Sheets recorta el cero final). " +
          "Recarga la app y vuelve a intentarlo; si el problema sigue, avísame.",
      };
    }
    return { ok: false, error: "No se encontró ese item en el presupuesto de la obra" };
  }

  var filaEncontrada = indiceElegido + 2;
  var filaOrigen = valores[indiceElegido][6];

  // setNumberFormat("@") ANTES de escribir el numero de item fuerza la
  // celda a texto plano, para que Sheets no le recorte ceros finales ni lo
  // convierta en numero (la misma causa del problema de arriba) si el
  // numero nuevo tambien "parece" numerico.
  hojaPres.getRange(filaEncontrada, 3).setNumberFormat("@").setValue(itemNuevo);
  if (descripcionNueva !== null) hojaPres.getRange(filaEncontrada, 4).setValue(descripcionNueva);
  if (unidadNueva !== null) hojaPres.getRange(filaEncontrada, 5).setValue(unidadNueva);
  if (cantidadNueva !== null) hojaPres.getRange(filaEncontrada, 6).setValue(cantidadNueva);
  if (vrUnitarioNuevo !== null) hojaPres.getRange(filaEncontrada, 8).setValue(vrUnitarioNuevo);
  if (origenNuevo !== null) hojaPres.getRange(filaEncontrada, 9).setValue(origenNuevo);

  var hojaMemoria = ss.getSheetByName("Memoria");
  var lastRowMem = hojaMemoria.getLastRow();
  if (lastRowMem > 1) {
    var valoresMem = hojaMemoria.getRange(2, 1, lastRowMem - 1, 14).getValues();
    for (var j = 0; j < valoresMem.length; j++) {
      if (normalizarTexto_(valoresMem[j][2]) !== direccion || normalizarTexto_(valoresMem[j][3]) !== itemViejo) continue;
      // Mismo cuidado que arriba: si el numero de item era ambiguo, solo se
      // migran las mediciones cuya descripcion coincide con la del item que
      // se esta editando, para no arrastrar mediciones que en realidad son
      // del OTRO item con el mismo numero.
      if (candidatos.length > 1 && normalizarTexto_(valoresMem[j][4]) !== descripcionActual) continue;
      var filaMem = j + 2;
      hojaMemoria.getRange(filaMem, 4).setNumberFormat("@").setValue(itemNuevo);
      if (descripcionNueva !== null) hojaMemoria.getRange(filaMem, 5).setValue(descripcionNueva);
      if (unidadNueva !== null) hojaMemoria.getRange(filaMem, 6).setValue(unidadNueva);
    }
  }

  marcarObraPendiente_(body.obraId); // ver nota en guardarMedida_

  var resultado = { ok: true, masterActualizado: false, avisoMaster: "" };
  if (obra.fileIdPresupuesto && filaOrigen) {
    try {
      actualizarFilaPresupuestoOriginal_(obra.fileIdPresupuesto, Number(filaOrigen), {
        item: itemNuevo, descripcion: descripcionNueva, unidad: unidadNueva,
      });
      resultado.masterActualizado = true;
    } catch (errMaster) {
      Logger.log("No se pudo actualizar el archivo original: " + errMaster);
      resultado.avisoMaster = "El item se corrigió en la obra, pero no se pudo actualizar el archivo original en Drive: " + errMaster;
    }
  } else {
    resultado.avisoMaster = "Esta obra no tiene ligado el archivo original (se creó antes de esta función), así que el cambio solo quedó en la app.";
  }

  return resultado;
}

// Reescribe una fila especifica del archivo de presupuesto original (el
// .xls/.xlsx que el usuario subio a la carpeta "Presupuestos" de Drive): lo
// convierte temporalmente a Google Sheets (igual que convertirYLeer_), edita
// la fila indicada, lo vuelve a exportar como Excel y sobreescribe el
// archivo original -- conserva el mismo ID, nombre, carpeta y permisos de
// Drive, solo cambia el contenido.
//
// OJO: esto asume que las filas del archivo no se han movido desde que se
// creo la obra (por ejemplo, insertando o borrando filas a mano en Excel
// despues de crear la obra). Si eso paso, esta funcion editaria la fila
// equivocada -- por eso las correcciones de item/unidad se deben hacer
// siempre desde la app una vez la obra ya existe, y no a mano en el archivo
// original.
function actualizarFilaPresupuestoOriginal_(fileId, filaOrigen, cambios) {
  if (!filaOrigen || filaOrigen < 1) throw new Error("Fila de origen inválida");

  var archivoOriginal = DriveApp.getFileById(fileId);
  var fileBlob = archivoOriginal.getBlob();
  var resource = { name: "temp_edicion_" + fileId, mimeType: MimeType.GOOGLE_SHEETS };
  var convertido = Drive.Files.create(resource, fileBlob);

  try {
    var ssTemp = SpreadsheetApp.openById(convertido.id);
    var hojaTemp = ssTemp.getSheets()[0];
    if (hojaTemp.getLastRow() < filaOrigen) {
      throw new Error("El archivo original ya no tiene esa fila (¿cambió su estructura?)");
    }
    if (cambios.item) hojaTemp.getRange(filaOrigen, 2).setValue(cambios.item);
    if (cambios.descripcion) hojaTemp.getRange(filaOrigen, 3).setValue(cambios.descripcion);
    if (cambios.unidad) hojaTemp.getRange(filaOrigen, 4).setValue(cambios.unidad);
    SpreadsheetApp.flush();

    var blobExportado = DriveApp.getFileById(convertido.id).getAs(MimeType.MICROSOFT_EXCEL);
    Drive.Files.update({}, fileId, blobExportado);
  } finally {
    try { Drive.Files.remove(convertido.id); } catch (eLimpieza) { /* no critico */ }
  }
}

// Reconstruye por completo la hoja "Memoria de Cálculo": para cada
// direccion/capitulo/item que tenga al menos una medida cargada, arma una
// mini-tabla con una fila por medida (columnas Foto/Descripcion/Longitud/
// Ancho/Alto/Volumen/Km/Cantidad, cada una llena solo si aplica segun la
// unidad) y una fila de TOTAL al final sumando la columna Cantidad. Se
// vuelve a armar desde cero cada vez (se borra y se crea de nuevo la hoja)
// para no arrastrar fusiones de celdas ni formato de una version anterior
// con mas o menos filas.
//
// Las fotos NO se incrustan aqui (una imagen grande dentro de esta tabla
// corre las filas y columnas de todo lo demas). En su lugar, cada medida
// con foto queda con una referencia "Imagen N" con hipervinculo hacia su
// fila exacta en la hoja "Registro Fotografico", que es donde se ve la
// foto en grande.
function regenerarMemoriaCalculo_(ss) {
  var _t = { inicio: Date.now() };
  var NOMBRE_HOJA = "Memoria de Cálculo";
  var COLS = 9;

  var hojaPres = ss.getSheetByName("Presupuesto");
  var ultimaFilaPres = hojaPres.getLastRow();
  var filasPres = ultimaFilaPres > 1 ? hojaPres.getRange(2, 1, ultimaFilaPres - 1, 7).getValues() : [];

  var hojaMemoria = ss.getSheetByName("Memoria");
  var ultimaFilaMemoria = hojaMemoria.getLastRow();
  var filasMemoria = ultimaFilaMemoria > 1 ? hojaMemoria.getRange(2, 1, ultimaFilaMemoria - 1, 14).getValues() : [];
  _t.leerHojas = Date.now();

  var medidasPorItem = {};
  filasMemoria.forEach(function (r) {
    var clave = r[2] + "||" + r[3];
    if (!medidasPorItem[clave]) medidasPorItem[clave] = [];
    medidasPorItem[clave].push({
      longitud: r[6], ancho: r[7], alto: r[8], volumen: r[9], distanciaKm: r[10],
      cantidad: r[11], fotoUrl: r[12], observacion: r[13],
    });
  });

  var totalPorItemClave = {};
  filasPres.forEach(function (r) {
    var clave = r[0] + "||" + r[2];
    var medidas = medidasPorItem[clave];
    if (!medidas || !medidas.length) return;
    var tipo = tipoUnidad_(r[4]);
    var suma = 0;
    medidas.forEach(function (m) { suma += calcularCantidadParcial_(tipo, m); });
    totalPorItemClave[clave] = suma;
  });
  _t.calcularTotales = Date.now();

  var hojaVieja = ss.getSheetByName(NOMBRE_HOJA);
  if (hojaVieja) ss.deleteSheet(hojaVieja);
  var hoja = ss.insertSheet(NOMBRE_HOJA);
  _t.borrarCrearHoja = Date.now();

  // Antes, cada fila (cada medida, cada encabezado de item/capitulo/
  // direccion) se escribia con varias llamadas individuales a la API de
  // Sheets (setValue + merge + setFontWeight + setBackground...). Con
  // decenas de mediciones eso son cientos de llamadas, y cada una tiene su
  // propio costo de red/ejecucion -- medido: mas de 20 segundos para
  // reconstruir esta hoja con solo 81 mediciones. Ahora se arma todo en
  // memoria (valores, formulas, colores, negritas) fila por fila igual que
  // antes, pero se aplica al final con un puñado de llamadas en bloque
  // sobre el rango completo. Los merge() no se pueden agrupar en Sheets
  // (cada uno es una operacion propia), pero son muchos menos que las
  // celdas individuales que se estaban escribiendo antes.
  var valores = [];
  var formulasCol9 = [];
  var backgrounds = [];
  var negritas = [];
  var coloresTexto = [];
  var alineaciones = [];
  var merges = [];
  var fotos = [];
  var pendientesFoto = [];

  function filaVacia() { return new Array(COLS).fill(""); }

  // Agrega una fila a los arrays en memoria y devuelve su numero de fila
  // (1-based) dentro de la hoja, para poder referenciarla en formulas o
  // fusiones despues.
  function agregarFila(vals, opts) {
    opts = opts || {};
    valores.push(vals);
    formulasCol9.push("");
    backgrounds.push(new Array(COLS).fill(opts.bg || "#ffffff"));
    negritas.push(new Array(COLS).fill(opts.bold ? "bold" : "normal"));
    coloresTexto.push(new Array(COLS).fill(opts.color || "#000000"));
    alineaciones.push(new Array(COLS).fill(opts.align || "left"));
    return valores.length;
  }

  var direccionActual = null;
  var capituloActual = null;

  filasPres.forEach(function (r) {
    var direccion = r[0], capitulo = r[1], item = r[2], descripcion = r[3], unidad = r[4];
    var clave = direccion + "||" + item;
    var medidas = medidasPorItem[clave];
    if (!medidas || !medidas.length) return;

    if (direccion !== direccionActual) {
      direccionActual = direccion;
      capituloActual = null;
      var valsDir = filaVacia(); valsDir[0] = direccion;
      var filaDir = agregarFila(valsDir, { bg: "#1F4E78", bold: true, color: "#FFFFFF" });
      merges.push({ fila: filaDir, columnas: COLS });
    }
    if (capitulo !== capituloActual) {
      capituloActual = capitulo;
      var valsCap = filaVacia(); valsCap[0] = capitulo;
      var filaCap = agregarFila(valsCap, { bg: "#DCE6F1", bold: true });
      merges.push({ fila: filaCap, columnas: COLS });
    }

    var valsItem = filaVacia();
    valsItem[0] = item + " — " + descripcion + " (" + (unidad || "sin unidad") + ")";
    var filaItem = agregarFila(valsItem, { bold: true });
    merges.push({ fila: filaItem, columnas: COLS });

    agregarFila(
      ["Foto", "Descripción / Observación", "Longitud", "Ancho", "Alto", "Volumen", "Km", "Cantidad", "Cantidad Parcial"],
      { bg: "#F0F0F0", bold: true }
    );

    var tipo = tipoUnidad_(unidad);
    var filaInicioMedidas = valores.length + 1;
    medidas.forEach(function (m) {
      var filaMedida = agregarFila([
        "", m.observacion || "", m.longitud || "", m.ancho || "", m.alto || "", m.volumen || "", m.distanciaKm || "", m.cantidad || 0, "",
      ]);
      formulasCol9[filaMedida - 1] = formulaCantidadParcial_(tipo, filaMedida);
      var urlsMedida = separarFotos_(m.fotoUrl);
      if (urlsMedida.length) {
        var indiceInicio = fotos.length;
        var numeros = [];
        urlsMedida.forEach(function (url) {
          var numero = fotos.length + 1;
          numeros.push(numero);
          fotos.push({
            numero: numero, direccion: direccion, item: item, descripcion: descripcion,
            observacion: m.observacion || "", fotoUrl: url,
          });
        });
        pendientesFoto.push({ fila: filaMedida, indiceFoto: indiceInicio, numeros: numeros });
      }
    });
    var filaFinMedidas = valores.length;

    // El TOTAL solo suma la columna "Cantidad Parcial" (el resultado final).
    // La columna "Cantidad" es el multiplicador manual por fila (cuantos
    // elementos iguales), sumarla no tiene sentido.
    var valsTotal = filaVacia(); valsTotal[0] = "TOTAL";
    var filaTotal = agregarFila(valsTotal, { bold: true, align: "right" });
    formulasCol9[filaTotal - 1] = "=SUM(I" + filaInicioMedidas + ":I" + filaFinMedidas + ")";
    merges.push({ fila: filaTotal, columnas: 8 });

    agregarFila(filaVacia()); // fila en blanco antes del siguiente item
  });
  _t.construirArrays = Date.now();

  var totalFilas = valores.length;
  if (totalFilas > 0) {
    hoja.getRange(1, 1, totalFilas, COLS).setValues(valores);
    hoja.getRange(1, 1, totalFilas, COLS).setBackgrounds(backgrounds);
    hoja.getRange(1, 1, totalFilas, COLS).setFontWeights(negritas);
    hoja.getRange(1, 1, totalFilas, COLS).setFontColors(coloresTexto);
    hoja.getRange(1, 1, totalFilas, COLS).setHorizontalAlignments(alineaciones);
    hoja.getRange(1, 9, totalFilas, 1).setFormulas(formulasCol9.map(function (f) { return [f]; }));
    _t.escriturasEnBloque = Date.now();
    merges.forEach(function (mrg) {
      hoja.getRange(mrg.fila, 1, 1, mrg.columnas).merge();
    });
    _t.merges = Date.now();
  } else {
    _t.escriturasEnBloque = Date.now();
    _t.merges = Date.now();
  }

  hoja.setColumnWidth(1, 100);
  hoja.setColumnWidth(2, 220);
  for (var c = 3; c <= 7; c++) hoja.setColumnWidth(c, 80);
  hoja.setColumnWidth(8, 90);
  hoja.setColumnWidth(9, 110);
  hoja.setFrozenRows(0);
  _t.anchoColumnas = Date.now();

  var meta = leerMetaContrato_(ss);
  var gidFotografico = regenerarRegistroFotografico_(ss, fotos, meta);
  _t.registroFotografico = Date.now();

  var detalleEjecucionReal = null;
  try {
    detalleEjecucionReal = regenerarEjecucionReal_(ss, meta.numeroContrato, totalPorItemClave);
  } catch (errEjec) {
    Logger.log("No se pudo actualizar Ejecucion Real: " + errEjec);
    detalleEjecucionReal = { error: String(errEjec) };
  }
  _t.ejecucionReal = Date.now();

  pendientesFoto.forEach(function (p) {
    var primera = fotos[p.indiceFoto];
    var etiqueta = "Imagen " + p.numeros.join(", ");
    if (primera.filaRegistro) {
      hoja.getRange(p.fila, 1).setFormula(
        '=HYPERLINK("#gid=' + gidFotografico + '&range=A' + primera.filaRegistro + '";"' + etiqueta + '")'
      );
    } else {
      hoja.getRange(p.fila, 1).setValue(etiqueta);
    }
  });
  _t.pendientesFoto = Date.now();

  ss.setActiveSheet(hoja);
  ss.moveActiveSheet(3);
  _t.activarHoja = Date.now();

  return {
    leerHojasMs: _t.leerHojas - _t.inicio,
    calcularTotalesMs: _t.calcularTotales - _t.leerHojas,
    borrarCrearHojaMs: _t.borrarCrearHoja - _t.calcularTotales,
    construirArraysMs: _t.construirArrays - _t.borrarCrearHoja,
    escriturasEnBloqueMs: _t.escriturasEnBloque - _t.construirArrays,
    mergesMs: _t.merges - _t.escriturasEnBloque,
    anchoColumnasMs: _t.anchoColumnas - _t.merges,
    registroFotograficoMs: _t.registroFotografico - _t.anchoColumnas,
    ejecucionRealMs: _t.ejecucionReal - _t.registroFotografico,
    pendientesFotoMs: _t.pendientesFoto - _t.ejecucionReal,
    activarHojaMs: _t.activarHoja - _t.pendientesFoto,
    totalMs: _t.activarHoja - _t.inicio,
    numFilas: totalFilas,
    numFotos: fotos.length,
    numMerges: merges.length,
    detalleEjecucionReal: detalleEjecucionReal,
  };
}

// Mismo criterio que tipoUnidad() en el frontend (frontend/index.html), para
// saber que formula de calculo le corresponde a cada fila segun su unidad.
function tipoUnidad_(unidadRaw) {
  var u = (unidadRaw || "").trim().toUpperCase().replace(/[.\s]/g, "");
  if (u === "M3/KM" || u === "M3KM") return "volumen_km";
  if (u === "M3") return "volumen";
  if (u === "M2") return "area";
  if (u === "M" || u === "ML") return "longitud";
  if (u === "UN" || u === "UND") return "unidad";
  if (u === "DIA" || u === "DIAS" || u === "DÍA" || u === "DÍAS") return "dias";
  return "otro";
}

// Arma la formula de la columna "Cantidad Parcial" para la fila dada, segun
// el tipo de unidad del item: deja visible cual operacion se esta haciendo
// (largo x ancho x alto para volumen, largo x ancho para area, etc.) y la
// multiplica por la columna "Cantidad" (H, cuantos elementos iguales tienen
// esta misma medida). Para UND/DIAS/otro no hay multiplicador: la columna
// "Cantidad" ya es directamente el valor final.
function formulaCantidadParcial_(tipo, fila) {
  if (tipo === "longitud") return "=C" + fila + "*H" + fila;
  if (tipo === "area") return "=C" + fila + "*D" + fila + "*H" + fila;
  if (tipo === "volumen") {
    return "=IF(N(F" + fila + ")>0;F" + fila + ";C" + fila + "*D" + fila + "*E" + fila + ")*H" + fila;
  }
  if (tipo === "volumen_km") return "=F" + fila + "*G" + fila + "*H" + fila;
  return "=H" + fila;
}

// Igual que formulaCantidadParcial_ pero calculado en JS (no como formula de
// hoja), para usar donde se necesita el numero ya resuelto: al leer una obra
// (cantidadEjecutada acumulada contra lo presupuestado) y en el historial de
// medidas que ve el usuario en la app.
function calcularBase_(tipo, m) {
  if (tipo === "longitud") return Number(m.longitud) || 0;
  if (tipo === "area") return (Number(m.longitud) || 0) * (Number(m.ancho) || 0);
  if (tipo === "volumen") {
    var directo = Number(m.volumen) || 0;
    if (directo > 0) return directo;
    return (Number(m.longitud) || 0) * (Number(m.ancho) || 0) * (Number(m.alto) || 0);
  }
  if (tipo === "volumen_km") return (Number(m.volumen) || 0) * (Number(m.distanciaKm) || 0);
  return Number(m.cantidad) || 0;
}

function calcularCantidadParcial_(tipo, m) {
  if (tipo === "unidad" || tipo === "dias" || tipo === "otro") return calcularBase_(tipo, m);
  return calcularBase_(tipo, m) * (Number(m.cantidad) || 1);
}

function leerMetaContrato_(ss) {
  var hoja = ss.getSheetByName("Config");
  var valores = hoja.getRange(1, 1, 5, 2).getValues();
  var meta = {};
  valores.forEach(function (r) {
    if (r[0] === "Nombre Obra") meta.nombreObra = r[1];
    else if (r[0] === "Número Contrato") meta.numeroContrato = r[1];
    else if (r[0] === "Objeto") meta.objeto = r[1];
    else if (r[0] === "Contratista") meta.contratista = r[1];
    else if (r[0] === "Supervisor") meta.supervisor = r[1];
  });
  return meta;
}

// Reconstruye la hoja "Registro Fotografico": un anexo separado de la
// Memoria de Calculo para que las fotos se puedan ver grandes y bien
// presentadas (listas para imprimir/adjuntar) sin desordenar las filas y
// columnas de la tabla de medidas. Encabezado con los datos del contrato,
// agrupado por direccion, cada foto con el nombre de la actividad y el
// mismo numero "Imagen N" que queda referenciado en la Memoria de Calculo.
function regenerarRegistroFotografico_(ss, fotos, meta) {
  var NOMBRE_HOJA = "Registro Fotográfico";
  var COLS = 4;

  var hojaVieja = ss.getSheetByName(NOMBRE_HOJA);
  if (hojaVieja) ss.deleteSheet(hojaVieja);
  var hoja = ss.insertSheet(NOMBRE_HOJA);

  // Mismo patron que regenerarMemoriaCalculo_/regenerarEjecucionReal_: se
  // arma todo en memoria (valores, formatos, filas a fusionar) y se aplica
  // con llamadas en bloque en vez de una por celda/fila. Las formulas
  // IMAGE()/HYPERLINK() de las fotos se aplican aparte al final (son pocas,
  // una por foto, y solo existen en la columna 1 de esas filas puntuales).
  var valores = [];
  var backgrounds = [];
  var negritas = [];
  var coloresTexto = [];
  var alineaciones = [];
  var tamanosFuente = [];
  var wraps = [];
  var merges = [];
  var filasAltura = [];
  var formulasImagen = []; // { fila, formula }

  function filaVacia() { return new Array(COLS).fill(""); }
  function agregarFila(vals, opts) {
    opts = opts || {};
    valores.push(vals);
    backgrounds.push(new Array(COLS).fill(opts.bg || "#ffffff"));
    negritas.push(new Array(COLS).fill(opts.bold ? "bold" : "normal"));
    coloresTexto.push(new Array(COLS).fill(opts.color || "#000000"));
    alineaciones.push(new Array(COLS).fill(opts.align || "left"));
    tamanosFuente.push(new Array(COLS).fill(opts.size || 11));
    wraps.push(new Array(COLS).fill(!!opts.wrap));
    return valores.length;
  }

  var valsTitulo = filaVacia(); valsTitulo[0] = "REGISTRO FOTOGRÁFICO";
  var filaTitulo = agregarFila(valsTitulo, { bg: "#1F4E78", color: "#FFFFFF", bold: true, align: "center", size: 14 });
  merges.push({ fila: filaTitulo, columnas: COLS });
  agregarFila(filaVacia()); // fila en blanco despues del titulo

  [
    ["Obra", meta.nombreObra], ["Contrato", meta.numeroContrato], ["Objeto", meta.objeto],
    ["Contratista", meta.contratista], ["Supervisor", meta.supervisor],
  ].forEach(function (d) {
    var vals = filaVacia(); vals[0] = d[0]; vals[1] = d[1] || "";
    var f = agregarFila(vals, { bold: false, wrap: true });
    // La etiqueta ("Obra", "Contrato"...) si va en negrita, el valor no; se
    // corrige encima del array recien agregado en vez de complicar agregarFila.
    negritas[f - 1][0] = "bold";
    merges.push({ fila: f, columnas: COLS, desde: 2 });
  });
  agregarFila(filaVacia()); // fila en blanco antes de las fotos

  if (!fotos.length) {
    var valsVacio = filaVacia(); valsVacio[0] = "Todavía no hay fotos registradas.";
    var filaVacioIdx = agregarFila(valsVacio, { align: "center", color: "#6B7686" });
    merges.push({ fila: filaVacioIdx, columnas: COLS });
  } else {
    var direccionActual = null;
    fotos.forEach(function (f) {
      if (f.direccion !== direccionActual) {
        direccionActual = f.direccion;
        var valsDir = filaVacia(); valsDir[0] = direccionActual;
        var filaDir = agregarFila(valsDir, { bg: "#1F4E78", color: "#FFFFFF", bold: true, align: "center" });
        merges.push({ fila: filaDir, columnas: COLS });
      }

      var etiqueta = "Imagen " + f.numero + "  —  " + f.item + " — " + f.descripcion +
        (f.observacion ? " (" + f.observacion + ")" : "");
      var valsEtiqueta = filaVacia(); valsEtiqueta[0] = etiqueta;
      var filaEtiqueta = agregarFila(valsEtiqueta, { bold: true, align: "center", wrap: true });
      merges.push({ fila: filaEtiqueta, columnas: COLS });

      f.filaRegistro = filaEtiqueta + 1;
      var filaImagen = agregarFila(filaVacia(), { align: "center" });
      merges.push({ fila: filaImagen, columnas: COLS });
      filasAltura.push({ fila: filaImagen, altura: 260 });

      var miniaturaUrl = miniaturaFoto_(f.fotoUrl);
      if (miniaturaUrl) {
        formulasImagen.push({ fila: filaImagen, formula: '=IMAGE("' + miniaturaUrl + '";1)' });
      } else if (f.fotoUrl) {
        formulasImagen.push({ fila: filaImagen, formula: '=HYPERLINK("' + f.fotoUrl + '";"Ver foto")' });
      }

      agregarFila(filaVacia()); // fila en blanco entre fotos
    });
  }

  var totalFilas = valores.length;
  hoja.getRange(1, 1, totalFilas, COLS).setValues(valores);
  hoja.getRange(1, 1, totalFilas, COLS).setBackgrounds(backgrounds);
  hoja.getRange(1, 1, totalFilas, COLS).setFontWeights(negritas);
  hoja.getRange(1, 1, totalFilas, COLS).setFontColors(coloresTexto);
  hoja.getRange(1, 1, totalFilas, COLS).setHorizontalAlignments(alineaciones);
  hoja.getRange(1, 1, totalFilas, COLS).setFontSizes(tamanosFuente);
  hoja.getRange(1, 1, totalFilas, COLS).setWraps(wraps);

  merges.forEach(function (mrg) {
    var desde = mrg.desde || 1;
    hoja.getRange(mrg.fila, desde, 1, COLS - desde + 1).merge();
  });
  filasAltura.forEach(function (fa) { hoja.setRowHeight(fa.fila, fa.altura); });
  formulasImagen.forEach(function (fi) { hoja.getRange(fi.fila, 1).setFormula(fi.formula); });

  for (var c = 1; c <= COLS; c++) hoja.setColumnWidth(c, 160);

  ss.setActiveSheet(hoja);
  ss.moveActiveSheet(4);

  return hoja.getSheetId();
}

function miniaturaFoto_(fotoUrl) {
  var m = (fotoUrl || "").match(/\/d\/([^/]+)/);
  return m ? "https://drive.google.com/thumbnail?id=" + m[1] + "&sz=w600" : "";
}

// ---------- Ejecucion Real (presupuesto oficial con precios vs ejecutado) ----------
//
// Busca en la carpeta "Presupuestos" (la misma donde el usuario sube el
// presupuesto SIN precios para crear la obra) un archivo cuyo nombre
// contenga el numero de contrato de esta obra. Ese archivo es el
// presupuesto OFICIAL con valores unitarios, un documento aparte que el
// usuario sube manualmente a esa carpeta (no lo genera la app).
function buscarPresupuestoOficialPorContrato_(numeroContrato) {
  if (!numeroContrato) return null;
  var folder = DriveApp.getFolderById(carpetaPresupuestos_());
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    if (f.getName().indexOf(numeroContrato) !== -1) return f.getId();
  }
  return null;
}

// regenerarEjecucionReal_ se ejecuta en CADA guardado de medida (a traves de
// regenerarMemoriaCalculo_), y sin cache eso significa reconvertir el .xlsx
// del presupuesto oficial de cero (Drive.Files.create + lectura +
// Drive.Files.remove) cada vez, aunque ese archivo no haya cambiado desde el
// guardado anterior -- es la parte mas cara de todo el guardado y la
// principal causa de que "tarda mucho tiempo". En vez de convertir y borrar
// el archivo temporal cada vez, se guarda una copia "sombra" (Google Sheet)
// y se reutiliza mientras el archivo original no cambie: se compara contra
// su fecha de ultima modificacion (getLastUpdated) guardada en
// PropertiesService. Solo se vuelve a convertir cuando esa fecha cambia
// (el usuario subio una version nueva del presupuesto oficial).
function convertirYLeerCacheado_(fileId) {
  var props = PropertiesService.getScriptProperties();
  var claveShadow = "shadowSheet_" + fileId;
  var claveFecha = "shadowFecha_" + fileId;
  var archivo = DriveApp.getFileById(fileId);
  var fechaActual = archivo.getLastUpdated().getTime();
  var shadowId = props.getProperty(claveShadow);
  var fechaCacheada = Number(props.getProperty(claveFecha) || 0);

  if (shadowId && fechaCacheada === fechaActual) {
    try {
      var ssShadow = SpreadsheetApp.openById(shadowId);
      return ssShadow.getSheets()[0].getDataRange().getValues();
    } catch (eShadow) {
      // La copia sombra ya no existe (se borro a mano en Drive); se
      // regenera abajo como si no hubiera cache.
    }
  }

  var fileBlob = archivo.getBlob();
  var resource = { name: "(cache interno) " + archivo.getName(), mimeType: MimeType.GOOGLE_SHEETS };
  var nuevoShadow = Drive.Files.create(resource, fileBlob);

  if (shadowId) {
    try { Drive.Files.remove(shadowId); } catch (eBorrar) { /* ya no existia, sin problema */ }
  }
  props.setProperty(claveShadow, nuevoShadow.id);
  props.setProperty(claveFecha, String(fechaActual));

  return SpreadsheetApp.openById(nuevoShadow.id).getSheets()[0].getDataRange().getValues();
}

var AIU_KEYWORDS_ = ["TOTAL COSTOS DIRECTOS", "ADMINISTRACION", "IMPREVISTOS", "UTILIDAD", "COSTO TOTAL OBRA", "COSTO DIRECTO OBRA"];

// Lee y clasifica el presupuesto oficial (con precios) fila por fila:
// nivel 0 = direccion, 1 = capitulo, 2 = subtitulo sin cantidad propia,
// 3 = item real (el unico nivel que se cruza contra las medidas de campo),
// 4 = fila de AIU (Administracion/Imprevistos/Utilidad/Costo total), que
// solo se muestra de forma informativa (sin ejecucion).
function leerPresupuestoOficialConPrecios_(fileId, direccionesConocidas) {
  var filasRaw = convertirYLeerCacheado_(fileId);
  var startIdx = -1;
  for (var i = 0; i < filasRaw.length; i++) {
    if (normalizarTexto_(filasRaw[i][1]) === "ITEM") { startIdx = i + 1; break; }
  }
  if (startIdx === -1) throw new Error("No se encontro la fila de encabezado ITEM en el presupuesto oficial");

  var filas = [];
  var direccionCtx = "";
  for (var i = startIdx; i < filasRaw.length; i++) {
    var item = filasRaw[i][1], desc = filasRaw[i][2], und = filasRaw[i][3], cant = filasRaw[i][4], vrUnit = filasRaw[i][5], vrParcial = filasRaw[i][6];
    var itemTxt = (item === "" || item === null || item === undefined) ? "" : String(item);
    var undTxt = normalizarTexto_(und);
    var descTxt = normalizarTexto_(desc);
    if (itemTxt === "" && undTxt === "" && descTxt === "") continue;
    if (itemTxt === "ITEM" || descTxt === "DESCRIPCION" || descTxt === "DESCRIPCIÓN") continue;
    var esDireccionConocida = direccionesConocidas[descTxt];
    var esAIU = AIU_KEYWORDS_.some(function (k) { return descTxt.toUpperCase().indexOf(k) === 0; });
    var nivel;
    if (itemTxt === "" && undTxt === "" && esDireccionConocida) { nivel = 0; direccionCtx = descTxt; }
    else if (itemTxt === "" && undTxt === "" && esAIU) { nivel = 4; }
    else if (itemTxt === "" && undTxt === "") { continue; }
    else if (undTxt === "" && /^\d+(\.0)?$/.test(itemTxt)) { nivel = 1; }
    else if (undTxt === "") { nivel = 2; }
    else { nivel = 3; }
    filas.push({
      nivel: nivel, direccion: direccionCtx, item: itemTxt, descripcion: descTxt, und: undTxt,
      cant: (cant === "" ? "" : Number(cant)), vrUnit: (vrUnit === "" ? "" : Number(vrUnit)), vrParcial: (vrParcial === "" ? "" : Number(vrParcial)),
    });
  }
  return filas;
}

// Etiquetas de las filas de cierre que trae el archivo oficial (usadas
// para reconstruirlas con formulas en regenerarEjecucionReal_ -- ver ahi).
// Basta con que la descripcion CONTENGA alguna de estas frases.
var ETIQUETAS_CIERRE_OFICIAL_ = [
  "COSTO DIRECTO OBRA", "TOTAL COSTOS DIRECTOS", "ADMINISTRACION", "IMPREVISTOS", "UTILIDAD",
  "COSTO TOTAL OBRA", "COSTO TOTAL PROYECTO",
];
// Revisa tanto el item como la descripcion: el archivo oficial no siempre
// pone la etiqueta de cierre en la misma columna (a veces "COSTO DIRECTO
// OBRA"/"COSTO TOTAL PROYECTO" quedan en la columna ITEM en vez de
// DESCRIPCION, lo que antes hacia que esta fila se colara sin filtrar).
function esFilaDeCierreOficial_(item, descripcion) {
  var texto = ((item || "") + " " + (descripcion || "")).toUpperCase();
  return ETIQUETAS_CIERRE_OFICIAL_.some(function (e) { return texto.indexOf(e) !== -1; });
}

function colLetraEjecucionReal_(n) {
  var s = "";
  while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// Reconstruye por completo la hoja "Ejecucion Real": mismo patron que
// regenerarMemoriaCalculo_ (se borra y se arma de nuevo cada vez que hay un
// cambio en las medidas), pero cruzando el presupuesto OFICIAL con precios
// (archivo aparte en la carpeta "Presupuestos") contra totalPorItemClave
// (cantidad ya ejecutada por direccion+item, calculada en
// regenerarMemoriaCalculo_ con la misma formula que usa la Memoria de
// Calculo). Si no se encuentra el presupuesto oficial de esta obra
// (todavia no se ha subido, o el numero de contrato no coincide con el
// nombre del archivo), no hace nada -- no rompe el guardado de medidas.
function regenerarEjecucionReal_(ss, numeroContrato, totalPorItemClave) {
  var _t = { inicio: Date.now() };
  var fileId = buscarPresupuestoOficialPorContrato_(numeroContrato);
  _t.buscarArchivo = Date.now();
  if (!fileId) {
    Logger.log("Ejecucion Real: no se encontro presupuesto oficial para el contrato " + numeroContrato);
    return { sinArchivoOficial: true, buscarArchivoMs: _t.buscarArchivo - _t.inicio };
  }

  var hojaPres = ss.getSheetByName("Presupuesto");
  var direccionesConocidas = {};
  var ultimaFilaPres = hojaPres.getLastRow();
  if (ultimaFilaPres > 1) {
    hojaPres.getRange(2, 1, ultimaFilaPres - 1, 1).getValues().forEach(function (r) {
      if (r[0]) direccionesConocidas[normalizarTexto_(r[0])] = true;
    });
  }
  _t.direccionesConocidas = Date.now();

  var filas = leerPresupuestoOficialConPrecios_(fileId, direccionesConocidas);
  _t.leerPresupuestoOficial = Date.now();

  // Los items agregados manualmente desde la app (boton "+ Agregar item
  // nuevo") no estan en el archivo oficial con precios -- ese es un
  // documento aparte que el usuario sube a mano, la app nunca lo edita.
  // Para que igual queden reflejados aqui, se buscan en la hoja
  // "Presupuesto" de la obra los items con Origen="APP" (ver nota en
  // crearObra_/agregarItemPresupuesto_ -- OJO, no basta con mirar si
  // FilaOrigen esta vacia: en obras creadas antes de que existiera esa
  // columna, TODOS los items originales la tienen vacia tambien, y por eso
  // el chequeo real y confiable es la columna Origen) que ademas no
  // aparezcan ya en "filas" (por si alguna vez coincide con un item que si
  // esta en el archivo oficial).
  var clavesOficiales = {};
  filas.forEach(function (f) { if (f.nivel === 3) clavesOficiales[f.direccion + "||" + f.item] = true; });

  var filasPresCompleta = ultimaFilaPres > 1 ? hojaPres.getRange(2, 1, ultimaFilaPres - 1, 9).getValues() : [];
  var itemsNuevos = [];
  filasPresCompleta.forEach(function (r) {
    var direccionR = normalizarTexto_(r[0]), itemR = normalizarTexto_(r[2]);
    if (!direccionR || !itemR) return;
    if (r[8] !== "APP") return; // no fue agregado desde la app: no tocar
    if (clavesOficiales[direccionR + "||" + itemR]) return; // ya esta en el archivo oficial
    itemsNuevos.push({
      direccion: direccionR, item: itemR, descripcion: normalizarTexto_(r[3]), und: normalizarTexto_(r[4]),
      vrUnit: (r[7] === "" || r[7] == null) ? "" : Number(r[7]),
    });
  });

  // Cada item nuevo se inserta justo DEBAJO del ultimo renglon de su
  // propia direccion (no en una seccion aparte al final de toda la hoja):
  // "direccion" queda marcado en cada renglon de "filas" con el contexto
  // de la ultima cabecera nivel 0 vista al leer el archivo oficial (ver
  // leerPresupuestoOficialConPrecios_), asi que buscar la ULTIMA fila con
  // ese mismo texto de direccion encuentra el final exacto de su bloque.
  // Queda como item nivel 3 normal justo antes de la siguiente direccion,
  // lo que ademas hace que SI sume dentro del subtotal de su capitulo y
  // direccion (igual que cualquier item real) -- correcto, porque si tiene
  // cantidad ejecutada es porque de verdad se ejecuto en esa direccion.
  var nuevosPorDireccion = {};
  itemsNuevos.forEach(function (n) {
    if (!nuevosPorDireccion[n.direccion]) nuevosPorDireccion[n.direccion] = [];
    nuevosPorDireccion[n.direccion].push(n);
  });
  Object.keys(nuevosPorDireccion).forEach(function (direccion) {
    var nuevasFilas = nuevosPorDireccion[direccion].map(function (n) {
      return {
        nivel: 3, direccion: n.direccion, item: n.item, descripcion: n.descripcion + " (agregado desde la app)", und: n.und,
        cant: "", vrUnit: n.vrUnit, vrParcial: "", // "cant"/"vrParcial" (CONTRATADO) vacios: no fueron parte del contrato original
      };
    });
    // Se busca especificamente el ultimo item real (nivel 3) de la
    // direccion, no solo "la ultima fila con esa direccion": las filas de
    // cierre (TOTAL COSTOS DIRECTOS, ADMINISTRACION, IMPREVISTOS,
    // UTILIDAD... nivel 4) y alguna fila de subtotal sin etiqueta que trae
    // el archivo oficial (nivel 2) quedan con el mismo contexto de
    // direccion pero van DESPUES del ultimo item real -- si se contaran,
    // el item nuevo quedaria despues de esas filas en vez de justo debajo
    // del ultimo item real, que es donde se pidio.
    var ultimaFilaDir = -1;
    for (var k = 0; k < filas.length; k++) { if (filas[k].direccion === direccion && filas[k].nivel === 3) ultimaFilaDir = k; }
    if (ultimaFilaDir === -1) {
      // La direccion no aparece para nada en el archivo oficial (caso
      // raro): se agrega al final de todo con su propio encabezado, para
      // no perderla.
      filas.push({
        nivel: 0, direccion: "", item: "", descripcion: direccion + " (agregado desde la app, sin match en el presupuesto oficial)",
        und: "", cant: "", vrUnit: "", vrParcial: "",
      });
      nuevasFilas.forEach(function (f) { filas.push(f); });
    } else {
      filas.splice.apply(filas, [ultimaFilaDir + 1, 0].concat(nuevasFilas));
    }
  });

  // Se reconstruye el cierre de cada direccion (TOTAL COSTOS DIRECTOS TCD /
  // ADMINISTRACION / IMPREVISTOS / UTILIDAD / COSTO TOTAL OBRA) con
  // FORMULAS en vez de los valores fijos que trae el archivo oficial -- asi
  // se recalculan solos si cambia algo (y las tres direcciones quedan
  // consistentes entre si, incluida la que no traia "COSTO TOTAL OBRA").
  // Tambien se deja el encabezado de cada direccion (nivel 0) SOLO con el
  // nombre, sin ningun subtotal en sus columnas.
  var filasFinal = [];
  var tcdRowsIdx = []; // indice (en filasFinal) del TCD de cada direccion, para el resumen general de mas abajo
  for (var fi = 0; fi < filas.length; fi++) {
    var fEnc = filas[fi];
    if (fEnc.nivel !== 0) { filasFinal.push(fEnc); continue; } // no deberia pasar si el archivo esta bien formado, pero por seguridad

    filasFinal.push({ nivel: 0, direccion: fEnc.direccion, item: "", descripcion: fEnc.descripcion, und: "" });

    var finDireccion = filas.length;
    for (var fj = fi + 1; fj < filas.length; fj++) { if (filas[fj].nivel === 0) { finDireccion = fj; break; } }

    var capRowsIdx = [];
    var primerItemIdx = -1, ultimoItemIdx = -1;
    for (var fk = fi + 1; fk < finDireccion; fk++) {
      var fc = filas[fk];
      // Solo se conservan capitulos/subitems/items reales (nivel 1/2/3).
      // Las filas de cierre que trae el archivo oficial (TOTAL COSTOS
      // DIRECTOS TCD, ADMINISTRACION, IMPREVISTOS, UTILIDAD, COSTO TOTAL
      // OBRA/COSTO DIRECTO OBRA/COSTO TOTAL PROYECTO -- normalmente
      // nivel 4, pero a veces el archivo las trae con texto en la columna
      // ITEM en vez de vacia, lo que las clasifica como nivel 2 por error
      // -- y cualquier fila suelta de subtotal sin etiqueta) se descartan
      // a proposito: se reemplazan por las 5 de abajo, formuladas. Por eso
      // ademas del nivel se filtra por texto: si el item o la descripcion
      // coincide con una etiqueta de cierre conocida, se descarta aunque
      // haya quedado mal clasificada. Como salvaguarda extra, tambien se
      // descarta cualquier fila sin item NI descripcion (no puede ser un
      // item/capitulo real) aunque no matchee ninguna etiqueta.
      var sinTextoAlguno = !normalizarTexto_(fc.item) && !normalizarTexto_(fc.descripcion);
      if ((fc.nivel === 1 || fc.nivel === 2 || fc.nivel === 3) && !sinTextoAlguno && !esFilaDeCierreOficial_(fc.item, fc.descripcion)) {
        filasFinal.push(fc);
        var idxFinal = filasFinal.length - 1;
        if (fc.nivel === 1) capRowsIdx.push(idxFinal);
        if (primerItemIdx === -1) primerItemIdx = idxFinal;
        ultimoItemIdx = idxFinal;
      }
    }

    if (primerItemIdx !== -1) {
      filasFinal.push({
        nivel: 10, direccion: fEnc.direccion, item: "", descripcion: "TOTAL COSTOS DIRECTOS TCD", und: "",
        _capRows: capRowsIdx.length ? capRowsIdx : null,
        _rangoItems: capRowsIdx.length ? null : [primerItemIdx, ultimoItemIdx],
      });
      tcdRowsIdx.push(filasFinal.length - 1);
      filasFinal.push({ nivel: 11, direccion: fEnc.direccion, item: "", descripcion: "ADMINISTRACION (30,4% DEL TCD)", und: "" });
      filasFinal.push({ nivel: 12, direccion: fEnc.direccion, item: "", descripcion: "IMPREVISTOS (1% DEL TCD)", und: "" });
      filasFinal.push({ nivel: 13, direccion: fEnc.direccion, item: "", descripcion: "UTILIDAD (6% DEL TCD)", und: "" });
      filasFinal.push({ nivel: 14, direccion: fEnc.direccion, item: "", descripcion: "COSTO TOTAL OBRA", und: "" });
    }

    fi = finDireccion - 1; // el for exterior le suma 1 al terminar la vuelta
  }

  // Resumen general al final de toda la hoja: mismo patron de 5 filas
  // (TOTAL COSTOS DIRECTOS TCD / ADMINISTRACION / IMPREVISTOS / UTILIDAD /
  // COSTO TOTAL OBRA), pero el TCD de aca suma el TCD de CADA direccion en
  // vez de los capitulos de una sola -- se reutiliza tal cual la misma
  // logica de formulas (nivel 10-14 se procesan igual sin importar si
  // "_capRows" apunta a capitulos o a los TCD de cada direccion).
  if (tcdRowsIdx.length) {
    filasFinal.push({ nivel: 0, direccion: "", item: "", descripcion: "RESUMEN GENERAL DE LA OBRA", und: "" });
    filasFinal.push({
      nivel: 10, direccion: "", item: "", descripcion: "TOTAL COSTOS DIRECTOS TCD", und: "",
      _capRows: tcdRowsIdx, _rangoItems: null,
    });
    filasFinal.push({ nivel: 11, direccion: "", item: "", descripcion: "ADMINISTRACION (30,4% DEL TCD)", und: "" });
    filasFinal.push({ nivel: 12, direccion: "", item: "", descripcion: "IMPREVISTOS (1% DEL TCD)", und: "" });
    filasFinal.push({ nivel: 13, direccion: "", item: "", descripcion: "UTILIDAD (6% DEL TCD)", und: "" });
    filasFinal.push({ nivel: 14, direccion: "", item: "", descripcion: "COSTO TOTAL OBRA", und: "" });
  }

  filas = filasFinal;

  var NOMBRE_HOJA = "Ejecucion Real";
  var sh = ss.getSheetByName(NOMBRE_HOJA);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(NOMBRE_HOJA);
  _t.borrarCrearHoja = Date.now();

  var COL = { ITEM: 1, DESC: 2, UND: 3, CCANT: 4, CVU: 5, CVP: 6, ACANT: 7, AVR: 8, PCANT: 9, PVR: 10, TCANT: 11, TVR: 12, SCANT: 13, SVR: 14 };
  var HEADER_BG = "#1F4E78";

  // Igual razon que el resto: ~20 llamadas sueltas (7 merge+setValue, 12
  // setValue de subtitulos) para dos filas de encabezado tardaban casi 2
  // segundos. Un solo setValues() con las dos filas armadas en memoria, y
  // los merge() (esos si van uno por uno, Sheets no los agrupa) despues.
  var filaEnc1 = new Array(14).fill("");
  filaEnc1[COL.ITEM - 1] = "ITEM";
  filaEnc1[COL.DESC - 1] = "DESCRIPCIÓN";
  filaEnc1[COL.UND - 1] = "CONTRATADO";
  filaEnc1[COL.ACANT - 1] = "ACU ACTA ANTERIOR";
  filaEnc1[COL.PCANT - 1] = "PRESENTE ACTA FINAL";
  filaEnc1[COL.TCANT - 1] = "ACUMULADO TOTAL";
  filaEnc1[COL.SCANT - 1] = "SALDO";

  var sub2 = {};
  sub2[COL.UND] = "UNID"; sub2[COL.CCANT] = "CANT"; sub2[COL.CVU] = "VR. UNITARIO"; sub2[COL.CVP] = "VR. PARCIAL";
  sub2[COL.ACANT] = "CANT"; sub2[COL.AVR] = "VR TOR";
  sub2[COL.PCANT] = "CANT"; sub2[COL.PVR] = "VR. TOTAL";
  sub2[COL.TCANT] = "CANT"; sub2[COL.TVR] = "VR. TOTAL";
  sub2[COL.SCANT] = "CANT"; sub2[COL.SVR] = "VR. TOTAL";
  var filaEnc2 = new Array(14).fill("");
  Object.keys(sub2).forEach(function (c) { filaEnc2[Number(c) - 1] = sub2[c]; });

  sh.getRange(1, 1, 2, 14).setValues([filaEnc1, filaEnc2]);
  sh.getRange(1, COL.ITEM, 2, 1).merge();
  sh.getRange(1, COL.DESC, 2, 1).merge();
  sh.getRange(1, COL.UND, 1, 4).merge();
  sh.getRange(1, COL.ACANT, 1, 2).merge();
  sh.getRange(1, COL.PCANT, 1, 2).merge();
  sh.getRange(1, COL.TCANT, 1, 2).merge();
  sh.getRange(1, COL.SCANT, 1, 2).merge();

  sh.getRange(1, 1, 2, 14).setFontWeight("bold").setFontColor("#FFFFFF").setBackground(HEADER_BG).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  sh.setFrozenRows(2);
  sh.setFrozenColumns(2);
  _t.encabezado = Date.now();

  var DATA_START = 3;
  var numFilas = filas.length;
  if (numFilas === 0) return { numFilas: 0, totalMs: Date.now() - _t.inicio };

  // Igual que en regenerarMemoriaCalculo_: antes esto eran hasta 5 vueltas
  // sobre cada fila con varias llamadas individuales a la API de Sheets por
  // vuelta (facil superar los mil llamados con un presupuesto oficial de
  // tamaño normal -- medido: casi 11 segundos). Ahora se arma todo en
  // memoria y se aplica con un puñado de llamadas en bloque.
  //
  // Las columnas "dinamicas" mezclan, segun la fila, un valor plano (la
  // cantidad ejecutada de un item real, o el VR. Parcial que trae el
  // archivo oficial para items/capitulos) o una formula (sumas de
  // acumulados/saldos, o los totales de cierre por direccion) -- por eso
  // van todas por setFormulas: ahi un string que no empieza con "=" se
  // guarda igual que si se hubiera tecleado directo en la celda (Sheets lo
  // sigue detectando como numero), asi no hace falta separar "poner valor"
  // de "poner formula" celda por celda. Incluye VR. PARCIAL (F) y la
  // columna VR de "Acu Acta Anterior" (H) ademas de las que ya iban
  // (I..N) -- las 5 filas de cierre de cada direccion (TOTAL COSTOS
  // DIRECTOS TCD, ADMINISTRACION, IMPREVISTOS, UTILIDAD, COSTO TOTAL
  // OBRA) las necesitan formuladas tambien ahi, no solo en J/L/N. La
  // columna G (7, "Acu Acta Anterior" CANT) va incluida solo para que el
  // rango F:N quede contiguo (un solo setFormulas); nunca se le escribe
  // nada, siempre queda vacia.
  var COLS_DINAMICAS = [COL.CVP, COL.ACANT, COL.AVR, COL.PCANT, COL.PVR, COL.TCANT, COL.TVR, COL.SCANT, COL.SVR];

  var valoresBase = [];
  var dinamico = [];
  var backgrounds = [];
  var negritas = [];
  var coloresTexto = [];
  var estilosFuente = [];

  for (var idx = 0; idx < numFilas; idx++) {
    var f = filas[idx];
    var filaVals = ["", "", "", "", "", "", "", "", "", "", "", "", "", ""];
    filaVals[COL.ITEM - 1] = f.item;
    filaVals[COL.DESC - 1] = f.descripcion;
    filaVals[COL.UND - 1] = f.und;
    if (f.cant !== "" && !isNaN(f.cant)) filaVals[COL.CCANT - 1] = f.cant;
    if (f.vrUnit !== "" && !isNaN(f.vrUnit)) filaVals[COL.CVU - 1] = f.vrUnit;
    // OJO: VR. PARCIAL (columna F, CVP) ya NO se pone aqui -- ahora va por
    // el bloque "dinamico" de mas abajo (junto con el resto de columnas
    // dinamicas), porque las filas de cierre (nivel 10-14) necesitan una
    // FORMULA ahi, y F tiene que ir en el mismo rango contiguo/mismo
    // setFormulas que esas.
    valoresBase.push(filaVals);

    var filaDin = {};
    COLS_DINAMICAS.forEach(function (c) { filaDin[c] = ""; });
    dinamico.push(filaDin);

    var bg = "#ffffff", peso = "normal", color = "#000000", estilo = "normal";
    if (f.nivel === 0) { bg = "#1F4E78"; color = "#FFFFFF"; peso = "bold"; }
    else if (f.nivel === 1) { bg = "#D9E1F2"; peso = "bold"; }
    else if (f.nivel === 2) { bg = "#FCE4D6"; estilo = "italic"; }
    else if (f.nivel === 4) { bg = "#F2F2F2"; estilo = "italic"; } // ya no deberia aparecer (se reemplazan por 10-14), se deja por si acaso
    else if (f.nivel === 10 || f.nivel === 11 || f.nivel === 12 || f.nivel === 13) { bg = "#F2F2F2"; estilo = "italic"; }
    else if (f.nivel === 14) { bg = "#D9E1F2"; peso = "bold"; }
    backgrounds.push(new Array(14).fill(bg));
    negritas.push(new Array(14).fill(peso));
    coloresTexto.push(new Array(14).fill(color));
    estilosFuente.push(new Array(14).fill(estilo));
  }
  _t.arraysBase = Date.now();

  var COLS_FORMULA_AIU = [COL.CVP, COL.AVR, COL.PVR, COL.TVR, COL.SVR]; // F, H, J, L, N

  for (var idx = 0; idx < numFilas; idx++) {
    var f = filas[idx];
    var row = DATA_START + idx;
    // VR. PARCIAL (F) por defecto: pasa tal cual el valor que trae el
    // archivo oficial para items/capitulos (igual que antes, solo que
    // ahora viaja por este bloque). Las filas de cierre (nivel 10-14) lo
    // pisan mas abajo con su propia formula.
    if (f.vrParcial !== "" && !isNaN(f.vrParcial)) dinamico[idx][COL.CVP] = f.vrParcial;
    if (f.nivel === 3) {
      var clave = f.direccion + "||" + f.item;
      var cantEjecutada = totalPorItemClave[clave] || 0;
      dinamico[idx][COL.PCANT] = String(cantEjecutada);
      dinamico[idx][COL.PVR] = "=" + colLetraEjecucionReal_(COL.PCANT) + row + "*" + colLetraEjecucionReal_(COL.CVU) + row;
      dinamico[idx][COL.TCANT] = "=" + colLetraEjecucionReal_(COL.ACANT) + row + "+" + colLetraEjecucionReal_(COL.PCANT) + row;
      dinamico[idx][COL.TVR] = "=" + colLetraEjecucionReal_(COL.AVR) + row + "+" + colLetraEjecucionReal_(COL.PVR) + row;
      dinamico[idx][COL.SCANT] = "=" + colLetraEjecucionReal_(COL.CCANT) + row + "-" + colLetraEjecucionReal_(COL.TCANT) + row;
      dinamico[idx][COL.SVR] = "=" + colLetraEjecucionReal_(COL.CVP) + row + "-" + colLetraEjecucionReal_(COL.TVR) + row;
    } else if (f.nivel === 1) {
      var finRel = numFilas;
      // nivel >= 10 (las 5 filas de cierre formuladas, ver mas arriba)
      // TAMBIEN cierra el rango del ultimo capitulo de la direccion --
      // sin esto, el ultimo capitulo terminaba sumando esas mismas filas
      // de cierre (que a su vez lo suman a el), una referencia circular
      // que Sheets resuelve como #REF!.
      for (var j = idx + 1; j < numFilas; j++) { if (filas[j].nivel <= 1 || filas[j].nivel >= 10) { finRel = j; break; } }
      var filaIni = DATA_START + idx + 1;
      var filaFin = DATA_START + finRel - 1;
      if (filaFin >= filaIni) {
        dinamico[idx][COL.PCANT] = "=SUM(" + colLetraEjecucionReal_(COL.PCANT) + filaIni + ":" + colLetraEjecucionReal_(COL.PCANT) + filaFin + ")";
        dinamico[idx][COL.PVR] = "=SUM(" + colLetraEjecucionReal_(COL.PVR) + filaIni + ":" + colLetraEjecucionReal_(COL.PVR) + filaFin + ")";
        dinamico[idx][COL.TCANT] = "=" + colLetraEjecucionReal_(COL.ACANT) + row + "+" + colLetraEjecucionReal_(COL.PCANT) + row;
        dinamico[idx][COL.TVR] = "=" + colLetraEjecucionReal_(COL.AVR) + row + "+" + colLetraEjecucionReal_(COL.PVR) + row;
        dinamico[idx][COL.SCANT] = "=" + colLetraEjecucionReal_(COL.CCANT) + row + "-" + colLetraEjecucionReal_(COL.TCANT) + row;
        dinamico[idx][COL.SVR] = "=" + colLetraEjecucionReal_(COL.CVP) + row + "-" + colLetraEjecucionReal_(COL.TVR) + row;
      }
    } else if (f.nivel === 10) {
      // TOTAL COSTOS DIRECTOS TCD: suma de los subtotales de cada
      // capitulo (nivel 1) de la direccion -- igual criterio que antes
      // usaba el encabezado nivel 0 para PCANT/PVR, aplicado ahora a las
      // 5 columnas de plata (F/H/J/L/N). Si la direccion no tiene
      // capitulos (todos sus items van sueltos), se suma directo el
      // rango de filas de esos items.
      COLS_FORMULA_AIU.forEach(function (c) {
        var letra = colLetraEjecucionReal_(c);
        if (f._capRows) {
          dinamico[idx][c] = "=SUM(" + f._capRows.map(function (r) { return letra + (DATA_START + r); }).join(",") + ")";
        } else if (f._rangoItems) {
          dinamico[idx][c] = "=SUM(" + letra + (DATA_START + f._rangoItems[0]) + ":" + letra + (DATA_START + f._rangoItems[1]) + ")";
        }
      });
    } else if (f.nivel === 11) {
      // ADMINISTRACION (30,4% DEL TCD): siempre la fila justo arriba
      // (TOTAL COSTOS DIRECTOS TCD, nivel 10).
      COLS_FORMULA_AIU.forEach(function (c) { dinamico[idx][c] = "=" + colLetraEjecucionReal_(c) + (row - 1) + "*0.304"; });
    } else if (f.nivel === 12) {
      // IMPREVISTOS (1% DEL TCD): 2 filas arriba (TOTAL COSTOS DIRECTOS TCD).
      COLS_FORMULA_AIU.forEach(function (c) { dinamico[idx][c] = "=" + colLetraEjecucionReal_(c) + (row - 2) + "*0.01"; });
    } else if (f.nivel === 13) {
      // UTILIDAD (6% DEL TCD): 3 filas arriba (TOTAL COSTOS DIRECTOS TCD).
      COLS_FORMULA_AIU.forEach(function (c) { dinamico[idx][c] = "=" + colLetraEjecucionReal_(c) + (row - 3) + "*0.06"; });
    } else if (f.nivel === 14) {
      // COSTO TOTAL OBRA: suma de las 4 filas anteriores (TCD +
      // ADMINISTRACION + IMPREVISTOS + UTILIDAD).
      COLS_FORMULA_AIU.forEach(function (c) {
        var letra = colLetraEjecucionReal_(c);
        dinamico[idx][c] = "=" + letra + (row - 4) + "+" + letra + (row - 3) + "+" + letra + (row - 2) + "+" + letra + (row - 1);
      });
    }
  }

  _t.arraysDinamicos = Date.now();

  sh.getRange(DATA_START, 1, numFilas, 14).setValues(valoresBase);
  sh.getRange(DATA_START, 1, numFilas, 14).setBackgrounds(backgrounds);
  sh.getRange(DATA_START, 1, numFilas, 14).setFontWeights(negritas);
  sh.getRange(DATA_START, 1, numFilas, 14).setFontColors(coloresTexto);
  sh.getRange(DATA_START, 1, numFilas, 14).setFontStyles(estilosFuente);
  sh.getRange(DATA_START, COLS_DINAMICAS[0], numFilas, COLS_DINAMICAS.length).setFormulas(
    dinamico.map(function (fila) { return COLS_DINAMICAS.map(function (c) { return fila[c]; }); })
  );
  _t.escriturasEnBloque = Date.now();

  [COL.CVU, COL.CVP, COL.AVR, COL.PVR, COL.TVR, COL.SVR].forEach(function (c) { sh.getRange(DATA_START, c, numFilas, 1).setNumberFormat("$#,##0"); });
  [COL.CCANT, COL.ACANT, COL.PCANT, COL.TCANT, COL.SCANT].forEach(function (c) { sh.getRange(DATA_START, c, numFilas, 1).setNumberFormat("#,##0.00"); });
  _t.formatosNumero = Date.now();

  sh.setColumnWidth(COL.ITEM, 55);
  sh.setColumnWidth(COL.DESC, 320);
  sh.setColumnWidths(COL.UND, 13, 90);

  ss.setActiveSheet(sh);
  ss.moveActiveSheet(5);
  _t.fin = Date.now();

  return {
    buscarArchivoMs: _t.buscarArchivo - _t.inicio,
    direccionesConocidasMs: _t.direccionesConocidas - _t.buscarArchivo,
    leerPresupuestoOficialMs: _t.leerPresupuestoOficial - _t.direccionesConocidas,
    borrarCrearHojaMs: _t.borrarCrearHoja - _t.leerPresupuestoOficial,
    encabezadoMs: _t.encabezado - _t.borrarCrearHoja,
    arraysBaseMs: _t.arraysBase - _t.encabezado,
    arraysDinamicosMs: _t.arraysDinamicos - _t.arraysBase,
    escriturasEnBloqueMs: _t.escriturasEnBloque - _t.arraysDinamicos,
    formatosNumeroMs: _t.formatosNumero - _t.escriturasEnBloque,
    finMs: _t.fin - _t.formatosNumero,
    totalMs: _t.fin - _t.inicio,
    numFilas: numFilas,
  };
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
