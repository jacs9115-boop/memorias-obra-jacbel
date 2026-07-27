// IMPORTANTE: en el editor de Apps Script, antes de implementar, ve a
// Servicios (icono +) y agrega "Drive API" (Advanced Google Services).
// Sin eso, la conversion de archivos .xls/.xlsx no va a funcionar.

function doGet(e) {
  try {
    if (e.parameter.obras === "1") return jsonOutput_(listarObras_());
    if (e.parameter.presupuestos === "1") return jsonOutput_(listarPresupuestos_());
    if (e.parameter.previsualizar) return jsonOutput_(parsearPresupuesto_(e.parameter.previsualizar));
    if (e.parameter.obra) return jsonOutput_(leerObra_(e.parameter.obra));
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
    if (body.accion === "editar_medida") return jsonOutput_(editarMedida_(body));
    if (body.accion === "borrar_medida") return jsonOutput_(borrarMedida_(body));
    if (body.accion === "borrar_obra") return jsonOutput_(borrarObra_(body));
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

function subirFotos_(fotos) {
  if (!fotos || !fotos.length) return [];
  var carpetaId = carpetaFotos_();
  var urls = [];
  fotos.forEach(function (f) {
    if (!f || !f.base64) return;
    var blob = Utilities.newBlob(
      Utilities.base64Decode(f.base64),
      f.tipo || "image/jpeg",
      (f.nombre || "foto") + ".jpg"
    );
    var archivo = DriveApp.getFolderById(carpetaId).createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    urls.push(archivo.getUrl());
  });
  return urls;
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
    hoja.appendRow(["ObraId", "Nombre", "SpreadsheetId", "NumeroContrato", "Objeto", "Contratista", "Supervisor", "FechaCreacion"]);
  }
  return hoja;
}

function listarObras_() {
  var hoja = obtenerHojaIndice_();
  var lastRow = hoja.getLastRow();
  if (lastRow < 2) return [];
  var valores = hoja.getRange(2, 1, lastRow - 1, 8).getValues();
  return valores.filter(function (r) { return r[0]; }).map(function (r) {
    return {
      obraId: r[0], nombre: r[1], spreadsheetId: r[2], numeroContrato: r[3],
      objeto: r[4], contratista: r[5], supervisor: r[6], fechaCreacion: r[7],
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
  var filasPres = [["Dirección", "Capítulo", "Item", "Descripción", "Unidad", "Cantidad Presupuestada"]];
  datos.direcciones.forEach(function (d) {
    d.items.forEach(function (it) {
      filasPres.push([d.nombre, it.capitulo, it.item, it.descripcion, it.unidad, it.cantidadPresupuestada]);
    });
  });
  hojaPres.getRange(1, 1, filasPres.length, 6).setValues(filasPres);
  hojaPres.getRange(1, 1, 1, 6).setFontWeight("bold");
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
    datos.meta.contratista, datos.meta.supervisor, new Date().toISOString()]);

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

function leerObra_(obraId) {
  var obra = buscarObra_(obraId);
  if (!obra) throw new Error("Obra no encontrada");

  var ss = SpreadsheetApp.openById(obra.spreadsheetId);

  var hojaPres = ss.getSheetByName("Presupuesto");
  var ultimaFilaPres = hojaPres.getLastRow();
  var filasPres = ultimaFilaPres > 1 ? hojaPres.getRange(2, 1, ultimaFilaPres - 1, 6).getValues() : [];

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

  // Reconstruye "Memoria de Cálculo" y "Registro Fotográfico" cada vez que se
  // abre la obra, sin importar si "Memoria" cambio por la app o porque el
  // usuario edito/borro filas a mano directamente en la hoja: asi las dos
  // hojas calculadas nunca quedan desincronizadas de la fuente real de datos.
  regenerarMemoriaCalculo_(ss);

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
    });
  });
  var direcciones = ordenDirecciones.map(function (nombre) {
    return { nombre: nombre, items: direccionesMap[nombre] };
  });

  return { obra: obra, direcciones: direcciones, medidas: medidas };
}
// ---------- Medidas (memoria de calculo) ----------

function guardarMedida_(body) {
  var obra = buscarObra_(body.obraId);
  if (!obra) return { ok: false, error: "Obra no encontrada" };

  var ss = SpreadsheetApp.openById(obra.spreadsheetId);
  var hoja = ss.getSheetByName("Memoria");

  var fotosUrls = subirFotos_(body.fotos);

  var id = Utilities.getUuid();
  var ahora = new Date();
  var fechaHora = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  var fila = [
    id, fechaHora, body.direccion || "", body.item || "", body.descripcion || "", body.unidad || "",
    body.longitud || "", body.ancho || "", body.alto || "", body.volumen || "", body.distanciaKm || "",
    Number(body.cantidad) || 0, fotosUrls.join("|"), body.observacion || "",
  ];
  hoja.appendRow(fila);
  regenerarMemoriaCalculo_(ss);

  return { ok: true, id: id, fechaHora: fechaHora, fotosUrls: fotosUrls };
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
      var fotosNuevas = subirFotos_(body.fotos);
      if (fotosNuevas.length) {
        var fotosActuales = separarFotos_(hoja.getRange(fila, 13).getValue());
        hoja.getRange(fila, 13).setValue(fotosActuales.concat(fotosNuevas).join("|"));
      }
      hoja.getRange(fila, 14).setValue(body.observacion || "");
      regenerarMemoriaCalculo_(ss);
      return { ok: true };
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
      regenerarMemoriaCalculo_(ss);
      return { ok: true };
    }
  }
  return { ok: false, error: "Medida no encontrada" };
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
  var NOMBRE_HOJA = "Memoria de Cálculo";

  var hojaPres = ss.getSheetByName("Presupuesto");
  var ultimaFilaPres = hojaPres.getLastRow();
  var filasPres = ultimaFilaPres > 1 ? hojaPres.getRange(2, 1, ultimaFilaPres - 1, 6).getValues() : [];

  var hojaMemoria = ss.getSheetByName("Memoria");
  var ultimaFilaMemoria = hojaMemoria.getLastRow();
  var filasMemoria = ultimaFilaMemoria > 1 ? hojaMemoria.getRange(2, 1, ultimaFilaMemoria - 1, 14).getValues() : [];

  var medidasPorItem = {};
  filasMemoria.forEach(function (r) {
    var clave = r[2] + "||" + r[3];
    if (!medidasPorItem[clave]) medidasPorItem[clave] = [];
    medidasPorItem[clave].push({
      longitud: r[6], ancho: r[7], alto: r[8], volumen: r[9], distanciaKm: r[10],
      cantidad: r[11], fotoUrl: r[12], observacion: r[13],
    });
  });

  var hojaVieja = ss.getSheetByName(NOMBRE_HOJA);
  if (hojaVieja) ss.deleteSheet(hojaVieja);
  var hoja = ss.insertSheet(NOMBRE_HOJA);

  var fila = 1;
  var direccionActual = null;
  var capituloActual = null;
  var fotos = [];
  var pendientesFoto = [];

  filasPres.forEach(function (r) {
    var direccion = r[0], capitulo = r[1], item = r[2], descripcion = r[3], unidad = r[4];
    var clave = direccion + "||" + item;
    var medidas = medidasPorItem[clave];
    if (!medidas || !medidas.length) return;

    if (direccion !== direccionActual) {
      direccionActual = direccion;
      capituloActual = null;
      hoja.getRange(fila, 1).setValue(direccion);
      hoja.getRange(fila, 1, 1, 9).merge().setFontWeight("bold").setBackground("#1F4E78").setFontColor("#FFFFFF");
      fila++;
    }
    if (capitulo !== capituloActual) {
      capituloActual = capitulo;
      hoja.getRange(fila, 1).setValue(capitulo);
      hoja.getRange(fila, 1, 1, 9).merge().setFontWeight("bold").setBackground("#DCE6F1");
      fila++;
    }

    hoja.getRange(fila, 1).setValue(item + " — " + descripcion + " (" + (unidad || "sin unidad") + ")");
    hoja.getRange(fila, 1, 1, 9).merge().setFontWeight("bold");
    fila++;

    hoja.getRange(fila, 1, 1, 9).setValues([["Foto", "Descripción / Observación", "Longitud", "Ancho", "Alto", "Volumen", "Km", "Cantidad", "Cantidad Parcial"]]);
    hoja.getRange(fila, 1, 1, 9).setFontWeight("bold").setBackground("#F0F0F0");
    fila++;

    var tipo = tipoUnidad_(unidad);
    var filaInicioMedidas = fila;
    medidas.forEach(function (m) {
      hoja.getRange(fila, 1, 1, 9).setValues([[
        "", m.observacion || "", m.longitud || "", m.ancho || "", m.alto || "", m.volumen || "", m.distanciaKm || "", m.cantidad || 0, "",
      ]]);
      hoja.getRange(fila, 9).setFormula(formulaCantidadParcial_(tipo, fila));
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
        pendientesFoto.push({ fila: fila, indiceFoto: indiceInicio, numeros: numeros });
      }
      fila++;
    });
    var filaFinMedidas = fila - 1;

    // El TOTAL solo suma la columna "Cantidad Parcial" (el resultado final).
    // La columna "Cantidad" es el multiplicador manual por fila (cuantos
    // elementos iguales), sumarla no tiene sentido.
    hoja.getRange(fila, 1).setValue("TOTAL");
    hoja.getRange(fila, 1, 1, 8).merge().setFontWeight("bold").setHorizontalAlignment("right");
    hoja.getRange(fila, 9).setFormula("=SUM(I" + filaInicioMedidas + ":I" + filaFinMedidas + ")").setFontWeight("bold");
    fila += 2;
  });

  hoja.setColumnWidth(1, 100);
  hoja.setColumnWidth(2, 220);
  for (var c = 3; c <= 7; c++) hoja.setColumnWidth(c, 80);
  hoja.setColumnWidth(8, 90);
  hoja.setColumnWidth(9, 110);
  hoja.setFrozenRows(0);

  var meta = leerMetaContrato_(ss);
  var gidFotografico = regenerarRegistroFotografico_(ss, fotos, meta);

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

  ss.setActiveSheet(hoja);
  ss.moveActiveSheet(3);
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

  var fila = 1;
  hoja.getRange(fila, 1, 1, COLS).merge().setValue("REGISTRO FOTOGRÁFICO")
    .setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center")
    .setBackground("#1F4E78").setFontColor("#FFFFFF");
  fila += 2;

  [
    ["Obra", meta.nombreObra], ["Contrato", meta.numeroContrato], ["Objeto", meta.objeto],
    ["Contratista", meta.contratista], ["Supervisor", meta.supervisor],
  ].forEach(function (d) {
    hoja.getRange(fila, 1).setValue(d[0]).setFontWeight("bold");
    hoja.getRange(fila, 2, 1, COLS - 1).merge().setValue(d[1] || "").setWrap(true);
    fila++;
  });
  fila += 1;

  if (!fotos.length) {
    hoja.getRange(fila, 1, 1, COLS).merge().setValue("Todavía no hay fotos registradas.")
      .setHorizontalAlignment("center").setFontColor("#6B7686");
  } else {
    var direccionActual = null;
    fotos.forEach(function (f) {
      if (f.direccion !== direccionActual) {
        direccionActual = f.direccion;
        hoja.getRange(fila, 1, 1, COLS).merge().setValue(direccionActual)
          .setFontWeight("bold").setBackground("#1F4E78").setFontColor("#FFFFFF")
          .setHorizontalAlignment("center");
        fila++;
      }

      var etiqueta = "Imagen " + f.numero + "  —  " + f.item + " — " + f.descripcion +
        (f.observacion ? " (" + f.observacion + ")" : "");
      hoja.getRange(fila, 1, 1, COLS).merge().setValue(etiqueta)
        .setFontWeight("bold").setHorizontalAlignment("center").setWrap(true);
      fila++;

      f.filaRegistro = fila;
      var miniaturaUrl = miniaturaFoto_(f.fotoUrl);
      hoja.getRange(fila, 1, 1, COLS).merge().setHorizontalAlignment("center").setVerticalAlignment("middle");
      if (miniaturaUrl) {
        hoja.getRange(fila, 1).setFormula('=IMAGE("' + miniaturaUrl + '";1)');
      } else if (f.fotoUrl) {
        hoja.getRange(fila, 1).setFormula('=HYPERLINK("' + f.fotoUrl + '";"Ver foto")');
      }
      hoja.setRowHeight(fila, 260);
      fila += 2;
    });
  }

  for (var c = 1; c <= COLS; c++) hoja.setColumnWidth(c, 160);

  ss.setActiveSheet(hoja);
  ss.moveActiveSheet(4);

  return hoja.getSheetId();
}

function miniaturaFoto_(fotoUrl) {
  var m = (fotoUrl || "").match(/\/d\/([^/]+)/);
  return m ? "https://drive.google.com/thumbnail?id=" + m[1] + "&sz=w600" : "";
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
