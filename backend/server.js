require("dotenv").config();
const express = require("express");
const cors = require("cors");
const path = require("path");
const multer = require("multer");
const Anthropic = require("@anthropic-ai/sdk");
const { Packer } = require("docx");
const { generarInformeContratistaDocx } = require("./informe");

const PORT = process.env.PORT || 3000;
const APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL;
// Mismo patron que la app de gastos (gastos-jacbel): Haiku primero (rapido y
// barato), y si la lectura queda incompleta se reintenta una vez con Sonnet
// antes de rendirse y devolver lo mejor que se haya podido leer.
const CLAUDE_MODEL = process.env.CLAUDE_MODEL || "claude-haiku-4-5-20251001";
const CLAUDE_MODEL_REINTENTO = process.env.CLAUDE_MODEL_REINTENTO || "claude-sonnet-5";
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const app = express();
app.use(cors());
// 25mb: algunos endpoints (guardar datos del informe) siguen viajando como
// JSON normal; los adjuntos en si (PDFs/fotos) van aparte por multipart
// (ver "upload" mas abajo), no dentro de este limite.
app.use(express.json({ limit: "25mb" }));
app.use(express.static(path.join(__dirname, "..", "frontend")));

const upload = multer({ limits: { fileSize: 12 * 1024 * 1024 } });

function requireAppsScriptUrl() {
  if (!APPS_SCRIPT_URL) throw new Error("Falta APPS_SCRIPT_URL");
}

async function llamarAppsScript(url, options) {
  const r = await fetch(url, options);
  return r.json();
}

async function llamarAppsScriptPost(body) {
  return llamarAppsScript(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
}

app.get("/api/health", (req, res) => res.json({ ok: true }));

// Boton "🔄 Actualizar reportes" de la app: reconstruye "Memoria de
// Calculo" / "Registro Fotografico" / "Ejecucion Real" de una obra puntual
// bajo demanda (ver nota en leerObra_/actualizarReportesObra_ en Code.gs).
app.post("/api/obras/:obraId/actualizar-reportes", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const data = await llamarAppsScriptPost({ accion: "actualizar_reportes", obraId: req.params.obraId });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.get("/api/obras", async (req, res) => {
  try {
    requireAppsScriptUrl();
    res.json(await llamarAppsScript(`${APPS_SCRIPT_URL}?obras=1`));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.get("/api/presupuestos", async (req, res) => {
  try {
    requireAppsScriptUrl();
    res.json(await llamarAppsScript(`${APPS_SCRIPT_URL}?presupuestos=1`));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.get("/api/presupuesto-preview", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const { fileId } = req.query;
    if (!fileId) return res.status(400).json({ error: "Falta fileId" });
    const data = await llamarAppsScript(`${APPS_SCRIPT_URL}?previsualizar=${encodeURIComponent(fileId)}`);
    if (data && data.error) return res.status(400).json({ error: data.error });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.post("/api/obras", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const { fileId, nombreObra } = req.body;
    if (!fileId || !nombreObra) return res.status(400).json({ error: "Falta fileId o nombreObra" });
    const data = await llamarAppsScriptPost({ accion: "crear_obra", fileId, nombreObra });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.get("/api/obras/:obraId", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const data = await llamarAppsScript(`${APPS_SCRIPT_URL}?obra=${encodeURIComponent(req.params.obraId)}`);
    if (data && data.error) return res.status(400).json({ error: data.error });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.delete("/api/obras/:obraId", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const data = await llamarAppsScriptPost({ accion: "borrar_obra", obraId: req.params.obraId });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.post("/api/medidas", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const data = await llamarAppsScriptPost({ accion: "guardar_medida", ...req.body });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.post("/api/obras/:obraId/copiar-medidas", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const { direccion, item, medidas } = req.body;
    if (!direccion || !item || !Array.isArray(medidas) || !medidas.length) {
      return res.status(400).json({ error: "Falta direccion, item o mediciones para copiar" });
    }
    const data = await llamarAppsScriptPost({ accion: "copiar_medidas", obraId: req.params.obraId, ...req.body });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.put("/api/obras/:obraId/items", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const { direccion, itemViejo } = req.body;
    if (!direccion || !itemViejo) return res.status(400).json({ error: "Falta direccion o itemViejo" });
    const data = await llamarAppsScriptPost({ accion: "editar_item", obraId: req.params.obraId, ...req.body });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.post("/api/obras/:obraId/items", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const { direccion, item } = req.body;
    if (!direccion || !item) return res.status(400).json({ error: "Falta direccion o item" });
    const data = await llamarAppsScriptPost({ accion: "agregar_item", obraId: req.params.obraId, ...req.body });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.delete("/api/obras/:obraId/items", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const { direccion, item } = req.body;
    if (!direccion || !item) return res.status(400).json({ error: "Falta direccion o item" });
    const data = await llamarAppsScriptPost({ accion: "borrar_item", obraId: req.params.obraId, ...req.body });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.put("/api/medidas/:medidaId", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const data = await llamarAppsScriptPost({ accion: "editar_medida", medidaId: req.params.medidaId, ...req.body });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.delete("/api/medidas/:medidaId", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const { obraId } = req.body;
    if (!obraId) return res.status(400).json({ error: "Falta obraId" });
    const data = await llamarAppsScriptPost({ accion: "borrar_medida", medidaId: req.params.medidaId, obraId });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

// ---------- Informe del Contratista (AP2-FO-024): lectura con IA de adjuntos ----------
//
// Mismo mecanismo que ya usa la app de gastos (gastos-jacbel): se manda la
// foto o el PDF a Claude con vision, pidiendole que devuelva SOLO un JSON
// con los campos que le corresponden a ese tipo de documento. La IA nunca
// escribe directo en la hoja de la obra -- lo que devuelve solo llega a
// precargar el formulario en el navegador; el usuario revisa/corrige y
// recien ahi, al presionar "Guardar datos del informe", queda guardado.
const PROMPT_POR_TIPO_DOCUMENTO = {
  acta_inicio: `Estas leyendo el Acta de Inicio de un contrato de obra publica en Colombia (puede ser una foto o un PDF, escaneado o digital). Extrae estos datos y responde UNICAMENTE con un JSON valido, sin texto adicional, con esta forma exacta:
{
  "numeroContrato": "numero o codigo del contrato, tal como aparece",
  "objeto": "objeto/descripcion del contrato",
  "contratista": "nombre de la persona o empresa contratista",
  "valorInicial": 12345678,
  "plazo": "plazo del contrato tal como aparece (ej: '6 meses', '180 dias')",
  "fechaActaInicio": "YYYY-MM-DD",
  "fechaTerminacionInicial": "YYYY-MM-DD"
}
Reglas: "valorInicial" es el valor del contrato en pesos colombianos (COP), numero entero sin puntos ni comas ni simbolo de moneda -- si no aparece en el acta, deja 0. "fechaTerminacionInicial" normalmente se calcula sumando el plazo a la fecha del acta si no aparece explicita -- si no puedes determinarla con confianza, deja "". Responde SIEMPRE con el JSON completo aunque el documento sea dificil de leer; nunca respondas con una disculpa. Si un campo especifico es ilegible o no aparece, usa "" (o 0 para valorInicial) solo en ese campo, sin inventar.`,
  polizas: `Estas leyendo un documento de aprobacion de polizas/garantias de un contrato de obra publica en Colombia (puede ser una foto o un PDF). Extrae estos datos y responde UNICAMENTE con un JSON valido, sin texto adicional, con esta forma exacta:
{
  "polizaCumplimiento": "numero de la poliza de cumplimiento",
  "polizaResponsabilidadCivil": "numero de la poliza de responsabilidad civil extracontractual",
  "companiaAseguradora": "nombre de la compañia aseguradora",
  "amparos": [
    { "tipo": "nombre del amparo, ej Cumplimiento, Calidad, Estabilidad, Pago de Salarios y Prestaciones, Responsabilidad Civil Extracontractual", "porcentaje": 10, "valorAsegurado": 12345678, "vigenciaDesde": "YYYY-MM-DD", "vigenciaHasta": "YYYY-MM-DD" }
  ]
}
Reglas: incluye en "amparos" TODOS los amparos/coberturas que encuentres en el documento, uno por elemento del arreglo (usualmente son varios: cumplimiento, calidad, estabilidad, salarios y prestaciones, responsabilidad civil, etc). "porcentaje" es el % del valor del contrato que cubre ese amparo, como numero (sin simbolo %). "valorAsegurado" es el valor asegurado de ese amparo en pesos colombianos (COP), numero entero. Si el documento tiene varias polizas (cumplimiento y responsabilidad civil por separado, cada una con sus propios amparos), junta todos los amparos de ambas en el mismo arreglo. Responde SIEMPRE con el JSON completo aunque el documento sea dificil de leer; nunca respondas con una disculpa. Si un campo especifico es ilegible, usa "" (o 0 para los numericos) solo en ese campo, sin inventar. Si no encuentras ningun amparo, usa un arreglo vacio [].`,
  contrato: `Estas leyendo el contrato (o su primera pagina/resumen) de una obra publica en Colombia (puede ser una foto o un PDF). Extrae estos datos y responde UNICAMENTE con un JSON valido, sin texto adicional, con esta forma exacta:
{
  "numeroContrato": "numero o codigo del contrato",
  "objeto": "objeto/descripcion del contrato",
  "contratista": "nombre de la persona o empresa contratista",
  "valorInicial": 12345678,
  "plazo": "plazo del contrato tal como aparece"
}
Reglas: "valorInicial" en pesos colombianos (COP), numero entero sin puntos ni comas ni simbolo de moneda -- si no aparece, deja 0. Responde SIEMPRE con el JSON completo aunque el documento sea dificil de leer; nunca respondas con una disculpa. Si un campo especifico es ilegible, usa "" (o 0 para valorInicial) solo en ese campo, sin inventar.`,
  comprobante_anticipo: `Estas leyendo un comprobante de pago/cobro de anticipo de un contrato de obra publica en Colombia (puede ser una foto o un PDF). Extrae estos datos y responde UNICAMENTE con un JSON valido, sin texto adicional, con esta forma exacta:
{
  "valorAnticipo": 12345678,
  "porcentajeAnticipo": 30
}
Reglas: "valorAnticipo" es el valor pagado/cobrado como anticipo en pesos colombianos (COP), numero entero sin puntos ni comas ni simbolo de moneda. "porcentajeAnticipo" es el % de anticipo pactado si aparece explicito en el comprobante (numero sin simbolo %); si no aparece, deja "". Responde SIEMPRE con el JSON completo aunque el documento sea dificil de leer; nunca respondas con una disculpa. Si un campo es ilegible, usa "" (o 0 para valorAnticipo) solo en ese campo, sin inventar.`,
};

// Un campo por tipo que, si viene lleno (o hay al menos un amparo), se
// considera que la lectura "sirvio" -- si ni siquiera eso vino, se
// reintenta con el modelo mas grande antes de rendirse.
function extraccionEsUtil_(tipo, extraido) {
  if (!extraido) return false;
  if (tipo === "polizas") return Array.isArray(extraido.amparos) && extraido.amparos.length > 0 || !!extraido.polizaCumplimiento;
  if (tipo === "comprobante_anticipo") return Number(extraido.valorAnticipo) > 0;
  return !!(extraido.numeroContrato || extraido.contratista || extraido.objeto);
}

async function leerDocumentoInformeConIA_(tipo, base64Data, mediaType) {
  const prompt = PROMPT_POR_TIPO_DOCUMENTO[tipo];
  if (!prompt) return { extraido: null, necesitaRevision: false };
  const esPdf = (mediaType || "").indexOf("pdf") !== -1;
  const bloqueArchivo = esPdf
    ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64Data } }
    : { type: "image", source: { type: "base64", media_type: mediaType || "image/jpeg", data: base64Data } };

  const intentos = [CLAUDE_MODEL, CLAUDE_MODEL_REINTENTO];
  let ultimoExtraido = null;
  for (let i = 0; i < intentos.length; i++) {
    try {
      const message = await anthropic.messages.create({
        model: intentos[i],
        max_tokens: 1200,
        messages: [{ role: "user", content: [bloqueArchivo, { type: "text", text: prompt }] }],
      });
      const textBlock = message.content.find((b) => b.type === "text");
      const jsonMatch = textBlock && textBlock.text.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        const extraido = JSON.parse(jsonMatch[0]);
        ultimoExtraido = extraido;
        if (extraccionEsUtil_(tipo, extraido)) return { extraido, necesitaRevision: false };
      }
    } catch (err) {
      console.error(`Lectura IA (${tipo}) fallo con modelo ${intentos[i]}:`, err.message);
    }
  }
  return { extraido: ultimoExtraido, necesitaRevision: true };
}

// ---------- Informe de Supervision (AP2-FO-024) ----------

app.get("/api/obras/:obraId/informe-supervision", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const data = await llamarAppsScript(`${APPS_SCRIPT_URL}?informeSupervision=${encodeURIComponent(req.params.obraId)}`);
    if (data && data.error) return res.status(400).json({ error: data.error });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.put("/api/obras/:obraId/informe-supervision", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const data = await llamarAppsScriptPost({ accion: "guardar_datos_informe", obraId: req.params.obraId, ...req.body });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.post("/api/obras/:obraId/informe-supervision/documentos", upload.single("archivo"), async (req, res) => {
  try {
    requireAppsScriptUrl();
    const tipo = req.body.tipo;
    if (!tipo) return res.status(400).json({ error: "Falta el tipo de documento" });
    if (!req.file) return res.status(400).json({ error: "Falta el archivo" });

    const base64 = req.file.buffer.toString("base64");
    const mime = req.file.mimetype || "application/pdf";

    // La lectura con IA y la subida a Drive corren en paralelo: si la IA
    // falla o tarda, el archivo igual queda guardado (no dependen una de
    // la otra).
    const [lectura, subida] = await Promise.all([
      leerDocumentoInformeConIA_(tipo, base64, mime),
      llamarAppsScriptPost({
        accion: "subir_documento_informe", obraId: req.params.obraId, tipo, base64, mime, nombre: req.file.originalname,
      }),
    ]);

    if (!subida.ok) return res.status(502).json({ error: subida.error || "No se pudo guardar el archivo" });
    res.json({ ok: true, url: subida.url, campo: subida.campo, extraido: lectura.extraido, necesitaRevision: lectura.necesitaRevision });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.get("/api/obras/:obraId/informe-supervision/historial", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const data = await llamarAppsScript(`${APPS_SCRIPT_URL}?historialInformes=${encodeURIComponent(req.params.obraId)}`);
    if (data && data.error) return res.status(400).json({ error: data.error });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

// Genera el .docx del Informe del Contratista para un periodo puntual:
// junta los datos manuales (Fase 1) + el balance/resumen calculado en Apps
// Script para ese rango de fechas, arma el documento con la libreria docx
// (ver informe.js), lo sube a Drive y lo deja registrado en el historial de
// la obra, y ademas lo devuelve directo para descarga inmediata.
app.post("/api/obras/:obraId/informe-supervision/generar", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const obraId = req.params.obraId;
    const { fechaDesde, fechaHasta, tipoInforme, numeroParcial } = req.body;
    if (!fechaDesde || !fechaHasta) return res.status(400).json({ error: "Falta la fecha de inicio o de fin del periodo" });
    if (tipoInforme !== "Parcial" && tipoInforme !== "Final") return res.status(400).json({ error: "El tipo de informe debe ser Parcial o Final" });

    const [datosInforme, resumen] = await Promise.all([
      llamarAppsScript(`${APPS_SCRIPT_URL}?informeSupervision=${encodeURIComponent(obraId)}`),
      llamarAppsScriptPost({ accion: "calcular_resumen_informe", obraId, fechaDesde, fechaHasta }),
    ]);
    if (!datosInforme.ok) return res.status(400).json({ error: datosInforme.error || "No se pudieron leer los datos del informe" });
    if (!resumen.ok) return res.status(400).json({ error: resumen.error || "No se pudo calcular el resumen del periodo" });

    const datos = Object.assign({}, datosInforme.datos, {
      amparos: datosInforme.amparos || [],
      tipoInforme, numeroParcial: numeroParcial || "",
      fechaDesde, fechaHasta,
    });

    const valorContrato = Number(datos.valorInicial) || 0;
    const porcentajeAnticipo = Number(datos.porcentajeAnticipo) || 0;
    const valorActaPeriodo = resumen.totalEjecutadoPeriodoVr || 0;
    const amortizacionAnticipo = valorActaPeriodo * (porcentajeAnticipo / 100);
    const acumuladoEjecutado = resumen.totalEjecutadoAcumuladoVr || 0;
    const balance = {
      valorContrato,
      valorAnticipoPagado: Number(datos.valorAnticipo) || 0,
      porcentajeAnticipo,
      valorActaPeriodo,
      amortizacionAnticipo,
      valorNetoAPagar: valorActaPeriodo - amortizacionAnticipo,
      acumuladoEjecutado,
      saldoPorEjecutar: Math.max(0, valorContrato - acumuladoEjecutado),
      saldoAFavor: Math.max(0, acumuladoEjecutado - valorContrato),
      tipoInforme, numeroParcial: numeroParcial || "",
    };

    const doc = await generarInformeContratistaDocx(datos, balance, resumen.items || [], resumen.fotos || []);
    const buffer = await Packer.toBuffer(doc);
    const nombreArchivo = `Informe ${tipoInforme === "Final" ? "Final" : "Parcial " + (numeroParcial || "")} - ${datos.numeroContrato || obraId} - ${fechaHasta}.docx`.replace(/\s+/g, " ");

    // Se guarda en Drive y queda en el historial ademas de devolverse para
    // descarga directa, para no perder el enlace si se cierra la pestaña
    // sin descargar.
    const subida = await llamarAppsScriptPost({ accion: "subir_informe_generado", obraId, base64: buffer.toString("base64"), nombre: nombreArchivo });
    if (subida.ok) {
      await llamarAppsScriptPost({
        accion: "registrar_informe_generado", obraId, tipo: tipoInforme, numeroParcial: numeroParcial || "",
        fechaDesde, fechaHasta, urlDocx: subida.url,
      });
    }

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${nombreArchivo.replace(/"/g, "")}"`);
    if (subida.ok) res.setHeader("X-Informe-Drive-Url", subida.url);
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.listen(PORT, () => console.log(`Memorias JACBEL escuchando en puerto ${PORT}`));
