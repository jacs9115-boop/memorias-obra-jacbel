// Genera el Informe del Contratista (.docx) desde cero con la libreria
// "docx" (no es una plantilla externa con placeholders): la estructura esta
// inspirada en el formato AP2-FO-024 de ACUAVALLE (tabla de info del
// contrato, tabla de actividades, cronologia, balance financiero, tabla de
// garantias/amparos, observaciones, registro fotografico), pero adaptada
// para que la genere el CONTRATISTA -- sin el encabezado/logo institucional
// de ACUAVALLE (eso lo pone cada contratista despues, en el espacio que se
// deja libre arriba).
const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel,
  AlignmentType, WidthType, ShadingType, BorderStyle, ImageRun, VerticalAlign,
} = require("docx");

const AZUL_OSCURO = "1F4E78";
const GRIS_CLARO = "F2F2F2";
const BORDE = { style: BorderStyle.SINGLE, size: 4, color: "999999" };
const BORDES_CELDA = { top: BORDE, bottom: BORDE, left: BORDE, right: BORDE };

function fmtCOP(n) {
  return "$ " + Math.round(Number(n) || 0).toLocaleString("es-CO");
}
function fmtFecha(iso) {
  if (!iso) return "";
  const [y, m, d] = String(iso).slice(0, 10).split("-");
  return y && m && d ? `${d}/${m}/${y}` : String(iso);
}
function fmtNum(n) {
  return Number(n || 0).toLocaleString("es-CO", { maximumFractionDigits: 2 });
}

function celda(texto, opts) {
  opts = opts || {};
  const runs = Array.isArray(texto) ? texto : [texto];
  return new TableCell({
    width: { size: opts.width || 2000, type: WidthType.DXA },
    shading: opts.bg ? { type: ShadingType.CLEAR, fill: opts.bg } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    borders: BORDES_CELDA,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    columnSpan: opts.colSpan,
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: runs.map((t) => new TextRun({ text: String(t == null ? "" : t), bold: !!opts.bold, size: opts.size || 18, color: opts.color })),
    })],
  });
}

function tituloSeccion(texto) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 120 },
    children: [new TextRun({ text: texto, bold: true, color: AZUL_OSCURO, size: 24 })],
  });
}

function parrafo(texto, opts) {
  opts = opts || {};
  return new Paragraph({
    spacing: { after: opts.after ?? 120 },
    alignment: opts.align,
    children: [new TextRun({ text: texto, size: opts.size || 20, italics: !!opts.italic, bold: !!opts.bold })],
  });
}

// ---------- Encabezado ----------
function bloqueEncabezado(datos) {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "CCCCCC" } },
      children: [new TextRun({ text: "[ ESPACIO PARA LOGO Y ENCABEZADO DEL CONTRATISTA ]", italics: true, color: "999999", size: 18 })],
    }),
    new Paragraph({ spacing: { after: 200 }, children: [] }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
      children: [new TextRun({
        text: datos.tipoInforme === "Final" ? "INFORME FINAL" : `INFORME PARCIAL No. ${datos.numeroParcial || ""}`,
        bold: true, size: 28, color: AZUL_OSCURO,
      })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 20 },
      children: [new TextRun({ text: "CONTRATISTA", bold: true, size: 20, color: AZUL_OSCURO })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: datos.contratista || "", bold: true, size: 20 })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 60 },
      children: [new TextRun({ text: `CONTRATO DE OBRA No. ${datos.numeroContrato || ""}`, bold: true, size: 22 })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: datos.objeto || "", size: 20 })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
      children: [new TextRun({ text: `PERIODO REPORTADO: ${fmtFecha(datos.fechaDesde)} — ${fmtFecha(datos.fechaHasta)}`, bold: true, size: 20 })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 300 },
      children: [new TextRun({ text: `Fecha del informe: ${fmtFecha(new Date().toISOString())}`, size: 18, color: "666666" })],
    }),
  ];
}

// ---------- 1. Descripcion de la necesidad (generada por IA) ----------
function seccionNecesidad(datos) {
  return [parrafo(datos.descripcionNecesidad || "[Agregar aquí la descripción de la necesidad que da origen al proyecto y lo que se soluciona con su ejecución.]")];
}

// ---------- 2. Informacion del contrato ----------
function tablaInfoContrato(datos) {
  const filas = [
    ["Contrato de Obra No.", datos.numeroContrato],
    ["Objeto", datos.objeto],
    ["Contratista", datos.contratista],
    ["Valor inicial", fmtCOP(datos.valorInicial)],
    ["Plazo inicial", datos.plazo],
    ["Fecha Acta de Inicio", fmtFecha(datos.fechaActaInicio)],
    ["Fecha terminación inicial", fmtFecha(datos.fechaTerminacionInicial)],
  ];
  if (datos.nuevaFechaTerminacion) filas.push(["Nueva fecha de terminación", fmtFecha(datos.nuevaFechaTerminacion)]);

  return new Table({
    width: { size: 9500, type: WidthType.DXA },
    columnWidths: [3200, 6300],
    rows: filas.map(([label, valor]) => new TableRow({
      children: [
        celda(label, { width: 3200, bold: true, bg: GRIS_CLARO }),
        celda(valor, { width: 6300 }),
      ],
    })),
  });
}

// ---------- 3. Resumen de actividades ejecutadas ----------
// Narrativo: lista las actividades (items) que se estan cobrando en el
// periodo de este acta -- los items del presupuesto con cantidad
// ejecutada > 0 dentro del rango de fechas -- usando la descripcion de
// COMO se ejecuto cada una (generada por IA en server.js;
// "descripcionEjecucion") en vez del simple nombre del item.
function resumenActividadesNarrativo(items) {
  const delPeriodo = (items || []).filter((it) => Number(it.cantidadEjecutadaPeriodo) > 0);
  if (!delPeriodo.length) {
    return [parrafo("No se registró ejecución de actividades dentro del periodo reportado.", { italic: true })];
  }
  const bloques = [parrafo("Durante el periodo reportado se ejecutaron y se cobran en la presente acta las siguientes actividades:")];
  delPeriodo.forEach((it) => {
    bloques.push(new Paragraph({
      indent: { left: 300 },
      spacing: { after: 80 },
      children: [new TextRun({
        text: `- Item ${it.item} (${fmtNum(it.cantidadEjecutadaPeriodo)} ${it.unidad}) — ${it.descripcionEjecucion || it.descripcion}`,
        size: 20,
      })],
    }));
  });
  return bloques;
}

// Cuadro de avance en PORCENTAJE (no cantidad): Contratado siempre es
// 100% (es la base contra la que se miden los demas); si un item no tiene
// cantidad contratada (agregado desde la app, sin match en el
// presupuesto oficial) no se puede calcular % sobre una base de 0, se deja
// "-" en vez de una division por cero.
function pctStr(numerador, base) {
  if (!base) return "-";
  return `${fmtNum((Number(numerador || 0) / base) * 100)}%`;
}

function tablaResumenPorcentaje(items) {
  const header = new TableRow({
    tableHeader: true,
    children: [
      celda("Item", { width: 700, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
      celda("Descripción", { width: 3700, bold: true, bg: AZUL_OSCURO, color: "FFFFFF" }),
      celda("Contratado", { width: 1250, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
      celda("Ejecutado este periodo", { width: 1450, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
      celda("Ejecutado acumulado", { width: 1450, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
      celda("Saldo por ejecutar", { width: 1350, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
    ],
  });
  const filas = (items || []).map((it) => {
    const contratada = Number(it.cantidadContratada || 0);
    const pctAcumulado = contratada ? (Number(it.cantidadEjecutadaAcumulada || 0) / contratada) * 100 : null;
    return new TableRow({
      children: [
        celda(it.item, { width: 700, align: AlignmentType.CENTER }),
        celda(it.descripcion, { width: 3700 }),
        celda(contratada ? "100%" : "-", { width: 1250, align: AlignmentType.CENTER }),
        celda(pctStr(it.cantidadEjecutadaPeriodo, contratada), { width: 1450, align: AlignmentType.CENTER }),
        celda(pctStr(it.cantidadEjecutadaAcumulada, contratada), { width: 1450, align: AlignmentType.CENTER }),
        celda(pctAcumulado == null ? "-" : `${fmtNum(Math.max(0, 100 - pctAcumulado))}%`, { width: 1350, align: AlignmentType.CENTER }),
      ],
    });
  });
  return new Table({
    width: { size: 9900, type: WidthType.DXA },
    columnWidths: [700, 3700, 1250, 1450, 1450, 1350],
    rows: [header].concat(filas),
  });
}

// "Grafica" de avance por capitulo dentro de cada direccion: una barra de
// progreso hecha con caracteres de bloque (texto, no imagen) -- se evita
// a proposito depender de una libreria de graficos con bindings nativos
// (fragil de instalar en el build de Render) o de un servicio externo de
// graficas (enviaria datos del proyecto a un tercero); esto siempre
// renderiza bien en Word sin dependencias extra. El avance se calcula por
// VALOR (cantidad x vr. unitario), no por cantidad cruda, porque un mismo
// capitulo mezcla items con distintas unidades (M2, M3, Un...) que no se
// pueden sumar directamente.
function barraTexto(pct) {
  const p = Math.max(0, Math.min(100, pct || 0));
  const llenos = Math.round((p / 100) * 20);
  return "█".repeat(llenos) + "░".repeat(20 - llenos);
}

function avancePorCapituloYDireccion(items) {
  const porDireccion = {};
  const ordenDireccion = [];
  (items || []).forEach((it) => {
    const direccion = it.direccion || "(sin dirección)";
    const capitulo = it.capitulo || "(sin capítulo)";
    if (!porDireccion[direccion]) { porDireccion[direccion] = {}; ordenDireccion.push(direccion); }
    if (!porDireccion[direccion][capitulo]) porDireccion[direccion][capitulo] = { vrContratado: 0, vrEjecutado: 0 };
    porDireccion[direccion][capitulo].vrContratado += Number(it.cantidadContratada || 0) * Number(it.vrUnitario || 0);
    porDireccion[direccion][capitulo].vrEjecutado += Number(it.vrEjecutadoAcumulado || 0);
  });
  return ordenDireccion
    .map((direccion) => ({
      direccion,
      capitulos: Object.keys(porDireccion[direccion])
        .filter((capitulo) => porDireccion[direccion][capitulo].vrContratado > 0)
        .map((capitulo) => {
          const c = porDireccion[direccion][capitulo];
          return { capitulo, pct: (c.vrEjecutado / c.vrContratado) * 100 };
        }),
    }))
    .filter((g) => g.capitulos.length);
}

function seccionAvancePorDireccion(items) {
  const grupos = avancePorCapituloYDireccion(items);
  if (!grupos.length) return [];
  const bloques = [parrafo("Avance de ejecución acumulado por capítulo, dentro de cada dirección/frente de obra (calculado sobre el valor, no la cantidad, de cada actividad):", { after: 140 })];
  grupos.forEach((g) => {
    bloques.push(new Paragraph({ spacing: { before: 160, after: 60 }, children: [new TextRun({ text: g.direccion, bold: true, size: 20 })] }));
    g.capitulos.forEach((c) => {
      bloques.push(new Paragraph({
        indent: { left: 200 },
        spacing: { after: 40 },
        children: [
          new TextRun({ text: barraTexto(c.pct), font: "Courier New", size: 18, color: AZUL_OSCURO }),
          new TextRun({ text: `  ${fmtNum(c.pct)}%  ${c.capitulo}`, size: 18 }),
        ],
      }));
    });
  });
  return bloques;
}

// ---------- 4. Balance financiero ----------
// SUMAS IGUALES: se arma con "Acumulado ejecutado a la fecha" (no con
// "Valor Acta del periodo") para que la cuenta cuadre SIEMPRE contra el
// valor del contrato, sin importar si este es el informe No. 1 o uno
// posterior -- en el primer informe ambos numeros coinciden (por eso en un
// informe No. 1 puede no notarse la diferencia), pero a partir del segundo
// ya no son iguales (el acta de cada periodo es solo el incremento, el
// acumulado es la suma de todos los periodos hasta la fecha).
function tablaBalanceFinanciero(b) {
  const filas = [
    ["Valor total del contrato", fmtCOP(b.valorContrato), ""],
    ["Anticipo pagado (no suma)", "", fmtCOP(b.valorAnticipoPagado)],
    [`Valor Acta ${b.tipoInforme === "Final" ? "Final" : "No. " + b.numeroParcial} (periodo reportado)`, "", fmtCOP(b.valorActaPeriodo)],
    [`Amortización anticipo (${fmtNum(b.porcentajeAnticipo)}% del valor del acta, no suma)`, "", fmtCOP(b.amortizacionAnticipo)],
    ["Valor neto a pagar en esta acta (no suma)", "", fmtCOP(b.valorNetoAPagar)],
    ["Acumulado ejecutado a la fecha", "", fmtCOP(b.acumuladoEjecutado)],
    ["Saldo por ejecutar", "", fmtCOP(b.saldoPorEjecutar)],
    ["Saldo a favor de la entidad contratante", "", fmtCOP(b.saldoAFavor)],
    ["SUMAS IGUALES", fmtCOP(b.valorContrato), fmtCOP(b.acumuladoEjecutado + b.saldoPorEjecutar + b.saldoAFavor)],
  ];
  return new Table({
    width: { size: 9500, type: WidthType.DXA },
    columnWidths: [5500, 2000, 2000],
    rows: filas.map(([label, izq, der], i) => {
      const esTotal = i === filas.length - 1;
      return new TableRow({
        children: [
          celda(label, { width: 5500, bold: esTotal, bg: esTotal ? GRIS_CLARO : undefined }),
          celda(izq, { width: 2000, align: AlignmentType.RIGHT, bold: esTotal, bg: esTotal ? GRIS_CLARO : undefined }),
          celda(der, { width: 2000, align: AlignmentType.RIGHT, bold: esTotal, bg: esTotal ? GRIS_CLARO : undefined }),
        ],
      });
    }),
  });
}

// Amortizacion del anticipo, en porcentaje del anticipo pagado (no del
// valor del contrato): cuanto ya se descontó de forma acumulada hasta
// esta acta, y cuanto queda pendiente.
function tablaAmortizacion(b) {
  const filas = [
    ["Valor total del anticipo pagado", fmtCOP(b.valorAnticipoPagado), "100%"],
    ["Amortizado acumulado a la fecha", fmtCOP(b.amortizacionAcumulada), `${fmtNum(b.pctAmortizado)}%`],
    ["Pendiente por amortizar", fmtCOP(Math.max(0, b.valorAnticipoPagado - b.amortizacionAcumulada)), `${fmtNum(b.pctPendienteAmortizar)}%`],
  ];
  return new Table({
    width: { size: 9500, type: WidthType.DXA },
    columnWidths: [5500, 2500, 1500],
    rows: filas.map(([label, valor, pct]) => new TableRow({
      children: [
        celda(label, { width: 5500, bold: true, bg: GRIS_CLARO }),
        celda(valor, { width: 2500, align: AlignmentType.RIGHT }),
        celda(pct, { width: 1500, align: AlignmentType.CENTER }),
      ],
    })),
  });
}

// ---------- 5. Garantias ----------
function tablaPolizas(datos) {
  const filas = [
    ["Compañía aseguradora", datos.companiaAseguradora],
    ["N° póliza de cumplimiento", datos.polizaCumplimiento],
    ["N° póliza de responsabilidad civil extracontractual", datos.polizaResponsabilidadCivil],
  ];
  return new Table({
    width: { size: 9500, type: WidthType.DXA },
    columnWidths: [4500, 5000],
    rows: filas.map(([label, valor]) => new TableRow({
      children: [celda(label, { width: 4500, bold: true, bg: GRIS_CLARO }), celda(valor, { width: 5000 })],
    })),
  });
}

function tablaAmparos(amparos) {
  const header = new TableRow({
    children: [
      celda("Amparo", { width: 3200, bold: true, bg: AZUL_OSCURO, color: "FFFFFF" }),
      celda("% del valor", { width: 1300, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
      celda("Valor asegurado", { width: 2000, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
      celda("Vigencia desde", { width: 1500, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
      celda("Vigencia hasta", { width: 1500, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
    ],
  });
  const filas = (amparos || []).map((a) => new TableRow({
    children: [
      celda(a.tipo, { width: 3200 }),
      celda(a.porcentaje !== "" && a.porcentaje != null ? `${a.porcentaje}%` : "", { width: 1300, align: AlignmentType.CENTER }),
      celda(a.valorAsegurado ? fmtCOP(a.valorAsegurado) : "", { width: 2000, align: AlignmentType.CENTER }),
      celda(fmtFecha(a.vigenciaDesde), { width: 1500, align: AlignmentType.CENTER }),
      celda(fmtFecha(a.vigenciaHasta), { width: 1500, align: AlignmentType.CENTER }),
    ],
  }));
  return new Table({
    width: { size: 9500, type: WidthType.DXA },
    columnWidths: [3200, 1300, 2000, 1500, 1500],
    rows: [header].concat(filas.length ? filas : [new TableRow({ children: [celda("Sin amparos registrados", { width: 9500, colSpan: 5, align: AlignmentType.CENTER })] })]),
  });
}

// ---------- 7. Cronologia y 8. Registro fotografico ----------
async function descargarImagenDrive_(fotoUrl) {
  const m = (fotoUrl || "").match(/\/d\/([^/]+)\//);
  if (!m) return null;
  try {
    const r = await fetch(`https://drive.google.com/thumbnail?id=${m[1]}&sz=w800`);
    if (!r.ok) return null;
    return Buffer.from(await r.arrayBuffer());
  } catch (e) {
    return null;
  }
}

async function seccionFotos(fotos) {
  if (!fotos || !fotos.length) {
    return [parrafo("No se registraron fotos dentro del periodo reportado.", { italic: true })];
  }
  // Agrupadas por direccion + item, un titulo por grupo.
  const grupos = {};
  const orden = [];
  fotos.forEach((f) => {
    const clave = `${f.direccion} — ${f.item} ${f.descripcion || ""}`.trim();
    if (!grupos[clave]) { grupos[clave] = []; orden.push(clave); }
    grupos[clave].push(f);
  });

  const bloques = [];
  for (const clave of orden) {
    bloques.push(new Paragraph({ spacing: { before: 200, after: 80 }, children: [new TextRun({ text: clave, bold: true, size: 20 })] }));
    for (const f of grupos[clave]) {
      const buf = await descargarImagenDrive_(f.fotoUrl);
      if (!buf) {
        bloques.push(parrafo(`(no se pudo incrustar la foto: ${f.fotoUrl})`, { italic: true, size: 16 }));
        continue;
      }
      bloques.push(new Paragraph({
        spacing: { after: 160 },
        children: [new ImageRun({ type: "jpg", data: buf, transformation: { width: 380, height: 285 } })],
      }));
    }
  }
  return bloques;
}

// ---------- Documento completo ----------
async function generarInformeContratistaDocx(datos, balance, items, fotos) {
  const children = [];
  children.push(...bloqueEncabezado(datos));

  children.push(tituloSeccion("1. Descripción de la necesidad"));
  children.push(...seccionNecesidad(datos));

  children.push(tituloSeccion("2. Información del contrato"));
  children.push(tablaInfoContrato(datos));

  children.push(tituloSeccion("3. Resumen de actividades ejecutadas"));
  children.push(...resumenActividadesNarrativo(items));
  children.push(new Paragraph({ spacing: { before: 160, after: 80 }, children: [] }));
  children.push(tablaResumenPorcentaje(items));
  children.push(new Paragraph({ spacing: { before: 200, after: 80 }, children: [] }));
  children.push(...seccionAvancePorDireccion(items));

  children.push(tituloSeccion("4. Balance financiero del contrato"));
  children.push(tablaBalanceFinanciero(balance));
  children.push(new Paragraph({ spacing: { before: 200, after: 80 }, children: [new TextRun({ text: "Amortización del anticipo", bold: true, size: 20 })] }));
  children.push(tablaAmortizacion(balance));

  children.push(tituloSeccion("5. Seguimiento y control de las garantías exigidas"));
  children.push(tablaPolizas(datos));
  children.push(new Paragraph({ spacing: { before: 160, after: 80 }, children: [] }));
  children.push(tablaAmparos(datos.amparos));

  children.push(tituloSeccion("6. Observaciones"));
  children.push(parrafo("[Agregar aquí observaciones adicionales, si aplica.]", { italic: true }));

  children.push(tituloSeccion("7. Cronología del periodo reportado"));
  children.push(parrafo(`Durante el periodo comprendido entre el ${fmtFecha(datos.fechaDesde)} y el ${fmtFecha(datos.fechaHasta)}, se ejecutaron las actividades descritas en el numeral 3.`));

  children.push(tituloSeccion("8. Registro fotográfico"));
  children.push(...(await seccionFotos(fotos)));

  children.push(new Paragraph({ spacing: { before: 600 }, children: [] }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    border: { top: { style: BorderStyle.SINGLE, size: 6, color: "000000" } },
    spacing: { before: 400 },
    children: [new TextRun({ text: datos.contratista || "", bold: true, size: 20 })],
  }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "CONTRATISTA", size: 18 })] }));

  const doc = new Document({
    sections: [{ properties: {}, children }],
  });
  return doc;
}

module.exports = { generarInformeContratistaDocx };
