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
      spacing: { after: 20 },
      children: [new TextRun({ text: "SUPERVISOR", bold: true, size: 20, color: AZUL_OSCURO })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: datos.supervisor || "", bold: true, size: 20 })],
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
  const texto = datos.descripcionNecesidad || "[Agregar aquí la descripción de la necesidad que da origen al proyecto y lo que se soluciona con su ejecución.]";
  // El texto largo generado por la IA trae parrafos separados por linea en
  // blanco -- se respetan como parrafos de Word aparte (un solo TextRun
  // con "\n" adentro no los separa visualmente).
  return texto.split(/\n\s*\n/).map((p) => p.trim()).filter((p) => p).map((p) => parrafo(p));
}

// ---------- 2. Informacion del contrato ----------
function tablaInfoContrato(datos) {
  const filas = [
    ["Contrato de Obra No.", datos.numeroContrato],
    ["Objeto", datos.objeto],
    ["Contratista", datos.contratista],
    ["Supervisor", datos.supervisor],
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

// Cuadro de avance por CANTIDAD (no porcentaje) -- el porcentaje ya se ve
// en el panel de barras por direccion/capitulo mas abajo, asi que este
// cuadro vuelve a mostrar los numeros crudos de cada item.
function tablaResumenActividades(items) {
  const header = new TableRow({
    tableHeader: true,
    children: [
      celda("Item", { width: 700, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
      celda("Descripción", { width: 3400, bold: true, bg: AZUL_OSCURO, color: "FFFFFF" }),
      celda("Und.", { width: 700, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
      celda("Cant. contratada", { width: 1200, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
      celda("Ejecutado este periodo", { width: 1300, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
      celda("Ejecutado acumulado", { width: 1300, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
      celda("Saldo por ejecutar", { width: 1300, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
    ],
  });
  const filas = (items || []).map((it) => new TableRow({
    children: [
      celda(it.item, { width: 700, align: AlignmentType.CENTER }),
      celda(it.descripcion, { width: 3400 }),
      celda(it.unidad, { width: 700, align: AlignmentType.CENTER }),
      celda(fmtNum(it.cantidadContratada), { width: 1200, align: AlignmentType.CENTER }),
      celda(fmtNum(it.cantidadEjecutadaPeriodo), { width: 1300, align: AlignmentType.CENTER }),
      celda(fmtNum(it.cantidadEjecutadaAcumulada), { width: 1300, align: AlignmentType.CENTER }),
      celda(fmtNum(Math.max(0, Number(it.cantidadContratada || 0) - Number(it.cantidadEjecutadaAcumulada || 0))), { width: 1300, align: AlignmentType.CENTER }),
    ],
  }));
  return new Table({
    width: { size: 9900, type: WidthType.DXA },
    columnWidths: [700, 3400, 700, 1200, 1300, 1300, 1300],
    rows: [header].concat(filas),
  });
}

// "Grafica" de avance por capitulo dentro de cada direccion: en vez de una
// imagen de grafica (necesitaria una libreria con bindings nativos fragil
// de instalar en el build de Render, o mandar los datos del proyecto a un
// servicio externo de graficas), se arma con TABLAS y celdas coloreadas
// (una barra real verde/gris, no texto) -- Word soporta tablas anidadas
// dentro de una celda, asi que cada fila de capitulo lleva su propia
// barrita adentro. El avance se calcula por VALOR (cantidad x vr.
// unitario), no por cantidad cruda, porque un mismo capitulo mezcla items
// con distintas unidades (M2, M3, Un...) que no se pueden sumar
// directamente.
const VERDE_AVANCE = "70AD47";
const GRIS_AVANCE = "D9D9D9";

function agruparAvancePorDireccion(items) {
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
    .map((direccion) => {
      const capitulos = Object.keys(porDireccion[direccion])
        .filter((capitulo) => porDireccion[direccion][capitulo].vrContratado > 0)
        .map((capitulo) => {
          const c = porDireccion[direccion][capitulo];
          return { capitulo, vrContratado: c.vrContratado, vrEjecutado: c.vrEjecutado };
        });
      const vrContratado = capitulos.reduce((s, c) => s + c.vrContratado, 0);
      const vrEjecutado = capitulos.reduce((s, c) => s + c.vrEjecutado, 0);
      return { direccion, capitulos, vrContratado, vrEjecutado };
    })
    .filter((g) => g.capitulos.length);
}

// Barra horizontal real (tabla de 1 fila x 2 celdas, verde/gris segun el
// %) del ancho que se le pida -- se usa tanto suelta (barra grande de cada
// direccion) como anidada dentro de una celda (barra chica por capitulo).
function barra(pct, anchoTotal) {
  const p = Math.max(0, Math.min(100, pct || 0));
  const anchoVerde = Math.max(0, Math.round((anchoTotal * p) / 100));
  const anchoGris = Math.max(0, anchoTotal - anchoVerde);
  const columnas = [];
  const celdas = [];
  if (anchoVerde > 0) {
    columnas.push(anchoVerde);
    celdas.push(new TableCell({
      width: { size: anchoVerde, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: VERDE_AVANCE },
      margins: { top: 30, bottom: 30, left: 0, right: 0 },
      children: [new Paragraph({ children: [] })],
    }));
  }
  if (anchoGris > 0) {
    columnas.push(anchoGris);
    celdas.push(new TableCell({
      width: { size: anchoGris, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: GRIS_AVANCE },
      margins: { top: 30, bottom: 30, left: 0, right: 0 },
      children: [new Paragraph({ children: [] })],
    }));
  }
  return new Table({
    width: { size: anchoTotal, type: WidthType.DXA },
    columnWidths: columnas,
    rows: [new TableRow({ children: celdas })],
  });
}

// Celda de KPI (etiqueta chica arriba, valor grande abajo) para el panel
// resumen de cada direccion.
function celdaKpi(etiqueta, valor, opts) {
  opts = opts || {};
  return new TableCell({
    width: { size: opts.width || 3100, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: opts.bg || "F7F7F7" },
    verticalAlign: VerticalAlign.CENTER,
    borders: BORDES_CELDA,
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    children: [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: etiqueta, size: 15, color: "666666" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: valor, bold: true, size: 22, color: opts.color || "000000" })] }),
    ],
  });
}

function panelResumenDireccion(g) {
  const pct = g.vrContratado > 0 ? (g.vrEjecutado / g.vrContratado) * 100 : 0;
  const pendiente = Math.max(0, g.vrContratado - g.vrEjecutado);
  const ancho = 9500;

  const bloques = [];
  bloques.push(new Paragraph({
    shading: { type: ShadingType.CLEAR, fill: AZUL_OSCURO },
    spacing: { before: 240, after: 100 },
    children: [new TextRun({ text: `  DIRECCIÓN: ${g.direccion}`, bold: true, color: "FFFFFF", size: 22 })],
  }));

  bloques.push(new Table({
    width: { size: ancho, type: WidthType.DXA },
    columnWidths: [3167, 3167, 3166],
    rows: [new TableRow({
      children: [
        celdaKpi("CONTRATADO", fmtCOP(g.vrContratado), { width: 3167 }),
        celdaKpi("EJECUTADO AL CORTE", fmtCOP(g.vrEjecutado), { width: 3167, bg: "EAF3E5", color: "375623" }),
        celdaKpi("PENDIENTE POR EJECUTAR", fmtCOP(pendiente), { width: 3166, bg: "F2F2F2", color: "595959" }),
      ],
    })],
  }));

  bloques.push(new Paragraph({ spacing: { before: 100, after: 40 }, children: [] }));
  bloques.push(barra(pct, ancho));
  bloques.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 40, after: 160 },
    children: [new TextRun({ text: `${fmtCOP(g.vrEjecutado)} ejecutado (${fmtNum(pct)}%)  —  ${fmtCOP(pendiente)} pendiente (${fmtNum(Math.max(0, 100 - pct))}%)`, size: 16, color: "666666" })],
  }));

  const anchoBarraCapitulo = 1600;
  const header = new TableRow({
    tableHeader: true,
    children: [
      celda("Capítulo", { width: 2400, bold: true, bg: AZUL_OSCURO, color: "FFFFFF" }),
      celda("Contratado", { width: 1900, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
      celda("Ejecutado al corte", { width: 1900, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
      celda("Pendiente", { width: 1700, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
      celda("Avance", { width: anchoBarraCapitulo, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
    ],
  });
  const filas = g.capitulos.map((c) => {
    const cpct = c.vrContratado > 0 ? (c.vrEjecutado / c.vrContratado) * 100 : 0;
    return new TableRow({
      children: [
        celda(c.capitulo, { width: 2400 }),
        celda(fmtCOP(c.vrContratado), { width: 1900, align: AlignmentType.RIGHT }),
        celda(fmtCOP(c.vrEjecutado), { width: 1900, align: AlignmentType.RIGHT }),
        celda(fmtCOP(Math.max(0, c.vrContratado - c.vrEjecutado)), { width: 1700, align: AlignmentType.RIGHT }),
        new TableCell({
          width: { size: anchoBarraCapitulo, type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          borders: BORDES_CELDA,
          margins: { top: 60, bottom: 60, left: 80, right: 80 },
          children: [
            barra(cpct, anchoBarraCapitulo - 200),
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 30 }, children: [new TextRun({ text: `${fmtNum(cpct)}%`, size: 14 })] }),
          ],
        }),
      ],
    });
  });
  const filaTotal = new TableRow({
    children: [
      celda("TOTAL DIRECCIÓN", { width: 2400, bold: true, bg: GRIS_CLARO }),
      celda(fmtCOP(g.vrContratado), { width: 1900, align: AlignmentType.RIGHT, bold: true, bg: GRIS_CLARO }),
      celda(fmtCOP(g.vrEjecutado), { width: 1900, align: AlignmentType.RIGHT, bold: true, bg: GRIS_CLARO }),
      celda(fmtCOP(pendiente), { width: 1700, align: AlignmentType.RIGHT, bold: true, bg: GRIS_CLARO }),
      celda(`${fmtNum(pct)}%`, { width: anchoBarraCapitulo, align: AlignmentType.CENTER, bold: true, bg: GRIS_CLARO }),
    ],
  });

  bloques.push(new Table({
    width: { size: ancho, type: WidthType.DXA },
    columnWidths: [2400, 1900, 1900, 1700, anchoBarraCapitulo],
    rows: [header].concat(filas).concat([filaTotal]),
  }));

  return bloques;
}

function seccionAvancePorDireccion(items) {
  const grupos = agruparAvancePorDireccion(items);
  if (!grupos.length) return [];

  const bloques = [parrafo("Resumen de ejecución por dirección y capítulo, comparando lo contratado con lo ejecutado y lo pendiente (calculado sobre el valor, no la cantidad, de cada actividad):", { after: 100 })];
  grupos.forEach((g) => { bloques.push(...panelResumenDireccion(g)); });

  const totalContratado = grupos.reduce((s, g) => s + g.vrContratado, 0);
  const totalEjecutado = grupos.reduce((s, g) => s + g.vrEjecutado, 0);
  const totalPendiente = Math.max(0, totalContratado - totalEjecutado);
  const totalPct = totalContratado > 0 ? (totalEjecutado / totalContratado) * 100 : 0;

  bloques.push(new Paragraph({
    shading: { type: ShadingType.CLEAR, fill: "000000" },
    spacing: { before: 260, after: 100 },
    children: [new TextRun({ text: "  TOTAL GENERAL DE LA OBRA", bold: true, color: "FFFFFF", size: 22 })],
  }));
  bloques.push(new Table({
    width: { size: 9500, type: WidthType.DXA },
    columnWidths: [3167, 3167, 3166],
    rows: [new TableRow({
      children: [
        celdaKpi("TOTAL CONTRATADO", fmtCOP(totalContratado), { width: 3167 }),
        celdaKpi("TOTAL EJECUTADO AL CORTE", fmtCOP(totalEjecutado), { width: 3167, bg: "EAF3E5", color: "375623" }),
        celdaKpi("TOTAL PENDIENTE", fmtCOP(totalPendiente), { width: 3166, bg: "F2F2F2", color: "595959" }),
      ],
    })],
  }));
  bloques.push(new Paragraph({ spacing: { before: 100, after: 40 }, children: [] }));
  bloques.push(barra(totalPct, 9500));
  bloques.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 40, after: 160 },
    children: [new TextRun({ text: `${fmtNum(totalPct)}% ejecutado — ${fmtNum(Math.max(0, 100 - totalPct))}% pendiente`, size: 16, color: "666666" })],
  }));

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
  children.push(tablaResumenActividades(items));
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
