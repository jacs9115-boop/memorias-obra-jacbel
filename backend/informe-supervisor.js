// Genera el "Informe de Supervisión" (.docx), formato AP2-FO-024 de
// Acuavalle S.A. E.S.P., EL MISMO documento que hoy elabora el usuario como
// supervisor de la obra -- a diferencia del Informe del Contratista
// (informe.js), este SI lleva el logo y encabezado institucional reales,
// porque es el documento oficial que se entrega a la entidad.
//
// Reutiliza del Informe del Contratista (informe.js) todo lo que ya es
// identico entre ambos documentos: estilos base, balance financiero,
// polizas/amparos, y el registro fotografico -- la unica informacion nueva
// que hace falta es la que es propia del formato de supervision (encabezado
// institucional con logo, tabla de "actividades generales" SI/NO, y la
// narrativa de cronologia en primera persona del supervisor).
const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer,
  AlignmentType, WidthType, BorderStyle, ImageRun, VerticalAlign, PageBreak,
  PageNumber, UnderlineType,
} = require("docx");
const fs = require("fs");
const path = require("path");

const {
  AZUL_OSCURO, GRIS_CLARO, BORDE, BORDES_CELDA, SIN_BORDE_,
  fmtCOP, fmtFecha, fmtNum,
  celda, tituloSeccion, parrafo,
  tablaBalanceFinanciero, tablaPolizas, tablaAmparos,
  seccionFotos,
} = require("./informe");

const LOGO_PATH_ = path.join(__dirname, "assets", "acuavalle-logo.png");
// Color del formato AP2-FO-024 real (los numeros/datos variables del
// contrato van en este rojo, las etiquetas fijas en negro) -- ver la
// portada real que sirvio de referencia.
const ROJO_ = "C00000";
const FUENTE_ = "Arial";

// Los 8 tipos de "acta" del formato AP2-FO-024 -- el usuario marca cuales
// ocurrieron (SI/NO), con fecha y observaciones, en la pantalla de
// generacion del informe de supervision.
const EVENTOS_CONTRATO_DEFAULT_ = [
  "Reunión visita previa al sitio de la obra",
  "Reunión en las oficinas del Municipio",
  "Iniciación",
  "Suspensión",
  "Reinicio",
  "Recibo",
  "Liquidación",
  "Otros",
];

function mesAnioEs_(iso) {
  const meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
  if (!iso) return "";
  const [y, m] = String(iso).slice(0, 10).split("-");
  const mi = parseInt(m, 10) - 1;
  return mi >= 0 && mi < 12 ? `${meses[mi]} de ${y}` : "";
}

// ---------- Encabezado institucional (se repite en cada pagina) ----------
// Tabla de 3 celdas igual al formato real: logo solo | PROCESO DE
// CONTRATACION + INFORME... + CONTRATO... | Codigo/Version. El
// "SUBGERENCIA TECNICA" del membrete va en el cuerpo (bloqueTitulo_), no
// aca -- asi es como esta en el documento real de referencia.
function headerAcuavalle_(datos) {
  const logoBuf = fs.existsSync(LOGO_PATH_) ? fs.readFileSync(LOGO_PATH_) : null;
  const numeroInforme = datos.tipoInforme === "Final" ? "Final" : String(datos.numeroParcial || "");

  const celdaLogo = new TableCell({
    width: { size: 2200, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    borders: BORDES_CELDA,
    margins: { top: 100, bottom: 100, left: 100, right: 100 },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: logoBuf ? [new ImageRun({ type: "png", data: logoBuf, transformation: { width: 110, height: 110 } })] : [],
      }),
    ],
  });

  const celdaTitulo = new TableCell({
    width: { size: 7300, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    borders: BORDES_CELDA,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PROCESO DE CONTRATACION", bold: true, size: 18, font: FUENTE_ })] }),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { before: 60 },
        children: [
          new TextRun({ text: "INFORME DE SUPERVISION No. ", bold: true, size: 18, font: FUENTE_ }),
          new TextRun({ text: numeroInforme, bold: true, size: 18, color: ROJO_, font: FUENTE_ }),
        ],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { before: 40 },
        children: [
          new TextRun({ text: "CONTRATO DE OBRA No. ", bold: true, size: 18, font: FUENTE_ }),
          new TextRun({ text: datos.numeroContrato || "", bold: true, size: 18, color: ROJO_, font: FUENTE_ }),
        ],
      }),
    ],
  });

  const celdaCodigo = new TableCell({
    width: { size: 1600, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    borders: BORDES_CELDA,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    children: [
      new Paragraph({ children: [new TextRun({ text: "Código: AP2-FO-024", size: 15, font: FUENTE_ })] }),
      new Paragraph({ spacing: { before: 40 }, children: [new TextRun({ text: "Versión No.: 002", size: 15, font: FUENTE_ })] }),
    ],
  });

  return new Header({
    children: [
      new Table({
        width: { size: 11100, type: WidthType.DXA },
        columnWidths: [2200, 7300, 1600],
        rows: [new TableRow({ children: [celdaLogo, celdaTitulo, celdaCodigo] })],
      }),
    ],
  });
}

// Pie de pagina institucional (se repite en cada pagina): frase de
// propiedad + "Pagina X de Y" con los campos automaticos de Word, igual
// al documento real de referencia.
function footerAcuavalle_() {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Propiedad de ACUAVALLE S.A. E.S.P. – Prohibida su reproducción", size: 14, color: "808080", font: FUENTE_ })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "Página ", size: 14, color: "808080", font: FUENTE_ }),
          new TextRun({ children: [PageNumber.CURRENT], size: 14, color: "808080", font: FUENTE_ }),
          new TextRun({ text: " de ", size: 14, color: "808080", font: FUENTE_ }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 14, color: "808080", font: FUENTE_ }),
        ],
      }),
    ],
  });
}

// ---------- Portada (ocupa exactamente la pagina 1, ver el PageBreak al
// final del bloque en generarInformeSupervisorDocx) ----------
function bloqueTitulo_(datos) {
  const numeroInforme = datos.tipoInforme === "Final" ? "FINAL" : `PARCIAL No. ${datos.numeroParcial || ""}`;
  const cargoSupervisor = datos.supervisorCargo ? `, ${datos.supervisorCargo}.` : "";

  return [
    new Paragraph({
      spacing: { before: 200, after: 400 },
      children: [new TextRun({ text: "SUBGERENCIA TECNICA", bold: true, color: ROJO_, size: 24, font: FUENTE_ })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 400, after: 300 },
      children: [
        new TextRun({ text: "INFORME ", bold: true, size: 32, font: FUENTE_, underline: { type: UnderlineType.SINGLE } }),
        new TextRun({ text: `${numeroInforme} `, bold: true, size: 32, color: ROJO_, font: FUENTE_, underline: { type: UnderlineType.SINGLE } }),
        new TextRun({ text: "DE SUPERVISIÓN", bold: true, size: 32, font: FUENTE_, underline: { type: UnderlineType.SINGLE } }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 60 },
      children: [
        new TextRun({ text: "CONTRATO DE OBRA No. ", bold: true, size: 24, font: FUENTE_ }),
        new TextRun({ text: datos.numeroContrato || "", bold: true, size: 24, color: ROJO_, font: FUENTE_ }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 500 },
      children: [new TextRun({ text: (datos.objeto || "").toUpperCase(), bold: true, size: 22, color: ROJO_, font: FUENTE_ })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 300, after: 40 },
      children: [new TextRun({ text: "SUPERVISOR:", bold: true, size: 20, font: FUENTE_ })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 500 },
      children: [new TextRun({ text: `${datos.supervisor || ""}${cargoSupervisor}`, bold: true, size: 20, color: ROJO_, font: FUENTE_ })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 300 },
      children: [
        new TextRun({ text: "FECHA: Cali - Valle, ", size: 20, font: FUENTE_ }),
        new TextRun({ text: mesAnioEs_(datos.fechaHasta), size: 20, color: ROJO_, font: FUENTE_ }),
      ],
    }),
    // La portada ocupa toda la pagina 1 -- el numeral "1. Informacion del
    // contrato" arranca en la pagina 2 aunque la portada no llene la hoja.
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ---------- 1. Informacion del contrato ----------
function tablaInfoContratoSupervisor_(datos) {
  const filas = [
    ["Contrato de Obra No.", datos.numeroContrato],
    ["Objeto", datos.objeto],
    ["Contratista", datos.contratista],
    ["Valor inicial", fmtCOP(datos.valorInicial)],
    ["Plazo inicial", datos.plazo],
    ["Fecha de iniciación", fmtFecha(datos.fechaActaInicio)],
    ["Fecha de terminación", fmtFecha(datos.nuevaFechaTerminacion || datos.fechaTerminacionInicial)],
    ["Supervisor", datos.supervisor],
  ];
  if (datos.supervisorDesignacion) filas.push(["Designación del supervisor", datos.supervisorDesignacion]);

  return new Table({
    width: { size: 9500, type: WidthType.DXA },
    columnWidths: [3200, 6300],
    rows: filas.map(([label, valor]) => new TableRow({
      children: [celda(label, { width: 3200, bold: true, bg: GRIS_CLARO }), celda(valor, { width: 6300 })],
    })),
  });
}

// ---------- 2. Actividades generales realizadas durante el contrato ----------
function tablaActividadesGenerales_(eventos) {
  const lista = (eventos && eventos.length ? eventos : EVENTOS_CONTRATO_DEFAULT_.map((e) => ({ evento: e, ocurrio: false, fecha: "", observaciones: "" })));
  const header = new TableRow({
    tableHeader: true,
    children: [
      celda("Acta", { width: 3200, bold: true, bg: AZUL_OSCURO, color: "FFFFFF" }),
      celda("SI", { width: 500, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
      celda("NO", { width: 500, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
      celda("Fecha", { width: 1500, bold: true, bg: AZUL_OSCURO, color: "FFFFFF", align: AlignmentType.CENTER }),
      celda("Observaciones", { width: 3300, bold: true, bg: AZUL_OSCURO, color: "FFFFFF" }),
    ],
  });
  const filas = lista.map((e) => new TableRow({
    children: [
      celda(e.evento, { width: 3200 }),
      celda(e.ocurrio ? "X" : "", { width: 500, align: AlignmentType.CENTER, bold: true }),
      celda(e.ocurrio ? "" : "X", { width: 500, align: AlignmentType.CENTER, bold: true }),
      celda(e.fecha ? fmtFecha(e.fecha) : "", { width: 1500, align: AlignmentType.CENTER }),
      celda(e.observaciones || "", { width: 3300 }),
    ],
  }));
  return new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: [3200, 500, 500, 1500, 3300],
    rows: [header].concat(filas),
  });
}

// ---------- 3. Cronologia del contrato ----------
function seccionCronologia_(datos, items) {
  const bloques = [];
  bloques.push(parrafo(
    datos.textoCronologia ||
    `El supervisor y el contratista ${datos.contratista || ""} realizaron el seguimiento correspondiente al desarrollo del Contrato de Obra No. ${datos.numeroContrato || ""}, cuyo objeto consiste en ${(datos.objeto || "").toLowerCase()}. Durante las visitas de seguimiento se verificó el cumplimiento de las obligaciones contractuales, así como la afiliación del personal al sistema de seguridad social en salud, pensión y riesgos profesionales.`
  ));

  const delPeriodo = (items || []).filter((it) => Number(it.cantidadEjecutadaPeriodo) > 0);
  if (delPeriodo.length) {
    bloques.push(new Paragraph({ spacing: { before: 120, after: 80 }, children: [new TextRun({ text: "ACTIVIDADES DEL CONTRATISTA DE OBRA VERIFICADAS EN CAMPO:", bold: true, size: 20 })] }));
    delPeriodo.forEach((it) => {
      bloques.push(new Paragraph({
        indent: { left: 300 },
        spacing: { after: 80 },
        alignment: AlignmentType.JUSTIFIED,
        children: [new TextRun({
          text: `- Item ${it.item} (${fmtNum(it.cantidadEjecutadaPeriodo)} ${it.unidad}) — ${it.descripcionEjecucion || it.descripcion}`,
          size: 20,
        })],
      }));
    });
  }
  return bloques;
}

// ---------- Documento completo ----------
async function generarInformeSupervisorDocx(datos, balance, items, fotos) {
  const children = [];
  children.push(...bloqueTitulo_(datos));

  children.push(tituloSeccion("1. Información del contrato"));
  children.push(tablaInfoContratoSupervisor_(datos));

  children.push(tituloSeccion("2. Actividades generales realizadas durante el contrato"));
  children.push(tablaActividadesGenerales_(datos.eventosContrato));

  children.push(tituloSeccion("3. Cronología del contrato, terminación y recibo"));
  children.push(...seccionCronologia_(datos, items));

  children.push(tituloSeccion(`4. Balance financiero del contrato N° ${datos.numeroContrato || ""}`));
  children.push(tablaBalanceFinanciero(balance));

  children.push(tituloSeccion("5. Seguimiento y control de las garantías exigidas"));
  children.push(parrafo("En mi calidad de supervisor, y una vez realizado el seguimiento y control a las garantías, se deja constancia en la siguiente tabla de las mismas con sus respectivos amparos, valores y vigencias así:"));
  children.push(tablaPolizas(datos));
  children.push(new Paragraph({ spacing: { before: 160, after: 80 }, children: [] }));
  children.push(tablaAmparos(datos.amparos));

  children.push(tituloSeccion("6. Observaciones"));
  children.push(parrafo(datos.observacionesSupervisor || "[Agregar aquí observaciones adicionales sobre el desarrollo del contrato durante este periodo, si aplica.]", { italic: !datos.observacionesSupervisor }));

  children.push(tituloSeccion("7. Balance del cumplimiento del objeto contractual"));
  children.push(parrafo("En el resumen de la presente acta."));

  children.push(tituloSeccion("8. Registro fotográfico"));
  children.push(...(await seccionFotos(fotos)));

  children.push(new Paragraph({ spacing: { before: 600 }, children: [] }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    border: { top: { style: BorderStyle.SINGLE, size: 6, color: "000000" } },
    spacing: { before: 400 },
    children: [new TextRun({ text: datos.supervisor || "", bold: true, size: 20 })],
  }));
  if (datos.supervisorCargo) {
    children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: datos.supervisorCargo, size: 18 })] }));
  }
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ACUAVALLE S.A. E.S.P.", size: 18, bold: true })] }));

  const doc = new Document({
    styles: { default: { document: { run: { font: FUENTE_ } } } },
    sections: [{
      properties: {},
      headers: { default: headerAcuavalle_(datos) },
      footers: { default: footerAcuavalle_() },
      children,
    }],
  });
  return doc;
}

module.exports = { generarInformeSupervisorDocx, EVENTOS_CONTRATO_DEFAULT_ };
