import {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, TableOfContents, PageBreak,
} from "docx";

const NAVY = "1F2A44";
const ACCENT = "B7791F";
const GREY = "5A5A5A";
const HEADER_BG = "1F2A44";
const ZEBRA_BG = "F1F3F7";

// ---- helpers -------------------------------------------------------------
const h1 = (text) => new Paragraph({ text, heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 120 } });
const h2 = (text) => new Paragraph({ text, heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 80 } });

function body(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120, line: 276 },
    children: [new TextRun({ text: text ?? "", size: 21, color: "222222", ...opts })],
  });
}

function label(lbl, val) {
  return new Paragraph({
    spacing: { after: 40 },
    children: [
      new TextRun({ text: `${lbl}: `, bold: true, size: 21, color: NAVY }),
      new TextRun({ text: val ?? "—", size: 21, color: "222222" }),
    ],
  });
}

function codeBlock(code) {
  const lines = String(code || "").split("\n");
  return new Paragraph({
    spacing: { before: 40, after: 120 },
    shading: { fill: "F5F5F0" },
    border: { left: { style: BorderStyle.SINGLE, size: 18, color: ACCENT, space: 6 } },
    children: lines.flatMap((ln, i) => {
      const run = new TextRun({ text: ln, font: "Consolas", size: 18, color: "333333" });
      return i === 0 ? [run] : [new TextRun({ break: 1 }), run];
    }),
  });
}

function cell(text, { bold = false, color = "222222", bg, align = AlignmentType.LEFT, font } = {}) {
  return new TableCell({
    shading: bg ? { fill: bg } : undefined,
    margins: { top: 60, bottom: 60, left: 110, right: 110 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text: text ?? "", bold, size: 19, color, font })],
    })],
  });
}

function dataTable(headers, rows, widths) {
  const border = { style: BorderStyle.SINGLE, size: 4, color: "D5D9E0" };
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border },
    columnWidths: widths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((hh) => cell(hh, { bold: true, color: "FFFFFF", bg: HEADER_BG })),
      }),
      ...rows.map((r, ri) => new TableRow({
        children: r.map((c, ci) => cell(
          typeof c === "object" ? c.text : c,
          { bg: ri % 2 ? ZEBRA_BG : undefined, align: ci === 0 ? AlignmentType.LEFT : (typeof c === "object" ? c.align : AlignmentType.LEFT), font: typeof c === "object" ? c.font : undefined },
        )),
      })),
    ],
  });
}

// ---- document ------------------------------------------------------------
export function buildDocxDocument(docData, parsed) {
  const secById = Object.fromEntries((docData.secciones || []).map((s) => [s.id, s]));
  const tDescByName = Object.fromEntries(((secById.tablas_columnas?.items) || []).map((i) => [i.tabla, i]));
  const mDescByName = Object.fromEntries(((secById.medidas_dax?.items) || []).map((i) => [i.nombre, i]));
  const pDescByName = Object.fromEntries(((secById.paginas_reporte?.items) || []).map((i) => [i.pagina, i]));

  const children = [];

  // 1. PORTADA
  children.push(
    new Paragraph({ spacing: { before: 2200 }, children: [] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: docData.titulo || "Documentación Técnica del Modelo", bold: true, size: 52, color: NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ACCENT, space: 8 } }, children: [new TextRun({ text: "Power BI · Modelo de Datos", size: 24, color: ACCENT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 300, after: 30 }, children: [new TextRun({ text: parsed.fileName || "", size: 22, color: GREY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 30 }, children: [new TextRun({ text: docData.fecha || "", size: 20, color: GREY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200 }, children: [new TextRun({ text: `${parsed.tables.length} tablas · ${parsed.measures.length} medidas · ${parsed.relationships.length} relaciones · ${parsed.pages.length} páginas`, size: 20, color: NAVY, bold: true })] }),
    new Paragraph({ children: [new PageBreak()] }),
  );

  // 2. TABLA DE CONTENIDOS
  children.push(
    new Paragraph({ text: "Tabla de Contenidos", heading: HeadingLevel.HEADING_1, spacing: { after: 120 } }),
    new TableOfContents("Contenido", { hyperlink: true, headingStyleRange: "1-2" }),
    new Paragraph({ children: [new PageBreak()] }),
  );

  // 3. INTRODUCCIÓN
  children.push(h1("1. Introducción"));
  children.push(body(docData.introduccion || "Este documento describe la estructura y el funcionamiento del modelo de datos de Power BI, incluyendo sus fuentes, tablas, medidas, relaciones y páginas de reporte."));

  // 4. RESUMEN EJECUTIVO
  children.push(h1("2. Resumen Ejecutivo"));
  children.push(body(docData.resumen_ejecutivo));

  // 5. FUENTES DE DATOS
  children.push(h1("3. Fuentes de Datos"));
  if (secById.fuentes_datos?.contenido) children.push(body(secById.fuentes_datos.contenido));
  const fuentes = (secById.fuentes_datos?.items || []);
  if (fuentes.length) {
    for (const f of fuentes) {
      children.push(h2(f.nombre || "Fuente"));
      children.push(label("Tipo", f.tipo));
      if (f.descripcion) children.push(body(f.descripcion));
    }
  } else if (parsed.sources?.length) {
    children.push(dataTable(["Conexión", "Tipo"], parsed.sources.map((s) => [s.name, s.type]), [70, 30]));
  }

  // 6. ARQUITECTURA DEL MODELO
  children.push(h1("4. Arquitectura del Modelo"));
  children.push(body(secById.arquitectura_modelo?.contenido));

  // 7. TABLAS Y COLUMNAS (completo desde parsed)
  children.push(h1("5. Tablas y Columnas"));
  if (secById.tablas_columnas?.contenido) children.push(body(secById.tablas_columnas.contenido));
  for (const t of parsed.tables) {
    const d = tDescByName[t.name];
    children.push(h2(t.name + (d?.tipo ? `  —  ${d.tipo}` : "")));
    if (d?.descripcion) children.push(body(d.descripcion));
    if (t.columns.length) {
      children.push(dataTable(
        ["Columna", "Tipo de dato", "Calculada"],
        t.columns.map((c) => [c.name, c.dataType, { text: c.calculated ? "Sí" : "—", align: AlignmentType.CENTER }]),
        [50, 32, 18],
      ));
    } else {
      children.push(body("Sin columnas visibles.", { italics: true, color: GREY }));
    }
  }

  // 8. MEDIDAS DAX (completo desde parsed, agrupado por tabla)
  children.push(h1("6. Medidas DAX"));
  if (secById.medidas_dax?.contenido) children.push(body(secById.medidas_dax.contenido));
  const byTable = {};
  for (const m of parsed.measures) (byTable[m.table] ||= []).push(m);
  const tableNames = Object.keys(byTable).sort();
  if (!tableNames.length) children.push(body("El modelo no contiene medidas.", { italics: true, color: GREY }));
  for (const tn of tableNames) {
    children.push(h2(`Tabla: ${tn}`));
    for (const m of byTable[tn]) {
      children.push(new Paragraph({ spacing: { before: 100, after: 20 }, children: [new TextRun({ text: `[${m.name}]`, bold: true, size: 22, color: ACCENT })] }));
      const desc = mDescByName[m.name];
      if (m.formatString) children.push(label("Formato", m.formatString));
      if (desc?.proposito) children.push(label("Propósito", desc.proposito));
      if (m.expression) children.push(codeBlock(m.expression));
    }
  }

  // 9. RELACIONES (completo desde parsed)
  children.push(h1("7. Relaciones del Modelo"));
  if (secById.relaciones?.contenido) children.push(body(secById.relaciones.contenido));
  if (parsed.relationships.length) {
    children.push(dataTable(
      ["Desde", "Hacia", "Cardinalidad", "Estado"],
      parsed.relationships.map((r) => [r.from, r.to, r.cardinality, { text: r.active ? "Activa" : "Inactiva", align: AlignmentType.CENTER }]),
      [30, 30, 22, 18],
    ));
  } else {
    children.push(body("El modelo no contiene relaciones.", { italics: true, color: GREY }));
  }

  // 10. PÁGINAS DEL REPORTE (completo desde parsed)
  children.push(h1("8. Páginas del Reporte"));
  if (parsed.pages.length) {
    for (const pg of parsed.pages) {
      children.push(h2(pg.name));
      const d = pDescByName[pg.name];
      if (d?.proposito) children.push(body(d.proposito));
      const visuals = (pg.visuals || []).map((v) => v.title ? `${v.type} ("${v.title}")` : v.type);
      if (visuals.length) children.push(label("Visuales", visuals.join(", ")));
    }
  } else {
    children.push(body("No se detectaron páginas en el reporte.", { italics: true, color: GREY }));
  }

  // 11. CONCLUSIONES
  children.push(h1("9. Conclusiones"));
  children.push(body(docData.conclusiones || "El modelo documentado provee una base estructurada para el análisis. Se recomienda mantener este documento actualizado ante cambios en tablas, medidas o relaciones."));

  return new Document({
    creator: "PBIX/PBIT Documenter",
    title: docData.titulo || "Documentación Técnica",
    styles: {
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 30, bold: true, color: NAVY }, paragraph: { spacing: { before: 300, after: 120 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "D5D9E0", space: 4 } } } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 24, bold: true, color: ACCENT }, paragraph: { spacing: { before: 200, after: 60 } } },
      ],
    },
    sections: [{
      properties: { page: { margin: { top: 1100, bottom: 1100, left: 1100, right: 1100 } } },
      children,
    }],
  });
}

export async function buildDocxBlob(docData, parsed) {
  return Packer.toBlob(buildDocxDocument(docData, parsed));
}
