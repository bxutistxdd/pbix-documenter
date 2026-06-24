import {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, TableOfContents, PageBreak,
} from "docx";
import { getColumnDescription, getTableContext } from "./prompts.js";

const FONT = "Calibri";
const PRIMARY = "1F4E79";   // azul oscuro profesional
const ACCENT = "2E74B5";    // azul medio (acento)
const GREY = "5A5A5A";
const HEADER_BG = "1F4E79";
const ZEBRA_BG = "EAF1F8";
const BORDER = "BDD2E8";
const CODE_BG = "EEF3F9";
const CODE_TX = "1F3864";

// Fuentes de datos (contexto Seguros del Estado).
const FUENTES = [
  { nombre: "Analytics OData - WorkItems", tipo: "OData / Azure DevOps Analytics", url: "https://analytics.dev.azure.com/{organization}/{project}/_odata/v3.0/WorkItems", funcion: "Provee work items del proyecto segurosdelestado: épicas, historias, bugs e incidentes para métricas de implementación." },
  { nombre: "Analytics OData - TestPlans", tipo: "OData / Azure DevOps Analytics", url: "https://analytics.dev.azure.com/{organization}/{project}/_odata/v3.0/TestPlans", funcion: "Provee planes de prueba, casos de prueba y resultados de ejecución desde Azure DevOps." },
  { nombre: "SharePoint Lista 1", tipo: "SharePoint List", url: "[PENDIENTE]", funcion: "Lista de SharePoint que contiene información de proyectos y entregas." },
  { nombre: "SharePoint Lista 2", tipo: "SharePoint List", url: "[PENDIENTE]", funcion: "Lista de SharePoint que contiene información de incidentes y bugs." },
];

const ARQUITECTURA = [
  ["Origen", "Azure DevOps OData (Analytics)", "Dos queries al proyecto segurosdelestado: WorkItems y TestPlans."],
  ["Origen", "SharePoint Lists", "Listas con datos de entrada, parámetros y fechas hábiles."],
  ["Origen", "VSTS", "Conexión a Visual Studio Team Services."],
  ["Staging / Jerarquía", "HierarchyFlat / HierarchyTest", "Tablas derivadas de OData que representan la jerarquía Epic→Bug y testing."],
  ["Modelo semántico", "Tablas tbl_ / Dim_ / SP_", "Hechos, dimensiones y tablas SharePoint del modelo."],
  ["Presentación", "Power BI Service", "Reporte con páginas publicadas y PowerApps embebida."],
];

// ---- helpers -------------------------------------------------------------
const body = (text, opts = {}) => new Paragraph({ spacing: { after: 120, line: 276 }, children: [new TextRun({ text: text ?? "", size: 21, color: "222222", ...opts })] });
const h2 = (text) => new Paragraph({ text, heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 70 } });
const h3 = (text) => new Paragraph({ text, heading: HeadingLevel.HEADING_3, spacing: { before: 140, after: 40 } });

function label(lbl, val) {
  return new Paragraph({ spacing: { after: 30 }, children: [
    new TextRun({ text: `${lbl}: `, bold: true, size: 21, color: PRIMARY }),
    new TextRun({ text: val ?? "—", size: 21, color: "222222" }),
  ] });
}

function codeBlock(code) {
  const lines = String(code || "").replace(/^\n+/, "").split("\n");
  return new Paragraph({
    spacing: { before: 40, after: 140 }, shading: { fill: CODE_BG },
    border: { left: { style: BorderStyle.SINGLE, size: 18, color: ACCENT, space: 6 } },
    children: lines.flatMap((ln, i) => {
      const run = new TextRun({ text: ln, font: FONT, size: 18, color: CODE_TX });
      return i === 0 ? [run] : [new TextRun({ break: 1 }), run];
    }),
  });
}

function cell(text, { bold = false, color = "222222", bg, align = AlignmentType.LEFT, font } = {}) {
  const lines = String(text ?? "").split("\n");
  return new TableCell({
    shading: bg ? { fill: bg } : undefined,
    margins: { top: 50, bottom: 50, left: 100, right: 100 },
    children: [new Paragraph({ alignment: align, children: lines.flatMap((ln, i) => {
      const run = new TextRun({ text: ln, bold, size: 18, color, font });
      return i === 0 ? [run] : [new TextRun({ break: 1 }), run];
    }) })],
  });
}

function dataTable(headers, rows, widths) {
  const b = { style: BorderStyle.SINGLE, size: 4, color: BORDER };
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: { top: b, bottom: b, left: b, right: b, insideHorizontal: b, insideVertical: b },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map(hh => cell(hh, { bold: true, color: "FFFFFF", bg: HEADER_BG })) }),
      ...rows.map((r, ri) => new TableRow({ children: r.map(c => {
        const o = typeof c === "object" && c !== null ? c : { text: c };
        return cell(o.text, { bg: ri % 2 ? ZEBRA_BG : undefined, align: o.align, font: o.font });
      }) })),
    ],
  });
}

// ---- heurísticas ---------------------------------------------------------
function tipoOrigen(name) {
  if (/^tbl_/i.test(name)) return { tipo: "Hecho", origen: "OData Azure DevOps" };
  if (/^Dim_/i.test(name)) return { tipo: "Dimensión", origen: "OData Azure DevOps" };
  if (/peso/i.test(name)) return { tipo: "Parámetro", origen: "Combinada" };
  if (/borrador/i.test(name)) return { tipo: "Staging", origen: "Combinada" };
  if (/Hierarchy|Jerarqu/i.test(name)) return { tipo: "Dimensión", origen: "OData Azure DevOps" };
  if (/_Medidas|Consolidado/i.test(name)) return { tipo: "Hecho", origen: "Calculada" };
  if (/^SP_|SharePoint/i.test(name)) return { tipo: "Staging", origen: "SharePoint" };
  return { tipo: "—", origen: "—" };
}
const yn = (v) => (v ? "Sí" : "No");

// ---- documento -----------------------------------------------------------
export function buildDocxDocument(narr, parsed) {
  const n = narr || {};
  const tMap = Object.fromEntries((n.tablas || []).map(i => [i.tabla, i]));
  const mMap = Object.fromEntries((n.medidas || []).map(i => [i.medida, i]));
  const pMap = Object.fromEntries((n.paginas || []).map(i => [i.pagina, i]));
  const rMap = Object.fromEntries((n.relaciones || []).map(i => [`${i.desde}__${i.hacia}`, i]));

  const realRels = parsed.relationships.filter(r => !/LocalDateTable|DateTableTemplate/i.test(`${r.fromTable}${r.toTable}`));
  const fecha = new Date().toLocaleDateString("es-CO", { year: "numeric", month: "long", day: "numeric" });
  const titulo = n.titulo || `Documentación Técnica: ${(parsed.fileName || "").replace(/\.(pbix|pbit)$/i, "")}`;
  const ch = [];

  // PORTADA
  ch.push(
    new Paragraph({ spacing: { before: 2200 }, children: [] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "DOCUMENTACIÓN TÉCNICA", bold: true, size: 50, color: PRIMARY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: parsed.fileName || "", size: 26, color: ACCENT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 220 }, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ACCENT, space: 8 } }, children: [new TextRun({ text: "Power BI — Modelo de Datos", size: 22, color: GREY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 30 }, children: [new TextRun({ text: `Fecha: ${fecha}`, size: 20, color: GREY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `Tablas: ${parsed.tables.length}  ·  Medidas: ${parsed.measures.length}  ·  Relaciones: ${realRels.length}  ·  Páginas: ${parsed.pages.length}`, bold: true, size: 20, color: PRIMARY })] }),
    new Paragraph({ children: [new PageBreak()] }),
  );

  // TABLA DE CONTENIDOS
  ch.push(
    new Paragraph({ text: "Tabla de Contenidos", heading: HeadingLevel.HEADING_1 }),
    new TableOfContents("Contenido", { hyperlink: true, headingStyleRange: "1-2" }),
    new Paragraph({ children: [new PageBreak()] }),
  );

  // 1. INTRODUCCIÓN
  ch.push(new Paragraph({ text: "1. Introducción", heading: HeadingLevel.HEADING_1 }));
  ch.push(body(n.introduccion || "La organización utiliza un reporte Power BI para medir métricas de implementación y testing de proyectos software."));

  // 2. RESUMEN EJECUTIVO
  ch.push(new Paragraph({ text: "2. Resumen Ejecutivo", heading: HeadingLevel.HEADING_1 }));
  ch.push(body(n.resumen_ejecutivo || ""));

  // 3. FUENTES DE DATOS
  ch.push(new Paragraph({ text: "3. Fuentes de Datos", heading: HeadingLevel.HEADING_1 }));
  ch.push(body("A continuación se detallan los orígenes de datos conectados al modelo. Los campos marcados como [PENDIENTE] deben ser completados con la URL real del entorno."));
  for (const f of FUENTES) {
    ch.push(h2(f.nombre));
    ch.push(label("Tipo de origen", f.tipo));
    ch.push(label("Endpoint / URL de acceso", f.url));
    ch.push(label("Función en el modelo", f.funcion));
  }

  // 4. ARQUITECTURA DEL MODELO
  ch.push(new Paragraph({ text: "4. Arquitectura del Modelo", heading: HeadingLevel.HEADING_1 }));
  ch.push(body(n.arquitectura || "El modelo de datos sigue un patrón de capas, con tablas de staging, dimensiones y hechos."));
  ch.push(h2("4.1 Diagrama de Flujo de Datos"));
  ch.push(dataTable(["Capa", "Componente", "Descripción"], ARQUITECTURA, [22, 30, 48]));

  // 5. TABLAS Y COLUMNAS
  ch.push(new Paragraph({ text: "5. Tablas y Columnas", heading: HeadingLevel.HEADING_1 }));
  ch.push(body(`El modelo contiene ${parsed.tables.length} tablas en total.`));
  for (const t of parsed.tables) {
    const d = tMap[t.name] || {};
    const heur = tipoOrigen(t.name);
    ch.push(h2(t.name));
    ch.push(label("Tipo", d.tipo || heur.tipo));
    ch.push(label("Origen", d.origen || heur.origen));
    const tdesc = getTableContext(t.name, d.descripcion);
    if (tdesc) ch.push(body(tdesc));
    const rows = t.columns.map(c => {
      const cdesc = getColumnDescription(t.name, c.name, undefined) || c.name.replace(/([a-z])([A-Z])/g, "$1 $2");
      return [c.name, c.dataType, { text: yn(c.calculated), align: AlignmentType.CENTER }, { text: yn(c.isHidden), align: AlignmentType.CENTER }, cdesc];
    });
    if (rows.length) ch.push(dataTable(["Columna", "Tipo de Dato", "Calculada", "Oculta", "Descripción"], rows, [26, 16, 12, 12, 34]));
    else ch.push(body("Sin columnas visibles.", { italics: true, color: GREY }));
  }

  // 6. MEDIDAS DAX
  ch.push(new Paragraph({ text: "6. Medidas DAX", heading: HeadingLevel.HEADING_1 }));
  ch.push(body(`El modelo contiene ${parsed.measures.length} medidas DAX.`));
  const byTable = {};
  for (const m of parsed.measures) (byTable[m.table] ||= []).push(m);
  for (const tn of Object.keys(byTable)) {
    ch.push(h2(`Tabla: ${tn}`));
    for (const m of byTable[tn]) {
      ch.push(h3(`[${m.name}]`));
      const d = mMap[m.name];
      ch.push(label("Propósito", d?.proposito || "Medida del modelo; revisar la expresión DAX para el detalle de cálculo."));
      if (m.formatString) ch.push(label("Formato", m.formatString));
      if (m.expression) ch.push(codeBlock(m.expression));
    }
  }

  // 7. RELACIONES DEL MODELO
  ch.push(new Paragraph({ text: "7. Relaciones del Modelo", heading: HeadingLevel.HEADING_1 }));
  ch.push(body(`El modelo define ${realRels.length} relaciones entre tablas (se excluyen las tablas de fecha automáticas).`));
  if (realRels.length) {
    const rows = realRels.map(r => {
      const motivo = rMap[`${r.from}__${r.to}`]?.motivo || `Relaciona ${r.fromTable} con ${r.toTable} a través de ${r.fromColumn} / ${r.toColumn}.`;
      return [`${r.from}\n→ ${r.to}`, "Uno a muchos", { text: r.active ? "✓ Activa" : "✗ Inactiva", align: AlignmentType.CENTER }, motivo];
    });
    ch.push(dataTable(["Relación", "Cardinalidad", "Estado", "Motivo y uso"], rows, [34, 16, 14, 36]));
  } else {
    ch.push(body("El modelo no contiene relaciones reales (fuera de tablas de fecha automáticas).", { italics: true, color: GREY }));
  }

  // 8. PÁGINAS DEL REPORTE
  ch.push(new Paragraph({ text: "8. Páginas del Reporte", heading: HeadingLevel.HEADING_1 }));
  ch.push(body(`El reporte contiene ${parsed.pages.length} páginas publicadas en Power BI Service.`));
  for (const pg of parsed.pages) {
    ch.push(h2(pg.name));
    const d = pMap[pg.name];
    ch.push(body(d?.descripcion || `Página "${pg.name}" del reporte.`));
    const tipos = [...new Set((pg.visuals || []).map(v => v.type))];
    ch.push(label("Tipos de visuales", tipos.join(", ") || "—"));
    ch.push(label("Total de visuales", String((pg.visuals || []).length)));
  }

  // 9. CONCLUSIONES Y RECOMENDACIONES
  ch.push(new Paragraph({ text: "9. Conclusiones y Recomendaciones", heading: HeadingLevel.HEADING_1 }));
  ch.push(body(n.conclusiones || "El reporte proporciona una visión integral del desempeño de los proyectos. Se recomienda mantener la documentación actualizada ante cambios en el modelo."));

  return new Document({
    creator: "PBIX/PBIT Documenter", title: titulo,
    styles: {
      default: { document: { run: { font: FONT, size: 21, color: "1A1A1A" } } },
      paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 30, bold: true, color: PRIMARY }, paragraph: { spacing: { before: 320, after: 140 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BORDER, space: 4 } } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 24, bold: true, color: ACCENT }, paragraph: { spacing: { before: 200, after: 60 } } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 21, bold: true, color: PRIMARY }, paragraph: { spacing: { before: 140, after: 40 } } },
    ] },
    sections: [{ properties: { page: { margin: { top: 1100, bottom: 1100, left: 1100, right: 1100 } } }, children: ch }],
  });
}

export async function buildDocxBlob(narr, parsed) {
  return Packer.toBlob(buildDocxDocument(narr, parsed));
}
