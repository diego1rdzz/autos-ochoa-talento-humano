// Genera la versión en Word (.docx) del informe, con el mismo contenido que informe.html.
// Las tres gráficas (clima, reporte 360° y calendario) se toman del HTML con Playwright.
// Uso: npm run build:docx
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { chromium } from 'playwright';
import {
  AlignmentType, BorderStyle, Document, Footer, Header, HeadingLevel, HeightRule, ImageRun,
  LevelFormat, LineRuleType, Packer, PageNumber, Paragraph, SectionType, ShadingType, Tab, Table, TableCell,
  TableLayoutType, TableRow, TabStopType, TextRun, VerticalAlign, WidthType,
} from 'docx';

const here = path.dirname(fileURLToPath(import.meta.url));
const htmlPath = path.join(here, 'informe.html');
const outPath = path.resolve(here, '..', 'entregables', 'Autos_Ochoa_Reto_Final_Desarrollo_Organizacional.docx');

// ---------- Estilo ----------
const C = {
  ink: '0F1B2D', ink2: '3A4556', muted: '667085', line: 'D9DDE4', line2: 'E9ECF1', soft: 'F4F6F9',
  navy: '15325C', deep: '0D2240', accent: 'C25A14', accentSoft: 'FCF0E5', accentLine: 'EFD3BB',
  oe1: '2A78D6', oe2: 'EB6834', oe3: '1BAF7A', white: 'FFFFFF',
};
const SANS = 'Calibri';
const SERIF = 'Georgia';
const SZ = { body: 19, small: 17, table: 17, th: 14, label: 13 };
const PAGE = { w: 12240, h: 15840, mx: 864, top: 900, bottom: 820 };
const CW = PAGE.w - 2 * PAGE.mx; // ancho útil en DXA
const GAP = 180;
const PX = 15; // 1 px CSS = 15 DXA

// ---------- Texto ----------
// **negritas**, __cursivas__ y {OE-1} para las etiquetas de objetivo.
function runs(str, o = {}) {
  const out = [];
  for (const part of String(str).split(/(\*\*[^*]+\*\*|__[^_]+__|\{OE-[123]\})/)) {
    if (!part) continue;
    let m;
    if ((m = part.match(/^\*\*(.+)\*\*$/))) out.push(new TextRun({ ...o, text: m[1], bold: true }));
    else if ((m = part.match(/^__(.+)__$/))) out.push(new TextRun({ ...o, text: m[1], italics: true }));
    else if ((m = part.match(/^\{OE-([123])\}$/))) {
      out.push(new TextRun({ ...o, text: '● ', color: C['oe' + m[1]] }));
      out.push(new TextRun({ ...o, text: `OE-${m[1]}   `, bold: true }));
    } else out.push(new TextRun({ ...o, text: part }));
  }
  return out;
}

function p(str, o = {}) {
  const { size = SZ.body, color = C.ink, bold, after = 80, before = 0, align, keepNext, line = 252, indent, font, caps, spacing } = o;
  return new Paragraph({
    alignment: align, keepNext, indent, spacing: { before, after, line, lineRule: LineRuleType.AUTO },
    children: runs(str, { size, color, bold, font, allCaps: caps, characterSpacing: spacing }),
  });
}
const small = (str, o = {}) => p(str, { size: SZ.small, color: C.ink2, ...o });
const label = (text, color = C.accent, o = {}) => new Paragraph({
  spacing: { before: o.before ?? 0, after: o.after ?? 20 }, keepNext: true,
  children: [new TextRun({ text, allCaps: true, bold: true, size: o.size ?? SZ.label, color, characterSpacing: 16 })],
});
const tiny = () => new Paragraph({ spacing: { before: 0, after: 0, lineRule: LineRuleType.AUTO, line: 200 }, run: { size: 2 }, children: [] });
const gap = () => new Paragraph({ spacing: { before: 0, after: 0, lineRule: LineRuleType.EXACT, line: 110 }, run: { size: 2 }, children: [] });

const h1 = (text, o = {}) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 0, after: o.after ?? 120 }, children: [new TextRun({ text })] });
const h2 = (text, o = {}) => new Paragraph({ heading: HeadingLevel.HEADING_2, pageBreakBefore: o.pageBreak, spacing: { before: o.before ?? 140, after: 60 }, children: [new TextRun({ text })] });
const h3 = (text, o = {}) => new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: o.before ?? 120, after: 50 }, children: [new TextRun({ text })] });

function moduleHead(num, eyebrow, title) {
  return [
    label(eyebrow, C.accent, { size: 15, after: 30 }),
    new Paragraph({
      heading: HeadingLevel.HEADING_1, spacing: { before: 0, after: 120 },
      children: [new TextRun({ text: `${num}  `, color: C.accent }), new TextRun({ text: title })],
    }),
  ];
}

const bullets = (items, o = {}) => items.map(it => new Paragraph({
  numbering: { reference: 'bul', level: 0 }, spacing: { after: o.after ?? 30, lineRule: LineRuleType.AUTO, line: 245 },
  children: runs(it, { size: o.size ?? SZ.body, color: o.color ?? C.ink }),
}));

// ---------- Tablas ----------
const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const ln = (color = C.line, size = 4) => ({ style: BorderStyle.SINGLE, size, color });
const noCell = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const box = (color = C.line, size = 4) => ({ top: ln(color, size), bottom: ln(color, size), left: ln(color, size), right: ln(color, size) });
const shade = fill => (fill ? { type: ShadingType.CLEAR, fill, color: 'auto' } : undefined);

function colWidths(weights, W) {
  const tot = weights.reduce((a, b) => a + b, 0);
  const ws = weights.map(w => Math.floor((w / tot) * W));
  ws[ws.length - 1] += W - ws.reduce((a, b) => a + b, 0);
  return ws;
}
const splitW = (weights, gap = GAP, W = CW) => colWidths(weights, W - gap * (weights.length - 1));

function tbl(widths, rows) {
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths, layout: TableLayoutType.FIXED,
    borders: { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE },
    rows,
  });
}
// Una celda debe terminar en párrafo; si termina en tabla se agrega uno mínimo.
const cellKids = kids => {
  const arr = (Array.isArray(kids) ? kids : [kids]).flat();
  if (!arr.length || arr[arr.length - 1] instanceof Table) arr.push(tiny());
  return arr;
};
const toParas = (c, o) => (Array.isArray(c) ? c : String(c).split('\n').map(s => p(s, { after: 0, line: 240, ...o })));

function cell(w, kids, o = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, children: cellKids(kids),
    borders: o.borders ?? noCell, shading: shade(o.fill), columnSpan: o.span, rowSpan: o.rowSpan,
    verticalAlign: o.vAlign ?? VerticalAlign.TOP,
    margins: o.margins ?? { top: 40, bottom: 40, left: 80, right: 80 },
  });
}
const spacerCell = w => cell(w, [tiny()], { margins: { top: 0, bottom: 0, left: 0, right: 0 } });

// Tabla de datos con encabezado. rows: arreglo de celdas; cada celda es texto, arreglo de párrafos
// o { c, span, rowSpan, fill, align, bold, color, size }.
function dataTable(weights, header, rows, o = {}) {
  const W = o.width ?? CW;
  const widths = colWidths(weights, W);
  const size = o.size ?? SZ.table;
  const pad = o.pad ?? 36;
  const out = [];
  if (header) {
    out.push(new TableRow({
      tableHeader: true, cantSplit: true,
      children: header.map((h, i) => {
        const spec = typeof h === 'object' ? h : { c: h };
        return cell(widths[i], [new Paragraph({
          alignment: spec.align, spacing: { after: 0, lineRule: LineRuleType.AUTO, line: 230 },
          children: [new TextRun({ text: spec.c, allCaps: true, bold: true, size: SZ.th, color: C.muted, characterSpacing: 10 })],
        })], { borders: { ...noCell, bottom: ln(C.ink, 10) }, vAlign: VerticalAlign.BOTTOM, margins: { top: 30, bottom: 30, left: 70, right: 70 } });
      }),
    }));
  }
  const occupied = new Array(widths.length).fill(0);
  rows.forEach((r, ri) => {
    let col = 0;
    const cells = [];
    for (const raw of r) {
      while (occupied[col] > 0) col++;
      const spec = raw && typeof raw === 'object' && !Array.isArray(raw) ? raw : { c: raw };
      const span = spec.span ?? 1;
      const w = widths.slice(col, col + span).reduce((a, b) => a + b, 0);
      const last = ri === rows.length - 1 || (spec.rowSpan && ri + spec.rowSpan >= rows.length);
      cells.push(cell(w, toParas(spec.c, { size: spec.size ?? size, color: spec.color ?? C.ink, bold: spec.bold, align: spec.align }), {
        span: spec.span, rowSpan: spec.rowSpan, fill: spec.fill,
        borders: spec.borders ?? { ...noCell, bottom: last ? ln(C.line, 6) : ln(C.line2, 4) },
        margins: spec.margins ?? { top: pad, bottom: pad, left: 70, right: 70 },
      }));
      if (spec.rowSpan > 1) for (let k = col; k < col + span; k++) occupied[k] = spec.rowSpan;
      col += span;
    }
    for (let k = 0; k < occupied.length; k++) if (occupied[k] > 0) occupied[k]--;
    out.push(new TableRow({ cantSplit: true, children: cells }));
  });
  return tbl(widths, out);
}

// Tarjetas con borde lado a lado (una sola fila, sin tablas anidadas). cols: [{ kids, fill, borders, w }]
// split: true deja que la fila se parta entre páginas en lugar de recortarse.
function cards(cols, o = {}) {
  const gap = o.gap ?? 120;
  const ws = o.widths ?? splitW(cols.map(c => c.w ?? 1), gap, o.width ?? CW);
  const widths = [], cells = [];
  cols.forEach((c, i) => {
    if (i > 0) { widths.push(gap); cells.push(spacerCell(gap)); }
    widths.push(ws[i]);
    cells.push(cell(ws[i], c.kids, {
      fill: c.fill ?? o.fill, borders: c.borders ?? o.borders ?? box(), vAlign: o.vAlign,
      margins: o.margins ?? { top: 80, bottom: 80, left: 110, right: 110 },
    }));
  });
  return tbl(widths, [new TableRow({ cantSplit: !o.split, children: cells })]);
}

function ficha(items) {
  const ws = colWidths([1.05, 0.95, 1.05, 1.35], CW);
  return tbl(ws, [new TableRow({
    cantSplit: true,
    children: items.map(([k, v], i) => cell(ws[i], [
      label(k, C.muted, { size: 12, after: 10 }),
      p(v, { size: SZ.small, bold: true, after: 0, line: 235 }),
    ], {
      fill: C.soft,
      borders: { top: ln(), bottom: ln(), left: i === 0 ? ln() : ln(C.line), right: ln() },
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
    })),
  })]);
}

function why(a, b) {
  const ws = colWidths([1, 1], CW);
  const blk = (k, v) => [label(k), p(v, { size: SZ.small, after: 0, line: 240 })];
  const kids = b === undefined ? [blk('Por qué lo hacemos y qué gana la empresa', a)] : [blk('Por qué lo hacemos', a), blk('Qué gana la empresa', b)];
  const widths = b === undefined ? [CW] : ws;
  return tbl(widths, [new TableRow({
    cantSplit: true,
    children: kids.map((k, i) => cell(widths[i], k, { fill: C.accentSoft, borders: box(C.accentLine), margins: { top: 70, bottom: 70, left: 120, right: 120 } })),
  })]);
}

function callout(title, text, o = {}) {
  const W = o.width ?? CW;
  return tbl([W], [new TableRow({
    cantSplit: true,
    children: [cell(W, [label(title), ...(Array.isArray(text) ? text : [p(text, { size: o.size ?? SZ.body, after: 0, line: 245 })])], {
      fill: C.accentSoft, borders: { ...noCell, left: ln(C.accent, 18) },
      margins: { top: 70, bottom: 70, left: 140, right: 120 },
    })],
  })]);
}

const img = (chart, maxW) => {
  const scale = Math.min(1, maxW / chart.w);
  return new Paragraph({
    spacing: { before: 20, after: 40, line: 240, lineRule: LineRuleType.AUTO },
    children: [new ImageRun({ type: 'png', data: chart.buf, transformation: { width: Math.round(chart.w * scale), height: Math.round(chart.h * scale) } })],
  });
};

// ---------- Gráficas tomadas del HTML ----------
async function captureCharts() {
  const browser = await chromium.launch();
  const page = await browser.newPage({ deviceScaleFactor: 3, viewport: { width: 1100, height: 1400 } });
  await page.goto(pathToFileURL(htmlPath).href, { waitUntil: 'load' });
  await page.evaluate(() => document.fonts.ready);
  await page.evaluate(() => {
    const wrap = (nodes, id) => {
      const w = document.createElement('div');
      w.id = id;
      w.style.cssText = 'background:#fff;padding:3px 2px';
      nodes[0].parentNode.insertBefore(w, nodes[0]);
      nodes.forEach(n => w.appendChild(n));
    };
    const bars = document.querySelector('.bars');
    wrap([bars, bars.nextElementSibling], 'cap-bars');
    const dots = document.querySelector('.dots').parentElement;
    dots.id = 'cap-dots';
    dots.style.background = '#fff';
    const gantt = document.querySelector('.gantt');
    wrap([gantt], 'cap-gantt');
  });
  const shot = async id => {
    const el = page.locator('#' + id);
    const bb = await el.boundingBox();
    return { buf: await el.screenshot(), w: bb.width, h: bb.height };
  };
  const charts = { bars: await shot('cap-bars'), dots: await shot('cap-dots'), gantt: await shot('cap-gantt') };
  await browser.close();
  return charts;
}

// ---------- Encabezados y pies ----------
const tabRight = [{ type: TabStopType.RIGHT, position: CW }];
function header(section) {
  return new Header({
    children: [new Paragraph({
      tabStops: tabRight, spacing: { after: 0 }, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.line, space: 4 } },
      children: [
        new TextRun({ text: 'Autos Ochoa · Plan de talento 2026–2027', allCaps: true, size: 13, color: C.muted, characterSpacing: 14 }),
        new TextRun({ children: [new Tab(), section], allCaps: true, bold: true, size: 13, color: C.accent, characterSpacing: 14 }),
      ],
    })],
  });
}
const footer = new Footer({
  children: [new Paragraph({
    tabStops: tabRight, spacing: { after: 0 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.line, space: 4 } },
    children: [
      new TextRun({ text: 'Reto Final · Administración del Talento Humano · Septiembre 2026', size: 13, color: C.muted }),
      new TextRun({ children: [new Tab(), PageNumber.CURRENT], bold: true, size: 14, color: C.ink }),
    ],
  })],
});
const pageProps = {
  page: {
    size: { width: PAGE.w, height: PAGE.h },
    margin: { top: PAGE.top, bottom: PAGE.bottom, left: PAGE.mx, right: PAGE.mx, header: 460, footer: 420 },
  },
};
const section = (name, children) => ({
  properties: { ...pageProps, type: SectionType.NEXT_PAGE },
  headers: { default: header(name) }, footers: { default: footer }, children,
});

// ---------- Portada ----------
// Sin tablas anidadas ni altura fija: una franja azul, los datos clave y la lista de módulos.
function cover() {
  const W = CW;
  const onDark = (text, o = {}) => p(text, { color: o.color ?? 'D4DCEA', size: o.size ?? 20, after: o.after ?? 0, before: o.before ?? 0, bold: o.bold, font: o.font, line: o.line ?? 260, caps: o.caps, spacing: o.spacing });
  const pad = 520;
  const band = tbl([W], [new TableRow({
    children: [cell(W, [
      new Paragraph({
        tabStops: [{ type: TabStopType.RIGHT, position: W - 2 * pad }], spacing: { after: 0 },
        children: [
          new TextRun({ text: 'Autos Ochoa — Grupo Ochoa', allCaps: true, size: 14, color: 'AEBBD0', characterSpacing: 30, bold: true }),
          new TextRun({ children: [new Tab(), 'Monterrey, Nuevo León'], allCaps: true, size: 14, color: 'AEBBD0', characterSpacing: 30, bold: true }),
        ],
      }),
      onDark('Plan de talento · Parte 2', { before: 1500, after: 160, size: 17, bold: true, color: 'F3B27F', caps: true, spacing: 28 }),
      new Paragraph({ spacing: { after: 0, lineRule: LineRuleType.AUTO, line: 240 }, children: [new TextRun({ text: 'Cómo vamos a formar, cuidar y hacer crecer a nuestra gente', font: SERIF, size: 64, color: C.white })] }),
      onDark('Ocho acciones de Recursos Humanos para abrir cuatro sucursales nuevas sin perder lo que nos hace Autos Ochoa. Septiembre 2026 – agosto 2027.', { before: 300, size: 21, line: 300 }),
    ], { fill: C.deep, margins: { top: pad, bottom: 620, left: pad, right: pad } })],
  })]);
  const sw = colWidths([1, 1, 1, 1], W);
  const stats = [
    ['1 → 5', 'sucursales en 12 meses', 'Guadalupe · San Nicolás · Apodaca · Saltillo'],
    ['8 → 34', 'personas en el equipo', '26 puestos nuevos'],
    ['5 → 60', 'autos vendidos al mes', 'entre todas las sucursales'],
    ['4.5 → 49', 'autos al mes para no perder', '(punto de equilibrio)'],
  ];
  const statsTbl = tbl(sw, [new TableRow({
    children: stats.map(([v, a, b2], i) => cell(sw[i], [
      onDark(v, { size: 36, bold: true, color: C.white, after: 40 }),
      onDark(a, { size: 15, color: 'C3CFE2', line: 230 }),
      onDark(b2, { size: 15, color: 'C3CFE2', line: 230 }),
    ], {
      fill: C.navy, borders: { ...noCell, left: i ? ln('33496B', 6) : NONE },
      margins: { top: 200, bottom: 200, left: i ? 200 : pad, right: 100 },
    })),
  })]);
  const mods = [
    ['1', 'Capacitación y Desarrollo', 'Academia Ochoa: formación presencial y en línea'],
    ['5', 'Responsabilidad Social', 'Tres políticas de inclusión y equidad'],
    ['2', 'Clima Laboral y Experiencia', 'Encuesta «Pulso Ochoa» y plan de 6 meses'],
    ['6', 'Sucesión y Talento Clave', 'Mapa de talento de 9 cuadros (9-Box)'],
    ['3', 'Marca Empleadora y Permanencia', 'Lo que ofrecemos y cómo cuidamos al equipo'],
    ['7', 'Digitalización y Agilidad', 'Herramientas digitales y ciclos de 2 semanas'],
    ['4', 'Desempeño y Competencias', 'Seis competencias y evaluación 360°'],
    ['8', 'Innovación y Excelencia', 'Las prácticas del fundador en el CRM y certificación'],
  ];
  const mw = colWidths([1, 1], W - 400);
  const modCell = ([n, t, d], w) => cell(w, [new Paragraph({
    spacing: { after: 0, lineRule: LineRuleType.AUTO, line: 240 },
    children: [
      new TextRun({ text: `${n}   `, font: SERIF, bold: true, size: 24, color: C.accent }),
      new TextRun({ text: t, size: 19, bold: true, color: C.ink }),
    ],
  }), p(d, { size: 16, color: C.muted, after: 0, line: 240, indent: { left: 300 } })], { borders: { ...noCell, bottom: ln(C.line, 4) }, margins: { top: 110, bottom: 110, left: 0, right: 0 } });
  const modRows = [];
  for (let i = 0; i < mods.length; i += 2) modRows.push(new TableRow({ cantSplit: true, children: [modCell(mods[i], mw[0]), spacerCell(400), modCell(mods[i + 1], mw[1])] }));
  return {
    properties: pageProps,
    children: [
      band,
      statsTbl,
      label('Las ocho acciones de este plan', C.accent, { size: 15, before: 520, after: 60 }),
      tbl([mw[0], 400, mw[1]], modRows),
      new Paragraph({
        tabStops: tabRight, spacing: { before: 700, after: 0 },
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: C.line, space: 8 } },
        children: [
          new TextRun({ text: 'Administración del Talento Humano · Reto Final', size: 16, color: C.muted }),
          new TextRun({ children: [new Tab(), 'Septiembre 2026'], size: 16, color: C.muted }),
        ],
      }),
    ],
  };
}

// ---------- Contenido ----------
function resumen() {
  const oe = [
    ['1', 'Gerente listo antes de abrir', 'Contratado y certificado 60 días antes de la apertura, en las 4 sucursales. Vacante cubierta en 45 días o menos.'],
    ['2', 'Sucursal rentable en 4 meses', 'Que cada sucursal venda al menos 10 autos al mes antes de 120 días. Meta: 4 de cada 5 gerentes.'],
    ['3', 'Que el talento se quede', 'Que se vayan menos del 10% de los gerentes al año, siempre cumpliendo la ley laboral.'],
  ];
  const [iw1, iw2] = colWidths([0.9, 3.1], CW);
  return [
    h1('En pocas palabras'),
    p('Autos Ochoa va a pasar de una a cinco sucursales en un año. Para lograrlo necesitamos a la persona correcta al frente de cada una: el **Gerente de Sucursal**. En la Parte 1 definimos cómo contratarlo y cuánto pagarle. Este documento explica **qué vamos a hacer a partir de ahora** para prepararlo, evaluarlo, cuidarlo y tener quién lo reemplace, y cómo eso ayuda a todo el equipo.', { size: 20, color: C.ink2, after: 100 }),
    p('**¿Por qué ahora?** Con cinco sucursales tendremos que vender casi 49 autos al mes solo para no perder dinero; hoy bastan 4.5. Cada mes que una sucursal opera sin un gerente preparado deja de ganar cerca de **$200,000**, y cuando un gerente se va, reemplazarlo cuesta alrededor de **$500,000**. Además, hoy el 97% de los clientes que contactamos no avanza porque nadie es responsable de darles seguimiento. Todo lo que proponemos apunta a tres objetivos, los mismos de la Parte 1:', { after: 100 }),
    cards(oe.map(([n, t, m]) => ({
      kids: [p(`{OE-${n}}`, { size: SZ.small, after: 20 }), p(t, { bold: true, after: 20 }), small(m, { after: 0, line: 240 })],
      borders: { ...box(), top: ln(C['oe' + n], 24) },
    }))),
    h2('Las ocho acciones'),
    dataTable([22, 31, 13, 25, 9], ['Acción', 'Qué problema resuelve', 'Objetivo', 'Cómo sabremos que funciona', { c: 'Costo al año', align: AlignmentType.RIGHT }], [
      ['**1 · Academia Ochoa**', 'Gerentes que llegan a abrir sin conocer nuestra forma de trabajar', '{OE-1}{OE-2}', '100% certificados 60 días antes de abrir', { c: '$60,000', align: AlignmentType.RIGHT }],
      ['**2 · Encuesta y plan de clima**', 'Nadie sabe quién da seguimiento, WhatsApp fuera de horario y reglas que se sienten distintas para la familia', '{OE-3}{OE-2}', 'Clima favorable de 54% a 70% o más en mayo 2027', { c: '$30,000', align: AlignmentType.RIGHT }],
      ['**3 · Marca empleadora**', 'Competimos por gerentes con Trefa y Credimotors, que tienen más presupuesto', '{OE-3}{OE-1}', 'Menos del 10% de los gerentes se va al año', { c: '$40,000', align: AlignmentType.RIGHT }],
      ['**4 · Evaluación del desempeño**', 'Hoy no evaluamos a nadie de forma formal', '{OE-2}', '4 de cada 5 gerentes en su meta a los 120 días', { c: '$10,000', align: AlignmentType.RIGHT }],
      ['**5 · Políticas de inclusión**', 'Giro con mayoría de hombres, comisiones por fuera y trato distinto a familiares', '{OE-3}{OE-1}', 'Diferencia de sueldo en el mismo puesto de 5% o menos', { c: '$20,000', align: AlignmentType.RIGHT }],
      ['**6 · Plan de sucesión**', 'Todo gerente viene de fuera y el conocimiento depende del fundador', '{OE-1}{OE-3}', '6 de cada 10 puestos clave con reemplazo listo', { c: '$12,000', align: AlignmentType.RIGHT }],
      ['**7 · Digital y ágil**', 'Una sola persona de RH para 26 contrataciones y 4 aperturas', '{OE-1}{OE-2}', 'Vacante de gerente cubierta en 45 días o menos', { c: '$0*', align: AlignmentType.RIGHT }],
      ['**8 · Innovación**', '25 años de experiencia concentrados en una sola persona', '{OE-2}{OE-3}', 'Gerente nuevo en su meta en 90 días (antes 120)', { c: '$69,000', align: AlignmentType.RIGHT }],
    ]),
    small('* La nómina con checador electrónico (~$24,500 al año) ya está en el presupuesto de la Parte 1; lo demás usa herramientas que ya tenemos. Montos estimados en pesos.', { before: 50, after: 120 }),
    tbl([iw1, iw2], [new TableRow({
      cantSplit: true,
      children: [
        cell(iw1, [p('$241,000', { size: 40, bold: true, color: C.white, after: 20 }), p('Costo total al año', { size: 14, color: 'B8C6DC', caps: true, spacing: 10, after: 0 })], { fill: C.navy, vAlign: VerticalAlign.CENTER, margins: { top: 110, bottom: 110, left: 200, right: 80 } }),
        cell(iw2, [p('Es **menos de la mitad de lo que cuesta perder a un solo gerente** (~$500,000) y cerca del **1% de la utilidad bruta del año** cuando lleguemos a la meta ($23 millones). Si evitamos una sola salida o una sucursal llega un mes antes a su meta, el plan se paga solo.', { color: 'DFE6F1', after: 0, line: 250 })], { fill: C.navy, vAlign: VerticalAlign.CENTER, margins: { top: 110, bottom: 110, left: 120, right: 200 } }),
      ],
    })]),
    gap(),
    callout('La idea central', 'Lo que construimos en 25 años con 8 personas no se copia solo en cinco sucursales. Hay que enseñarlo, cuidarlo y medirlo. Cada acción tiene un responsable, una fecha y una forma de saber si funcionó.'),
  ];
}

function modulo1() {
  const levels = [
    ['Nivel 1 · Lo básico de Autos Ochoa', 'Primeros 30 días · 24 horas', ['Lo que prometemos al cliente: auto revisado, con garantía y papeles en orden', 'Usar el CRM y responder en menos de 2 horas', 'El crédito con las 5 financieras y el crédito propio', 'Lo que la ley pide al líder: horarios, domingos y registro de asistencia']],
    ['Nivel 2 · Dirigir la sucursal', 'Días 31 a 120 · 32 horas', ['Que ningún cliente se quede sin seguimiento', 'Cerrar la venta entendiendo lo que el cliente necesita', 'Mover los autos que llevan más de 90 días', 'Entender los números: margen por auto y punto de equilibrio']],
    ['Nivel 3 · Liderar con datos', 'Meses 5 a 12 · 24 horas', ['Leer los tableros y decidir con datos', 'Liderar con respeto y cuidar la salud emocional del equipo (NOM-035)', 'Evaluar y preparar a quien te reemplazará', 'Usar la inteligencia artificial con responsabilidad']],
  ];
  const lw = splitW([1, 1, 1], 120);
  const lwAll = [lw[0], 120, lw[1], 120, lw[2]];
  const lvHead = levels.flatMap(([t, s], i) => [
    ...(i ? [spacerCell(120)] : []),
    cell(lw[i], [p(t, { bold: true, color: C.white, size: SZ.small, after: 0 }), p(s, { size: 14, color: 'C3CFE2', after: 0 })], { fill: C.navy, borders: box(C.navy), margins: { top: 60, bottom: 60, left: 110, right: 90 } }),
  ]);
  const lvBody = levels.flatMap(([, , items], i) => [
    ...(i ? [spacerCell(120)] : []),
    cell(lw[i], bullets(items, { size: SZ.small, after: 20 }), { borders: { ...box(), top: NONE }, margins: { top: 60, bottom: 60, left: 60, right: 90 } }),
  ]);
  const sw = colWidths([70, 20, 10], CW);
  const seg = (w, text, fill) => cell(w, [p(text, { size: 14, bold: true, color: C.white, align: AlignmentType.CENTER, after: 0, line: 220 })], { fill, borders: { ...noCell, right: ln(C.white, 12) }, margins: { top: 30, bottom: 30, left: 40, right: 40 }, vAlign: VerticalAlign.CENTER });
  const method = (title, share, items) => ({
    kids: [
      new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: splitW([1, 1], 140)[0] - 220 }], spacing: { after: 40 }, children: [new TextRun({ text: title, allCaps: true, bold: true, size: 15, color: C.navy, characterSpacing: 14 }), new TextRun({ children: [new Tab(), share], size: 15, color: C.muted })] }),
      ...bullets(items, { size: SZ.small, after: 20 }),
    ],
  });
  return [
    ...moduleHead('01', 'Módulo 1 · Capacitación y Desarrollo', 'Academia Ochoa: preparar a cada gerente antes de abrir'),
    ficha([['Qué entregamos', 'Plan de formación mixto: presencial y en línea'], ['Objetivo al que apoya', '{OE-1}{OE-2}'], ['A quién va dirigido', 'Gerentes y asesores que pueden llegar a gerente'], ['Cómo sabremos que funciona', '100% de gerentes certificados 60 días antes de abrir']]),
    h3('Objetivo del programa'),
    p('Que cada gerente conozca y domine la forma de trabajar de Autos Ochoa **antes de abrir su sucursal**, y que en máximo **120 días** su sucursal venda al menos **10 autos al mes** con un cierre de 18% o más. Al mismo tiempo, preparar a 4 asesores senior para que puedan ser los próximos gerentes.', { after: 0 }),
    h3('A quién va dirigido'),
    dataTable([34, 12, 40, 14], ['Quién', 'Personas', 'Qué cursa', { c: 'Horas al año', align: AlignmentType.RIGHT }], [
      ['**Gerentes de Sucursal**', '5', 'Niveles 1, 2 y 3', { c: '80', align: AlignmentType.RIGHT }],
      ['Asesores comerciales', '15', 'Nivel 1', { c: '48', align: AlignmentType.RIGHT }],
      ['Asesores senior con potencial', '4', 'Además, Nivel 3', { c: '+40', align: AlignmentType.RIGHT }],
      ['Coordinadores de expedientes', '5', 'Parte del Nivel 1', { c: '32', align: AlignmentType.RIGHT }],
      ['Padrinos y Dir. Comercial', '2–5', 'Cómo enseñar', { c: '12', align: AlignmentType.RIGHT }],
    ], { pad: 20 }),
    h3('Temas clave · tres niveles'),
    tbl(lwAll, [new TableRow({ cantSplit: true, children: lvHead }), new TableRow({ cantSplit: true, children: lvBody })]),
    h3('Cómo se enseña · aprender haciendo (modelo 70-20-10)'),
    tbl(sw, [new TableRow({ cantSplit: true, children: [seg(sw[0], '70% en el trabajo diario, con clientes y casos reales', C.navy), seg(sw[1], '20% con otros', '2A5A99'), seg(sw[2], '10% cursos', C.accent)] })]),
    gap(),
    cards([
      method('Presencial', '~45% de las horas', ['**Talleres** cada tres meses en la matriz, practicando con conversaciones reales', '**Un padrino** que acompaña 1 hora a la semana durante 90 días', '**Una charla al mes con el fundador**', '**Cursos de crédito** con Afirme, Motorfy y Banregio']),
      method('En línea', '~55% de las horas', ['**Videos cortos** de 10 a 15 minutos en Google Classroom', '**Simulador en el CRM** para practicar con 10 conversaciones reales', '**Una pregunta diaria por WhatsApp** con tabla de posiciones', '**Examen en línea** y constancia oficial DC-3 (STPS) por nivel']),
    ], { gap: 140 }),
    h3('Cómo sabremos que funciona'),
    dataTable([17, 40, 15, 28], ['Pregunta', 'Qué medimos', 'Meta', 'Cada cuándo'], [
      ['¿Les gustó?', 'Calificación que dan al curso', '4.5 de 5 o más', 'Al terminar cada módulo'],
      ['¿Aprendieron?', 'Gerentes certificados 60 días antes de abrir', '100% (5 de 5)', 'En cada apertura'],
      ['¿Lo aplican?', 'Clientes sin seguimiento al final del día · respuesta en menos de 2 horas', '0 · 90% o más', 'Cada semana'],
      ['¿Da resultados?', 'Días para que la sucursal venda 10 autos al mes', '120 o menos', 'Al día 120 de cada gerente'],
      ['¿Vale la pena?', 'Lo que se gana contra lo que cuesta · horas por persona', '3 a 1 · 40 h', 'Cada año (agosto 2027)'],
    ], { pad: 26 }),
    gap(),
    why('Nuestra forma de vender (CRM con IA, cinco financieras, garantía y papeles en regla) no se aprende en ningún otro lado. Si no la enseñamos nosotros, nadie lo va a hacer.', 'Cada mes sin un gerente preparado cuesta ~$200,000. Si cada gerente llega a su meta 30 días antes, en las 4 aperturas protegemos ~$800,000 con $60,000 de inversión.'),
  ];
}

function modulo2(charts) {
  const dims = [
    ['D1 · Claridad de mi rol', 1, ['Sé exactamente qué resultados se esperan de mí cada semana.', 'Está claro quién es responsable de dar seguimiento a cada cliente.', 'Conozco mis indicadores y cómo se miden.']],
    ['D5 · Trato justo', 13, ['Las reglas son iguales para todos, sean o no de la familia.', 'Los ascensos se deciden con criterios claros y conocidos.', 'Aquí nadie recibe un trato distinto por su género, edad u origen.']],
    ['D2 · Mi jefe directo', 4, ['Mi jefe me da retroalimentación útil al menos una vez por semana.', 'Mi jefe reconoce cuando hago bien mi trabajo.', 'Confío en las decisiones que toma mi jefe.']],
    ['D6 · Crecimiento', 16, ['Sé qué necesito para crecer al siguiente puesto.', 'Recibo la capacitación que necesito para hacer bien mi trabajo.', 'Me veo en Autos Ochoa dentro de dos años.']],
    ['D3 · Carga de trabajo y descanso', 7, ['Puedo hacer mi trabajo en mi horario sin quedarme de más seguido.', 'Fuera de mi horario no se espera que conteste mensajes de trabajo.', 'El rol de domingos es justo y se publica con anticipación.']],
    ['D7 · Herramientas', 19, ['El CRM me facilita el trabajo en lugar de complicarlo.', 'Tengo el equipo que necesito (celular, computadora, internet).', 'Encuentro la información que necesito cuando la necesito.']],
    ['D4 · Sueldo y reconocimiento', 10, ['Entiendo cómo se calculan mis comisiones y bonos.', 'Lo que gano es justo comparado con empresas parecidas.', 'Mis comisiones se pagan completas, a tiempo y en mi recibo de nómina.']],
    ['D8 · Orgullo y comunicación', 22, ['Me siento orgulloso(a) de decir que trabajo en Autos Ochoa.', 'Me entero de las decisiones importantes antes que por rumores.', 'Me siento parte del equipo aunque trabaje en otra sucursal.']],
  ];
  const qw = splitW([1, 1], 200);
  const qwAll = [qw[0], 200, qw[1]];
  const qHead = ([t, s], w) => cell(w, [new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: w - 160 }], spacing: { after: 0 },
    children: [new TextRun({ text: t, bold: true, size: SZ.small, color: C.navy }), new TextRun({ children: [new Tab(), `${s}–${s + 2}`], size: 14, color: C.muted })],
  })], { fill: C.soft, borders: { ...box(), bottom: ln() }, margins: { top: 30, bottom: 30, left: 80, right: 80 } });
  const qBody = ([, s, qs], w) => cell(w, qs.map((q, i) => p(`${s + i}.  ${q}`, { size: SZ.small, after: 10, line: 235, indent: { left: 260, hanging: 260 } })), { borders: { ...box(), top: NONE }, margins: { top: 40, bottom: 40, left: 80, right: 80 } });
  const qRows = [];
  for (let i = 0; i < dims.length; i += 2) {
    qRows.push(new TableRow({ cantSplit: true, children: [qHead(dims[i], qw[0]), spacerCell(200), qHead(dims[i + 1], qw[1])] }));
    qRows.push(new TableRow({ cantSplit: true, children: [qBody(dims[i], qw[0]), spacerCell(200), qBody(dims[i + 1], qw[1])] }));
    if (i < dims.length - 2) qRows.push(new TableRow({ height: { value: 80, rule: HeightRule.ATLEAST }, children: [spacerCell(qw[0]), spacerCell(200), spacerCell(qw[1])] }));
  }
  const scw = colWidths([1, 1, 1, 1, 1], CW);
  const scale = [['1', 'Totalmente en desacuerdo'], ['2', 'En desacuerdo'], ['3', 'Ni de acuerdo ni en desacuerdo'], ['4', 'De acuerdo'], ['5', 'Totalmente de acuerdo']];
  const scaleTbl = tbl(scw, [new TableRow({
    cantSplit: true,
    children: scale.map(([n, t], i) => cell(scw[i], [p(n, { bold: true, size: 17, align: AlignmentType.CENTER, after: 0 }), p(t, { size: 14, align: AlignmentType.CENTER, after: 0, line: 220 })], { fill: i >= 3 ? 'E3EEFB' : C.soft, borders: { ...noCell, right: i < 4 ? ln(C.white, 12) : NONE }, margins: { top: 40, bottom: 40, left: 30, right: 30 } })),
  })]);
  const journey = [['Postularse', 'Por WhatsApp, con respuesta en 48 horas y sueldo visible.'], ['Primer día', 'Contrato firmado, alta en el IMSS y carta del fundador.'], ['Primer mes', 'Un padrino y el Nivel 1 de la Academia.'], ['Primera venta', 'Reconocimiento público en el boletín del equipo.'], ['Día 120', 'Evaluación, plan personal y ruta de carrera.'], ['Si se va', 'Plática de salida y puertas abiertas para volver.']];
  const tiles = [['54%', 'del equipo ve el clima de forma favorable'], ['+19', '¿Recomendarías trabajar aquí? (eNPS): 44% sí, 25% no'], ['94%', 'contestó la encuesta (16 de 17 personas)'], ['Medio', 'riesgo NOM-035 en horarios y en la vida familiar']];

  const MONTHS = ['Dic', 'Ene', 'Feb', 'Mar', 'Abr', 'May'];
  const bar = s => [...s].map(ch => ({
    c: [tiny()], fill: ch === 'x' ? C.navy : ch === 'a' ? C.accent : undefined, margins: { top: 0, bottom: 0, left: 0, right: 0 },
    borders: { top: ln(C.white, 36), bottom: ln(C.white, 36), left: ln(C.white, 6), right: ln(C.white, 6) },
  }));
  const plan = [
    [{ c: '**D1 Claridad de rol**\n31% → 65%', rowSpan: 2 }, 'Escribir quién es responsable de cada cliente y de cada paso de la venta', 'Dir. Comercial y RH', ...bar('xx....')],
    ['Tablero personal en el CRM con 3 números de la semana', 'Sistemas', ...bar('.xx...')],
    [{ c: '**D3 Carga y descanso**\n38% → 60%', rowSpan: 2 }, 'Guardias pagadas; fuera de horario el asistente de IA contesta y agenda', 'Dir. General y RH', ...bar('xx....')],
    ['Rol de domingos con un mes de anticipación y jornada de 46 horas sin bajar sueldos (1 de enero)', 'RH y Finanzas', ...bar('xax...')],
    [{ c: '**D5 Trato justo**\n40% → 65%', rowSpan: 2 }, 'Publicar las bandas de sueldo, iguales para todos (Módulo 5)', 'Dirección General', ...bar('a.....')],
    ['Comité de ascensos con alguien que no sea de la familia y buzón anónimo', 'Dir. General y RH', ...bar('.xxxxx')],
    ['**D4 Sueldo**\n44% → 65%', 'Calculadora de comisiones en el CRM, todo en nómina, y reconocimiento «Club 10»', 'Finanzas y Sistemas', ...bar('.xxxxx')],
    ['**D6 Crecimiento**\n50% → 70%', 'Publicar el plan de carrera a 24 meses y abrir la Academia (Módulo 1)', 'RH', ...bar('xxxxxx')],
    ['**Para todos**', 'Encuesta rápida de 5 preguntas cada mes, pláticas de permanencia y encuesta completa en mayo', 'RH', ...bar('xxxxxa')],
  ];
  const key = (fill, text) => [new TextRun({ text: '■ ', color: fill, size: 16 }), new TextRun({ text: text + '     ', size: 15, color: C.ink2 })];

  return [
    ...moduleHead('02', 'Módulo 2 · Clima Laboral y Experiencia del Colaborador', 'Encuesta «Pulso Ochoa»: escuchar al equipo para mejorar'),
    ficha([['Qué entregamos', 'Encuesta de clima, resultados simulados y plan de 6 meses'], ['Objetivo al que apoya', '{OE-3}{OE-2}'], ['A quién va dirigido', 'Todo el equipo; cada gerente cuida el clima de su sucursal'], ['Cómo sabremos que funciona', 'Clima favorable de 54% a 70% o más en mayo 2027']]),
    gap(),
    p('**Cómo es la encuesta.** Son 27 preguntas: 24 se contestan del 1 al 5, una pregunta si recomendarías Autos Ochoa como lugar para trabajar y dos son abiertas. Es **anónima**, toma unos 8 minutos y se contesta desde el celular con un enlace de WhatsApp. Solo mostramos resultados de grupos de 3 personas o más, para que nadie pueda ser identificado. Complementa la encuesta que pide la ley (NOM-035, Guía de Referencia II).'),
    h3('Las 24 preguntas, en 8 temas', { before: 60 }),
    tbl(qwAll, qRows),
    h3('Cómo se contesta · preguntas 1 a 24'),
    scaleTbl,
    small('**Favorable** = respuestas 4 y 5. Tema **fuerte**: 75% o más · **a vigilar**: 60–74% · **urgente**: menos de 60%.', { before: 50, after: 0 }),
    h3('Tres preguntas más'),
    small('**25.** Del 0 al 10, ¿qué tanto recomendarías Autos Ochoa como lugar para trabajar? (eNPS: los que ponen 9–10 menos los que ponen 0–6)', { after: 20 }),
    small('**26.** ¿Qué es lo que más te hace querer quedarte?', { after: 20 }),
    small('**27.** Si pudieras cambiar una sola cosa de tu trabajo, ¿cuál sería?', { after: 0 }),
    h3('La experiencia del colaborador · seis momentos que vamos a cuidar'),
    cards(journey.map(([t, d], i) => ({ kids: [p(`${i + 1}`, { font: SERIF, bold: true, size: 22, color: C.accent, after: 0 }), p(t, { bold: true, size: SZ.small, after: 10 }), small(d, { size: 15, after: 0, line: 230 })], borders: { ...noCell, top: ln(C.navy, 18) } })), { gap: 100, margins: { top: 50, bottom: 20, left: 0, right: 40 } }),
    gap(),
    why('Con 8 personas el ambiente se nota en el pasillo; con 34 en cinco lugares, ya no. Además, la ley (NOM-035) nos pide revisar el estrés y la carga de trabajo.', 'Un equipo que no sabe quién atiende a cada cliente pierde ventas. Si mejora el clima, mejora el seguimiento y se va menos gente (OE-2 y OE-3).'),

    h2('Resultados simulados · noviembre de 2026', { before: 0, pageBreak: true }),
    small('Así podrían verse los resultados. Los armamos con lo que ya sabemos de la empresa por el diagnóstico interno y la Parte 1; la encuesta real se aplicará en noviembre.'),
    cards(tiles.map(([v, l]) => ({ kids: [p(v, { size: 30, bold: true, after: 10 }), small(l, { size: 15, after: 0, line: 230 })] })), { gap: 120 }),
    h3('% de respuestas favorables por tema · de mayor a menor'),
    img(charts.bars, CW / PX),
    callout('Lo que nos dice', bullets([
      '**Lo más bajo es la claridad de rol (31%):** la gente no sabe quién le da seguimiento a cada cliente. Es la misma falla que hoy deja caer al 97% de los prospectos.',
      '**Carga y descanso (38%):** los asesores contestan WhatsApp fuera de horario y no saben qué domingos les toca trabajar.',
      '**Trato justo (40%):** se siente que las reglas no son iguales para familiares y no familiares.',
      '**Lo mejor es el orgullo (85%):** la historia familiar de 25 años es nuestra fortaleza y la base del Módulo 3.',
    ], { size: SZ.small, after: 20 })),
    h3('Plan de mejora a 6 meses · diciembre 2026 a mayo 2027'),
    dataTable([16, 40, 16, 4.7, 4.7, 4.7, 4.7, 4.7, 4.7], ['Tema · meta', 'Qué vamos a hacer', 'Responsable', ...MONTHS.map(m => ({ c: m, align: AlignmentType.CENTER, margins: { top: 30, bottom: 30, left: 0, right: 0 } }))], plan, { pad: 24 }),
    new Paragraph({ spacing: { before: 40, after: 60 }, children: [...key(C.navy, 'Trabajo'), ...key(C.accent, 'Fecha clave'), new TextRun({ text: 'Los temas fuertes o a vigilar (D2, D7, D8) se siguen con la encuesta mensual.', size: 15, color: C.ink2 })] }),
    h3('Cómo sabremos que funciona'),
    dataTable([45, 25, 30], ['Qué medimos', 'Hoy → meta', 'Cada cuándo'], [
      ['**Clima favorable** (promedio de los 8 temas)', '54% → 70% o más', 'Cada 6 meses; corte en mayo 2027'],
      ['**¿Recomendarías trabajar aquí?** (eNPS)', '+19 → +30 o más', 'Cada 6 meses'],
      ['**Acciones del plan hechas a tiempo**', '90% o más', 'Cada mes, en la junta de Dirección'],
      ['**Faltas no programadas**', '2% o menos', 'Cada mes, con el checador'],
    ], { pad: 26 }),
  ];
}

function modulo3() {
  const [ew1, ew2] = colWidths([4, 1], CW);
  const pillars = [['1 · Decides', 'Diriges tu sucursal y opinas en precios y campañas cada mes.', 'Prueba: consejo de gerentes'], ['2 · Creces', 'De asesor a gerente regional en 24 meses, con reglas claras.', 'Prueba: ascensos publicados'], ['3 · Ganas claro', 'Tu comisión sale del margen real y se paga completa en nómina.', 'Prueba: sueldos publicados'], ['4 · Herramientas', 'CRM con IA, Academia Ochoa y constancias oficiales.', 'Prueba: simulador'], ['5 · Mismas reglas', 'Una familia de 25 años con las mismas reglas para todos.', 'Prueba: Módulo 5']];
  const grp = text => ({ c: [label(text, C.navy, { size: 13, after: 0 })], span: 5, fill: C.soft });
  return [
    ...moduleHead('03', 'Módulo 3 · Employee Branding y Fidelización', 'Por qué trabajar aquí y cómo cuidamos al equipo'),
    ficha([['Qué entregamos', 'Propuesta de valor, plan de comunicación y acciones para retener'], ['Objetivo al que apoya', '{OE-3}{OE-1}'], ['A quién va dirigido', 'Gerentes y asesores senior, y quienes queremos atraer'], ['Cómo sabremos que funciona', 'Menos del 10% de gerentes se va al año']]),
    gap(),
    tbl([ew1, ew2], [new TableRow({
      cantSplit: true,
      children: [
        cell(ew1, [p('«Aquí no vendes autos: diriges un negocio, con 25 años de respaldo y la tecnología que ningún otro lote de Monterrey tiene.»', { font: SERIF, size: 25, color: C.white, after: 0, line: 270 })], { fill: C.navy, vAlign: VerticalAlign.CENTER, margins: { top: 140, bottom: 140, left: 220, right: 120 } }),
        cell(ew2, [p('Lo que ofrecemos (propuesta de valor al empleado, EVP)', { size: 14, color: 'B8C6DC', caps: true, spacing: 8, align: AlignmentType.RIGHT, after: 0, line: 230 })], { fill: C.navy, vAlign: VerticalAlign.CENTER, margins: { top: 140, bottom: 140, left: 80, right: 220 } }),
      ],
    })]),
    gap(),
    cards(pillars.map(([t, d, pr]) => ({ kids: [p(t, { bold: true, color: C.navy, size: SZ.small, after: 20 }), small(d, { size: 15, after: 30, line: 230 }), p(pr, { size: 14, color: C.muted, after: 0 })] })), { gap: 90, margins: { top: 60, bottom: 60, left: 90, right: 80 } }),
    h3('Plan de comunicación · hacia dentro y hacia fuera'),
    dataTable([22, 33, 14, 13, 18], ['Qué hacemos', 'Qué contamos', 'Cada cuándo', 'Quién', 'Cómo lo medimos'], [
      [grp('Hacia dentro · para que el talento se quede')],
      ['**Bienvenida** con carta del fundador', 'Nuestra historia, lo que ofrecemos y las reglas', 'Al entrar', 'RH', 'Bienvenida en 30 días'],
      ['**«Junta Ochoa»** de 30 min', 'Cómo van las 5 sucursales y reconocimientos', 'Cada mes', 'Dir. General', 'Asistencia de 90%'],
      ['**Boletín «Motor Ochoa»**', 'Primeras ventas, ascensos y logros', 'Cada 15 días', 'Marketing y RH', 'Cuántos lo leen'],
      ['**Vacantes primero para casa**', 'Toda vacante se publica adentro 5 días antes', 'Cada vacante', 'RH', 'Puestos cubiertos de casa'],
      [grp('Hacia fuera · para atraer al gerente antes que la competencia')],
      ['**Videos «Detrás de Ochoa»**', 'El equipo, el fundador y los ascensos', '3 al mes', 'Contenido', 'Costo por candidato'],
      ['**Página de empleos**', 'Sueldos, carrera y postulación por WhatsApp', 'Siempre', 'Marketing', 'Visitas que se postulan'],
      ['**Campaña por apertura**', '«Buscamos a quien dirija Ochoa Guadalupe»', '5 meses antes', 'Marketing y RH', 'Candidatos por vacante'],
      ['**LinkedIn y universidades**', 'Testimonios y entrar sin experiencia en autos', 'Cada semana', 'RH', 'Contratos por canal'],
    ], { pad: 22 }),
    h3('Para que la gente se quede'),
    ...bullets([
        '**Plática cada 3 meses** con gerentes y asesores senior: ¿qué te haría quedarte? Respuesta en 15 días.',
        '**Aviso en el CRM** si alguien baja su actividad o su ánimo: su jefe lo busca en menos de 7 días.',
        '**Ascensos primero para la gente de casa:** 20% de las gerencias el primer año y 60% después.',
        '**Beneficios que crecen con los años:** fondo de ahorro, auto a precio de costo e inversión en unidades desde el año 2.',
        '**«Club 10»:** la sucursal que se mantiene en su meta 3 meses gana viernes corto y cena con el fundador.',
        '**$8,000 por recomendar** a alguien que se queda 90 días, y puertas abiertas para quien se fue bien.',
    ], { size: SZ.small, after: 20 }),
    h3('Cómo sabremos que funciona'),
    dataTable([50, 25, 25], ['Qué medimos', 'Meta', 'Cada cuándo'], [
      ['Gerentes que se van sin que queramos', '10% o menos', 'Cada año'],
      ['Nuevos que siguen al año', '85% o más', 'Cada año'],
      ['Contratados por recomendación', '40% o más', 'Cada año'],
      ['Ofertas de gerente aceptadas', '85% o más', 'Por vacante'],
      ['Pláticas de permanencia hechas', '100%', 'Trimestral'],
    ], { pad: 18 }),
    gap(),
    why('Autos Trefa tiene más de 166 mil seguidores y más presupuesto. No les ganamos con publicidad; les ganamos con nuestra historia y con cómo tratamos a la gente.', 'Cada gerente que no se va ahorra ~$500,000 (OE-3). Cubrir vacantes con recomendaciones y canales propios ayuda a tener al gerente listo antes de abrir (OE-1).'),
  ];
}

function modulo4(charts) {
  const comp = (title, desc) => [p(title, { bold: true, size: SZ.small, after: 10 }), small(desc, { size: 15, after: 0, line: 230 })];
  const exp = items => bullets(items, { size: 15, after: 10 });
  const scale = [['1 · Todavía no', 'No lo hace, aun con ayuda.'], ['2 · Con ayuda', 'Lo hace con su padrino o a veces.'], ['3 · Bien y siempre', 'Lo hace de forma constante. Es lo esperado.'], ['4 · Avanzado', 'Lo hace bien incluso en situaciones difíciles.'], ['5 · Ejemplo', 'Lo enseña a otros.']];
  const scw = colWidths([1, 1, 1, 1, 1], CW);
  const weights = [[40, '2A78D6', C.white, 'Su jefe'], [25, 'EB6834', C.white, 'Su equipo'], [15, '1BAF7A', C.ink, 'Otros gerentes'], [10, 'EDA100', C.ink, 'Áreas internas'], [10, 'E87BA4', C.ink, 'Él mismo']];
  const ww = colWidths(weights.map(w => w[0]), CW);
  const fw = colWidths([1, 0.12, 1, 0.12, 0.7], CW);
  const fbox = (w, kids, fill = C.soft, border = C.line) => cell(w, kids, { fill, borders: box(border), margins: { top: 60, bottom: 60, left: 110, right: 90 }, vAlign: VerticalAlign.CENTER });
  const sym = (w, s) => cell(w, [p(s, { size: 26, bold: true, align: AlignmentType.CENTER, after: 0 })], { vAlign: VerticalAlign.CENTER, margins: { top: 0, bottom: 0, left: 0, right: 0 } });
  const timeline = [['Cada mes', 'Plática 1 a 1', '30 minutos con su jefe para revisar el tablero.'], ['Gerente nuevo', 'Días 30, 90 y 120', 'El día 120 nos dice si contratamos bien.'], ['Febrero y agosto', 'Evaluación 360°', 'La primera será en febrero de 2027.'], ['Marzo y septiembre', 'Mapa de talento', 'Dirección y RH ubican a cada persona.'], ['Agosto', 'Cierre del año', 'Plan de desarrollo personal y revisión de sueldo.']];
  return [
    ...moduleHead('04', 'Módulo 4 · Gestión del Desempeño y Competencias', 'Qué esperamos de un gerente: seis competencias medibles'),
    ficha([['Qué entregamos', 'Modelo de 6 competencias y evaluación 360°'], ['Objetivo al que apoya', '{OE-2}{OE-3}'], ['A quién va dirigido', 'Gerentes (360°); asesores con una versión corta'], ['Cómo sabremos que funciona', '4 de cada 5 gerentes en su meta a los 120 días']]),
    gap(),
    p('Hoy no evaluamos formalmente a nadie. Proponemos **seis competencias**, tomadas del perfil del gerente de la Parte 1: **dos técnicas, dos de liderazgo y dos digitales**. Cada una dice qué esperamos ver en el día a día y **con qué dato del CRM lo comprobamos**, para que la evaluación no dependa de la simpatía.'),
    dataTable([11, 23, 37, 29], ['Tipo', 'Competencia', 'Qué esperamos ver', 'Con qué dato lo comprobamos'], [
      [{ c: '**Técnicas**', rowSpan: 2 }, { c: comp('T1 · Seguimiento a clientes', 'Que cada cliente tenga una siguiente acción con fecha.') }, { c: exp(['Termina el día sin pendientes vencidos.', 'Responde en menos de 2 horas.', 'Recupera a los clientes que se enfriaron.']) }, { c: 'Pendientes vencidos (hoy 29) · clientes que avanzan (hoy 3%) · tiempo de respuesta', size: 15 }],
      [{ c: comp('T2 · Cierre y crédito', 'Entender lo que el cliente necesita y puede pagar.') }, { c: exp(['Ofrece al menos 2 opciones de crédito.', 'Arma el expediente completo a la primera.', 'No regala el margen con descuentos sin autorizar.']) }, { c: 'Ventas sobre clientes calificados (meta 18%) · margen por auto (meta $32,000) · expedientes completos', size: 15 }],
      [{ c: '**Liderazgo**', rowSpan: 2 }, { c: comp('L1 · Formar al equipo', 'Que la sucursal no dependa solo del gerente.') }, { c: exp(['Platica 30 minutos a la semana con cada asesor.', 'Lo acompaña en conversaciones reales.', 'Prepara a alguien que lo pueda reemplazar.']) }, { c: 'Asesores que se van · asesores que venden 4 autos o más · reemplazo identificado', size: 15 }],
      [{ c: comp('L2 · Liderar con integridad', 'Mismas reglas para todos y sin atajos con la ley.') }, { c: exp(['Publica el rol de domingos con un mes de anticipación.', 'No paga por fuera ni deja trabajar sin contrato.', 'Trata igual a familiares y no familiares.']) }, { c: 'Incidencias (meta 0) · calificación de «trato justo» en la encuesta de su sucursal', size: 15 }],
      [{ c: '**Digitales**', rowSpan: 2 }, { c: comp('D1 · Decidir con datos', 'Usar los tableros para saber dónde enfocarse.') }, { c: exp(['Revisa el tablero cada mañana.', 'Ajusta el precio de los autos con más de 90 días.', 'Detecta a tiempo a quien necesita apoyo.']) }, { c: 'Días promedio en inventario · autos con más de 90 días (hoy 91%) · uso del tablero', size: 15 }],
      [{ c: comp('D2 · Trabajar con IA', 'Aprovechar el asistente del CRM y WhatsApp.') }, { c: exp(['Usa las respuestas que sugiere la IA.', 'Programa recordatorios automáticos.', 'Publica en redes los autos de su sucursal.']) }, { c: 'Conversaciones con apoyo de IA · clientes que llegan por redes', size: 15 }],
    ], { pad: 30 }),
    h3('Cómo calificamos cada competencia · escala de 1 a 5'),
    tbl(scw, [new TableRow({ cantSplit: true, children: scale.map(([t, d], i) => cell(scw[i], [p(t, { bold: true, size: SZ.small, after: 10 }), small(d, { size: 15, after: 0, line: 230 })], { fill: i === 2 ? 'DDE9F8' : C.soft, borders: { ...noCell, right: i < 4 ? ln(C.white, 12) : NONE }, margins: { top: 50, bottom: 50, left: 90, right: 70 } })) })]),
    gap(),
    why('En la Parte 1 nos pusimos la meta de medir si contratamos bien, pero hoy no tenemos cómo. Además, el periodo de prueba de un gerente (hasta 180 días) solo se puede terminar con evidencia.', 'Estas seis conductas son las que vienen antes de la venta. Un gerente que las cumple sostiene 10 a 12 autos al mes en su sucursal (OE-2).'),

    h2('Cómo vamos a evaluar al gerente: evaluación 360°', { before: 240 }),
    p('Al gerente lo evalúan las personas que ven su trabajo desde distintos ángulos. A los asesores y coordinadores los evalúan su jefe y ellos mismos (180°). Son 30 preguntas en línea, 5 por competencia, y los resultados salen solos en el tablero.'),
    h3('Quién evalúa y cuánto pesa su opinión'),
    tbl(ww, [new TableRow({ cantSplit: true, children: weights.map(([v, fill, color, name], i) => cell(ww[i], [p(`${v}% · ${name}`, { size: 15, bold: true, color, align: AlignmentType.CENTER, after: 0 })], { fill, borders: { ...noCell, right: i < 4 ? ln(C.white, 12) : NONE }, margins: { top: 40, bottom: 40, left: 20, right: 20 }, vAlign: VerticalAlign.CENTER })) })]),
    gap(),
    dataTable([17, 33, 8, 42], ['Quién evalúa', 'Quiénes son', 'Peso', 'En qué se fija más'], [
      ['**Su jefe**', 'Dirección Comercial', '40%', 'Las seis competencias y los datos del CRM'],
      ['**Su equipo**', '3 asesores y el coordinador (anónimo)', '25%', 'Si los forma y los trata con justicia'],
      ['**Otros gerentes**', '2 gerentes de otras sucursales', '15%', 'Si colabora y decide con datos'],
      ['**Áreas internas**', 'Compras, Finanzas y Marketing', '10%', 'Expedientes, margen, cumplimiento y contenido'],
      ['**Él mismo**', 'El propio gerente', '10%', 'La diferencia con los demás le muestra qué mejorar'],
    ], { pad: 24 }),
    h3('La calificación final: lo que logró y cómo lo logró pesan igual'),
    tbl(fw, [new TableRow({
      cantSplit: true,
      children: [
        fbox(fw[0], [p('**50%** Resultados (lo que logró)', { size: SZ.small, after: 10 }), small('autos vendidos · % de cierre · margen por auto · pendientes vencidos · rotación de su equipo', { size: 15, after: 0, line: 230 })]),
        sym(fw[1], '+'),
        fbox(fw[2], [p('**50%** Competencias (cómo lo logró)', { size: SZ.small, after: 10 }), small('promedio de las seis competencias en la evaluación 360°', { size: 15, after: 0, line: 230 })]),
        sym(fw[3], '='),
        fbox(fw[4], [p('Calificación 1 a 5', { size: SZ.small, bold: true, color: C.white, after: 10 }), p('se usa en el mapa de talento (Módulo 6)', { size: 15, color: 'C9D4E6', after: 0, line: 230 })], C.navy, C.navy),
      ],
    })]),
    h3('Cada cuándo'),
    cards(timeline.map(([s, b, d]) => ({ kids: [p(s, { size: 13, caps: true, bold: true, color: C.muted, spacing: 8, after: 10 }), p(b, { bold: true, size: SZ.small, after: 10 }), small(d, { size: 15, after: 0, line: 230 })], borders: { ...noCell, top: ln(C.navy, 18) } })), { gap: 100, margins: { top: 50, bottom: 20, left: 0, right: 40 } }),
    h3('Qué pasa con cada calificación'),
    dataTable([15, 18, 67], ['Calificación', 'Nivel', 'Qué sigue'], [
      ['4.5 – 5.0', '**Sobresaliente**', 'Candidato a crecer y bono de permanencia'],
      ['3.5 – 4.4', '**Supera**', 'Plan para crecer: Nivel 3 de la Academia y un proyecto de mejora'],
      ['2.8 – 3.4', '**Cumple**', 'Plan para reforzar sus competencias más bajas'],
      ['2.0 – 2.7', '**En desarrollo**', 'Plan de mejora de 90 días con su padrino'],
      ['< 2.0', '**No cumple**', 'Plan por escrito; en periodo de prueba se consulta a la Comisión Mixta antes de decidir'],
    ], { pad: 22 }),
    h3('Ejemplo · reporte 360° del Colaborador B, gerente de Guadalupe (simulado)'),
    img(charts.dots, charts.dots.w),
    callout('Cómo se lee', 'Él se da 3.8 en **formar a su equipo**, pero los demás le dan 2.7; coincide con que un asesor se fue. Calificación final: 50% × 4.6 (resultados) + 50% × 3.6 (competencias) = **4.1**. Siguiente paso: coaching para formar a su equipo.', { size: SZ.small }),
    gap(),
    callout('Reglas para que sea justa', bullets(['Las reglas se acuerdan antes de evaluar y quien evalúa recibe 1 hora de capacitación para no calificar por simpatía.', 'Nadie ve quién dijo qué: solo se muestran grupos de 3 o más.', 'Cada año revisamos si las calificaciones de verdad predicen las ventas.'], { size: SZ.small, after: 20 })),
    h3('Cómo sabremos que funciona'),
    dataTable([50, 25, 25], ['Qué medimos', 'Meta', 'Cada cuándo'], [
      ['Gerentes en su meta a los 120 días', '80% o más', 'Día 120'],
      ['Evaluaciones hechas a tiempo', '100%', 'Cada ciclo'],
      ['Gerentes con nivel 3 o más en las seis', '80% o más', 'Agosto 2027'],
      ['Planes personales con 80% de avance', '90% o más', 'Cada 6 meses'],
    ], { pad: 18 }),
  ];
}

function modulo5() {
  const policy = (tag, title, para, items, kpis, owner, base) => ({
    kids: [
      label(tag, C.accent, { size: 12, after: 20 }),
      p(title, { font: SERIF, bold: true, size: 21, after: 80, line: 240 }),
      label('Para qué', C.navy, { size: 12 }), small(para, { size: 15, line: 230 }),
      label('Qué vamos a hacer', C.navy, { size: 12, before: 40 }), ...bullets(items, { size: 15, after: 20 }),
      label('Cómo sabremos que funciona', C.navy, { size: 12, before: 60 }), ...kpis.map(([k, t]) => p(`${k} **·** ${t}`, { size: 15, after: 10, line: 230 })),
      label('Responsable', C.navy, { size: 12, before: 60 }), small(owner, { size: 15, after: 60 }),
      p(base, { size: 13, color: C.muted, after: 0, line: 220 }),
    ],
  });
  return [
    ...moduleHead('05', 'Módulo 5 · Responsabilidad Social y Talento', 'Tres políticas para que todos tengan las mismas oportunidades'),
    ficha([['Qué entregamos', 'Tres políticas de responsabilidad social (ASG)'], ['Objetivo al que apoya', '{OE-3}{OE-1}'], ['A quién va dirigido', 'Toda la empresa; cada gerente las aplica en su sucursal'], ['Cómo sabremos que funciona', 'Diferencia de sueldo en el mismo puesto de 5% o menos']]),
    gap(),
    p('**Por qué estas tres.** En la venta de autos trabajan sobre todo hombres, es común pagar comisiones por fuera y trabajar fines de semana sin control. Además, somos una empresa familiar en su segunda generación, donde es fácil que parezca que hay reglas distintas para la familia. Estas políticas atacan justo eso y se reportan con indicadores internacionales (GRI).', { after: 100 }),
    cards([
      policy('Política 1 · Equidad salarial y de oportunidades', 'Mismo puesto, mismo sueldo', 'Que lo que ganas y tus ascensos dependan de tu puesto y tu desempeño, no de tu apellido, tu género o de cuánto negociaste.',
        ['Publicar 6 bandas de sueldo; nadie queda fuera de su banda, tampoco los familiares.', 'Comisiones con la misma tabla para todos y pagadas completas en nómina.', 'Cada 6 meses revisar diferencias de sueldo entre hombres y mujeres y entre familiares y no familiares; si pasan de 5%, se corrigen.', 'Ascensos decididos con la evaluación y un comité que incluye a alguien que no es de la familia.'],
        [['Diferencia de sueldo de 5% o menos', 'cada 6 meses'], ['Todos dentro de su banda', 'cada 6 meses'], ['Ascensos documentados 100%', 'cada ascenso']],
        'Dirección General y RH', 'Base: LFT arts. 86 y 133 · NMX-R-025 · GRI 405 · ODS 5, 8 y 10'),
      policy('Política 2 · Inclusión en la selección', 'Contratamos por lo que sabes hacer', 'Tener más candidatos para cubrir 26 puestos a tiempo, quitando filtros que no dicen nada de cómo trabaja una persona.',
        ['Vacantes sin límite de edad y sin pedir foto, estado civil ni «buena presentación».', 'El primer filtro no muestra nombre, edad ni género; decidimos con la prueba práctica y la entrevista.', 'Entre los finalistas para gerente y asesor habrá al menos una mujer.', 'Abrimos la puerta a gente sin experiencia en autos (universitarios, mujeres, personas de más de 45 años) con 3 meses de capacitación pagada, y hacemos ajustes para personas con discapacidad.'],
        [['Vacantes sin requisitos que discriminen 100%', 'cada mes'], ['Finalistas con al menos una mujer 80%', 'cada 3 meses'], ['Mujeres contratadas en ventas 30%', 'cada año']],
        'RH y comité de selección', 'Base: LFT arts. 2, 3 y 133 · Ley Federal para Prevenir y Eliminar la Discriminación · GRI 405 y 406 · ODS 5 y 10'),
      policy('Política 3 · Balance vida-trabajo', 'Tiempo para tu vida', 'Abrir sábados y domingos sin que eso cueste la vida personal del equipo ni nos meta en problemas legales.',
        ['Nadie trabaja más de 2 domingos al mes; el rol se publica con un mes de anticipación y el domingo se paga con 25% extra.', 'Derecho a desconectarse: fuera de la guardia pagada, el asistente de IA contesta; nadie está obligado a responder.', 'Jornada de 46 horas desde el 1 de enero de 2027, sin bajar sueldos.', 'Oficinas: 1 día en casa por semana. Para todos: 5 días de paternidad extra, día libre en tu cumpleaños y sala de lactancia.'],
        [['Encuesta «carga y descanso» 60% o más', 'cada 6 meses'], ['Horas extra 4 o menos por semana', 'cada mes'], ['Nadie con más de 2 domingos', 'cada mes']],
        'Gerentes y RH', 'Base: LFT arts. 59, 66–71, 132 fr. XXXIV, 170 y 330-A · NOM-035 · GRI 401 · ODS 3 y 8'),
    ], { gap: 120, split: true, margins: { top: 90, bottom: 90, left: 110, right: 100 } }),
    gap(),
    why('Al pasar de 8 a 34 personas, cualquier mala costumbre se multiplica por cinco. Escribir estas reglas desde el principio evita repetir en cada sucursal lo que la ley sanciona.', 'Más candidatos para cubrir vacantes a tiempo (OE-1), menos rotación y menos riesgo de demandas (OE-3). Además nos prepara para la certificación de igualdad del Módulo 8.'),
  ];
}

function modulo6() {
  const axisW = 900;
  const cw3 = colWidths([1, 1, 1], CW - axisW);
  const fills = { l0: 'F3F4F6', l1: 'E8EEF7', l2: 'D9E6F7', l3: 'C3D8F4', l4: 'A9C8F0' };
  const nb = [
    ['Potencial alto', [['7', 'Diamante en bruto', 'Acelerar su aprendizaje', 'l2', 'C'], ['8', 'Futuro líder', 'Prepararlo para crecer', 'l3'], ['9', 'Estrella', 'Listo; cuidarlo', 'l4', 'A']]],
    ['Potencial medio', [['4', 'Irregular', 'Entender qué pasa', 'l1'], ['5', 'Pilar del equipo', 'Consolidar', 'l2'], ['6', 'Alto desempeño', 'Mantenerlo en su puesto', 'l3', 'B']]],
    ['Potencial bajo', [['1', 'Bajo desempeño', 'Plan de 90 días', 'l0'], ['2', 'Cumple', 'Motivar', 'l1'], ['3', 'Experto confiable', 'Reconocer; que enseñe', 'l2']]],
  ];
  const badge = who => new TextRun({ text: ` ${who} `, bold: true, size: 17, color: C.white, shading: { type: ShadingType.CLEAR, fill: C.deep, color: 'auto' } });
  const nbCell = ([n, t, a, l, who], w) => cell(w, [
    new Paragraph({ spacing: { after: 10, lineRule: LineRuleType.AUTO, line: 225 }, children: [new TextRun({ text: `${n} · `, size: 14, color: C.muted, bold: true }), new TextRun({ text: t, bold: true, size: 15 })] }),
    p(a, { size: 14, color: C.ink2, after: who ? 40 : 0, line: 220 }),
    ...(who ? [new Paragraph({ spacing: { after: 0 }, children: [badge(who)] })] : []),
  ], { fill: fills[l], borders: { top: ln(C.white, 16), bottom: ln(C.white, 16), left: ln(C.white, 16), right: ln(C.white, 16) }, margins: { top: 50, bottom: 50, left: 70, right: 50 } });
  const nbRows = nb.map(([pot, cells]) => new TableRow({
    height: { value: 820, rule: HeightRule.ATLEAST }, cantSplit: true,
    children: [cell(axisW, pot.split(' ').map(t => p(t, { size: 14, color: C.muted, bold: true, after: 0, line: 220 })), { vAlign: VerticalAlign.CENTER, margins: { top: 0, bottom: 0, left: 0, right: 60 } }), ...cells.map((c, i) => nbCell(c, cw3[i]))],
  }));
  nbRows.push(new TableRow({ children: [cell(axisW, [tiny()]), ...['Desempeño bajo', 'Desempeño medio', 'Desempeño alto'].map((t, i) => cell(cw3[i], [p(t, { size: 14, color: C.muted, bold: true, align: AlignmentType.CENTER, after: 0 })], { margins: { top: 30, bottom: 0, left: 0, right: 0 } }))] }));
  const nineBox = tbl([axisW, ...cw3], nbRows);
  const person = (who, name, sub, rows) => ({
    kids: [
      new Paragraph({ spacing: { after: 0 }, children: [badge(who), new TextRun({ text: `  ${name}`, bold: true, size: SZ.small })] }),
      p(sub, { size: 14, color: C.muted, after: 40 }),
      ...rows.map(([k, v, hl]) => new Paragraph({ spacing: { after: 20, lineRule: LineRuleType.AUTO, line: 230 }, children: [new TextRun({ text: `${k}:  `, size: 13, bold: true, allCaps: true, color: C.muted, characterSpacing: 6 }), ...runs(v, { size: 15, color: hl ? C.accent : C.ink, bold: hl || undefined })] })),
    ],
  });
  const persons = [
    person('A', 'Colaborador A · Asesor senior en la matriz', '3 años en la empresa', [['Desempeño', '**Alto (4.4):** vende 5.6 autos al mes, cierra el 22% y casi nunca deja clientes sin seguimiento.'], ['Potencial', '**Alto:** aprobó los niveles 1 y 2 muy rápido, ya enseña a 2 asesores y quiere ser gerente.'], ['Cuadro', '9 · Estrella — listo ahora', true], ['Qué haremos', 'Será gerente de Apodaca (abre en mayo). Antes: 30 días acompañando a un gerente, Nivel 3 y bono de permanencia. Como otras agencias lo pueden buscar, platicamos con él cada mes.']]),
    person('B', 'Colaborador B · Gerente de Guadalupe', 'Contratado de fuera · 120 días', [['Desempeño', '**Alto (4.1):** 11 autos al mes; llegó a su meta en el segundo mes.'], ['Potencial', '**Medio:** le cuesta formar a su equipo (2.7 en esa competencia) y un asesor ya se fue.'], ['Cuadro', '6 · Alto desempeño — mantenerlo en su puesto', true], ['Qué haremos', 'Coaching para formar a su equipo, con un padrino. Volvemos a revisar en septiembre. Todavía no es candidato a gerente regional.']]),
    person('C', 'Colaborador C · Asesor nuevo', 'Entró sin experiencia en autos · 4 meses', [['Desempeño', '**Bajo (2.5):** vende 2.3 autos al mes; la meta es 4. Sigue aprendiendo.'], ['Potencial', '**Alto:** la mejor calificación de su generación en la Academia (96 de 100) y muy bueno con la IA.'], ['Cuadro', '7 · Diamante en bruto — acelerar', true], ['Qué haremos', 'Plan de 90 días con el Colaborador A como padrino: 20 conversaciones acompañadas y meta de 4 autos al mes en el mes 7.']]),
  ];
  return [
    ...moduleHead('06', 'Módulo 6 · Sucesión y Talento Clave', 'Mapa de talento (9-Box): quién está listo para crecer'),
    ficha([['Qué entregamos', 'Mapa de talento de 9 cuadros y plan de sucesión'], ['Objetivo al que apoya', '{OE-1}{OE-3}'], ['A quién va dirigido', 'Gerentes y quienes podrían llegar a serlo'], ['Cómo sabremos que funciona', '6 de cada 10 puestos clave con reemplazo listo']]),
    gap(),
    p('**Cómo funciona.** El mapa cruza dos preguntas: **¿qué tan bien hace hoy su trabajo?** (desempeño: la calificación del Módulo 4) y **¿qué tanto puede crecer?** (potencial: qué tan rápido aprende en la Academia, cómo lidera, si quiere crecer y una prueba psicométrica). Dirección y RH lo revisan en marzo y septiembre. Ejemplo simulado a marzo de 2027.', { after: 100 }),
    nineBox,
    h3('Los tres colaboradores del ejemplo'),
    cards(persons, { gap: 120, split: true, margins: { top: 70, bottom: 70, left: 100, right: 90 } }),
    h3('Cómo sabremos que funciona'),
    dataTable([50, 25, 25], ['Qué medimos', 'Meta', 'Cada cuándo'], [
      ['Puestos clave con reemplazo listo', '60%', 'Cada 6 meses'],
      ['Gerencias cubiertas por gente de casa', '20% → 60%', 'Año 1 → año 2'],
      ['Talento de los cuadros 7 a 9 que se queda', '100%', 'Cada año'],
      ['Cuadros 7 a 9 con plan personal', '100%', 'Cada 6 meses'],
    ], { pad: 16 }),
    h3('Plan de sucesión que resulta'),
    dataTable([25, 14, 20, 11, 30], ['Puesto', 'Hoy lo ocupa', 'Quién lo podría ocupar', '¿Cuándo?', 'Siguiente paso'], [
      ['Gerente · Apodaca (Fase 2)', 'Vacante', 'Colaborador A', '**Ya**', 'Ascenso en marzo; ya tiene los niveles 1 y 2'],
      ['Gerente · Guadalupe', 'Colaborador B', 'Un asesor senior por definir', '1 a 2 años', 'Identificarlo en la evaluación de agosto'],
      ['Asesor senior · Matriz', 'Colaborador A', 'Colaborador C', '1 a 2 años', 'Plan de 90 días'],
      ['Gerente regional (año 2)', 'Puesto nuevo', 'El gerente mejor evaluado', '2 años o más', 'Nivel 3 y revisión de septiembre'],
    ], { pad: 22 }),
    gap(),
    why('Hoy todos los gerentes vienen de fuera y casi todo el conocimiento está en el fundador. Si un gerente se va sin reemplazo, la sucursal queda sin líder 45 días, más 120 de aprendizaje del nuevo.', 'Un reemplazo de casa ya conoce la empresa y ya está certificado: es la forma más barata de tener al gerente listo antes de abrir (OE-1) y de cuidar a quien más ofertas recibe (OE-3).'),
  ];
}

function modulo7() {
  const flow = [['150 días antes', 'Se aprueba la vacante'], ['120 días', 'Se publica'], ['90 días', 'Prueba y oferta'], ['60 días', 'Gerente certificado', true], ['30 días', '3 asesores listos'], ['Día 0', 'Apertura'], ['Día 120', '10 autos al mes', true]];
  const fw = colWidths(flow.map(() => 1), CW);
  const roleW = colWidths([0.9, 3.1], CW);
  const roles = [['Quién decide qué va primero', 'Coordinación de RH, según los 3 objetivos.'], ['Quién facilita', 'El especialista de sistemas, que ya usa tableros en ventas.'], ['Quién trabaja', 'RH, sistemas, marketing, finanzas y un gerente invitado.'], ['Quién revisa', 'Dirección General y Comercial al final de cada ciclo.'], ['Reuniones', 'Planeación de 1 hora · 15 minutos diarios · revisión con Dirección · 30 minutos para ver qué mejorar.'], ['Tablero', 'Por hacer → Haciendo (máximo 3 por persona) → En revisión → Listo. «Listo» = publicado, medido y comunicado.']];
  return [
    ...moduleHead('07', 'Módulo 7 · Digitalización y Metodologías Ágiles', 'Herramientas que ya tenemos y trabajo en ciclos cortos'),
    ficha([['Qué entregamos', '4 herramientas digitales y trabajo en ciclos de 2 semanas (Scrum)'], ['Objetivo al que apoya', '{OE-1}{OE-2}{OE-3}'], ['A quién va dirigido', 'RH y gerentes, que usarán los tableros a diario'], ['Cómo sabremos que funciona', 'Vacante de gerente cubierta en 45 días o menos']]),
    h3('Herramientas digitales · para qué sirven y cómo las ponemos en marcha'),
    dataTable([21, 21, 36, 10, 12], ['Herramienta', 'Para qué la usamos en RH', 'Cómo la ponemos en marcha', 'Apoya a', 'Costo'], [
      ['**1 · CRM de Autos Ochoa con IA** (ya lo tenemos)', 'Reclutar, evaluar (360°), plan de sucesión y avisar si alguien se quiere ir', 'El especialista de sistemas le agrega tres funciones: tablero de candidatos, formulario 360° y aviso de salida. La IA califica la prueba práctica igual para todos.', 'Módulos 1 · 2 · 4 · 6 · 8', '$0 en licencias; ~60 h de trabajo interno'],
      ['**2 · Google Classroom y Looker Studio** (gratuitos)', 'Cursos de la Academia y tablero de indicadores', 'Un salón en línea por nivel; un tablero que junta CRM, nómina y encuesta. Cada gerente ve su sucursal y Dirección ve todo.', 'Módulos 1 · 2 · 3 · 4', '$0'],
      ['**3 · Nómina en la nube con checador**', 'Pagar comisiones en el recibo, registrar horarios y pedir vacaciones desde el celular', 'Elegir proveedor en octubre (por ejemplo Runa, Worky o Buk), probar en la matriz y tenerlo en todas las sedes antes de diciembre de 2026, como pide la ley.', 'Módulos 2 · 5', '~$60 por persona al mes (Parte 1)'],
      ['**4 · WhatsApp Business**', 'Postularse, contestar encuestas, pregunta diaria y respuestas fuera de horario', 'Mensajes automáticos: postulación en 3 preguntas, encuesta mensual, pregunta del día y respuesta del asistente fuera de la guardia.', 'Módulos 1 · 2 · 3 · 5', 'Ya lo usamos'],
    ], { pad: 26 }),
    small('**Industria 4.0, en simple:** IA que sugiere y califica · aviso anticipado de salidas · todo en la nube y en el celular · mensajes automáticos · aprender jugando · decidir con datos al día.', { before: 60, after: 0 }),
    h3('Metodología ágil · trabajar en ciclos de dos semanas (Scrum) con un tablero de tareas (Kanban)'),
    tbl(roleW, roles.map(([k, v]) => new TableRow({ cantSplit: true, children: [cell(roleW[0], [p(k, { size: SZ.small, bold: true, color: C.navy, after: 0, line: 230 })], { borders: { ...noCell, bottom: ln(C.line2, 4) }, margins: { top: 30, bottom: 30, left: 0, right: 100 } }), cell(roleW[1], [p(v, { size: SZ.small, after: 0, line: 230 })], { borders: { ...noCell, bottom: ln(C.line2, 4) }, margins: { top: 30, bottom: 30, left: 0, right: 0 } })] }))),
    h3('Primeros seis ciclos · octubre a diciembre 2026'),
    dataTable([10, 90], ['Ciclo', 'Qué entregamos'], [
      ['1', 'Tablero de indicadores y tablero de candidatos en el CRM'],
      ['2', 'Academia Nivel 1 en Classroom y las 3 políticas publicadas'],
      ['3', 'Nómina y checador en la matriz; página de empleos'],
      ['4', 'Checador en todas las sedes y encuesta «Pulso Ochoa»'],
      ['5', 'Plan de mejora del clima y calculadora de comisiones'],
      ['6', 'Formulario 360° y competencias en el CRM'],
    ], { pad: 14 }),
    small('**Ejemplo de tarea:** «Como gerente nuevo, quiero mi curso del Nivel 1 en Classroom para certificarme 60 días antes de abrir».', { before: 50, after: 0 }),
    h3('Cada apertura es un tablero con fechas fijas'),
    tbl(fw, [new TableRow({ cantSplit: true, children: flow.map(([b, t, hit], i) => cell(fw[i], [p(b, { bold: true, size: 15, color: hit ? C.white : C.navy, after: 0 }), p(t, { size: 14, color: hit ? C.white : C.ink, after: 0, line: 220 })], { fill: hit ? C.navy : C.soft, borders: { ...noCell, right: i < flow.length - 1 ? ln(C.white, 16) : NONE }, margins: { top: 50, bottom: 50, left: 80, right: 50 } })) })]),
    gap(),
    h3('Cómo sabremos que funciona'),
    dataTable([50, 25, 25], ['Qué medimos', 'Meta', 'Cada cuándo'], [
      ['Días para cubrir una vacante de gerente', '45 o menos', 'Cada vacante'],
      ['Tareas del ciclo terminadas', '80% o más', 'Cada 2 semanas'],
      ['Gerentes certificados 60 días antes', '100%', 'Cada apertura'],
      ['Gerentes que usan el tablero a diario', '5 de 5', 'Cada mes'],
    ], { pad: 16 }),
    gap(),
    why('Una sola persona de RH tiene que cubrir 26 contrataciones y 4 aperturas en un año. Sin tableros ni ciclos cortos, RH se vuelve el cuello de botella. Con ellos, Dirección ve el avance cada dos semanas y ninguna sucursal abre sin gerente.'),
  ];
}

function modulo8() {
  const steps = [['1', 'Grabar', '12 charlas con el fundador (con su permiso) y el historial del CRM sin nombres de clientes.'], ['2', 'Ordenar', 'Unas 60 prácticas —valuar, tomar autos a cuenta, crédito, revisión— cargadas en el asistente de IA del CRM.'], ['3', 'Mejorar', 'Reuniones de 30 minutos cada dos semanas en cada sucursal (Kaizen): cada idea que funciona se suma al asistente.'], ['4', 'Llevar a todos', 'Prueba de 90 días en la matriz y Guadalupe (febrero–abril 2027) y después en las 5 sucursales.']];
  const tag = (text, fill = C.accent) => new TextRun({ text: ` ${text} `, bold: true, allCaps: true, size: 12, color: C.white, shading: { type: ShadingType.CLEAR, fill, color: 'auto' } });
  const certName = (name, sub, t, fill) => [new Paragraph({ spacing: { after: 10, lineRule: LineRuleType.AUTO, line: 230 }, children: [new TextRun({ text: name + ' ', bold: true, size: SZ.small }), tag(t, fill)] }), ...(sub ? [small(sub, { size: 15, after: 0, line: 230 })] : [])];
  return [
    ...moduleHead('08', 'Módulo 8 · Innovación Organizacional y Excelencia', 'Lo que sabe el fundador, al alcance de cada gerente'),
    ficha([['Qué entregamos', 'Asistente con las prácticas del fundador, mejora continua y certificación'], ['Objetivo al que apoya', '{OE-2}{OE-3}'], ['A quién va dirigido', 'Gerentes, que lo usan y lo alimentan'], ['Cómo sabremos que funciona', 'Gerente nuevo en su meta en 90 días (antes 120)']]),
    h2('Iniciativa de innovación: «Copiloto Modelo Ochoa»'),
    p('Lo más valioso de Autos Ochoa son **25 años de experiencia del fundador**: cómo valuar un auto, cómo negociar cuando el cliente deja el suyo a cuenta, qué financiera aprueba a qué cliente. Con cinco sucursales, él no puede estar en todas. La idea es **poner ese conocimiento dentro del CRM** para que cualquier gerente lo consulte, y mejorarlo cada dos semanas con lo que aprendan las sucursales.'),
    cards(steps.map(([n, t, d]) => ({ kids: [p(n, { font: SERIF, bold: true, size: 28, color: C.accent, after: 0 }), p(t, { bold: true, size: SZ.small, after: 10 }), small(d, { size: 15, after: 0, line: 230 })] })), { gap: 100, margins: { top: 60, bottom: 60, left: 100, right: 80 } }),
    gap(),
    callout('Reglas de uso', bullets(['**Un comité al mes:** Dirección, sistemas, RH y un gerente.', 'Los datos de clientes se usan sin nombres y con el aviso de privacidad al día.', 'La IA sugiere, el gerente decide: en precios y créditos siempre revisa una persona.', 'Costo: ~$24,000 al año por el uso de la IA.'], { size: SZ.small, after: 20 })),
    h3('Cómo sabremos que funciona'),
    dataTable([50, 25, 25], ['Qué medimos', 'Meta', 'Cada cuándo'], [
      ['Días para que un gerente nuevo llegue a su meta', '90 o menos', 'Cada generación'],
      ['Dudas que el asistente resuelve solo', '70% o más', 'Cada mes'],
      ['Ideas de mejora puestas en práctica', '2 por sucursal', 'Cada 3 meses'],
      ['Llamadas al fundador por dudas del día a día', '−50%', 'Cada 3 meses'],
    ], { pad: 16 }),
    h2('Certificación recomendada'),
    dataTable([24, 24, 9, 16, 27], ['Certificación', 'Qué reconoce', 'Módulos', 'Tiempo · costo', 'Por qué nos sirve'], [
      [{ c: certName('Great Place to Work®', '', 'Año 1') }, 'Que el equipo confía y está a gusto (encuesta Trust Index©)', '2 · 3 · 5', '8–10 semanas · ~$45,000 estimado', 'Es un sello que ven los candidatos y mide lo que mejora el Módulo 2 (OE-3)'],
      [{ c: certName('NMX-R-025', 'Igualdad Laboral y No Discriminación', 'Año 2') }, 'Igualdad y no discriminación, con auditoría externa', '5 · 3', '6–9 meses · auditoría', 'Las 3 políticas del Módulo 5 ya siguen sus requisitos (OE-3)'],
      [{ c: certName('ISO 9001', 'Calidad', 'Después', C.muted) }, 'Procesos ordenados y medidos', '7 · 8', '9–12 meses · consultoría', 'Útil si pasamos de 5 sucursales; hoy sería adelantarnos'],
    ], { pad: 26 }),
    gap(),
    callout('Nuestra recomendación', '**Certificarnos como Great Place to Work® en agosto de 2027**, al cerrar el año de expansión. Es posible para una empresa de 34 personas y usa el mismo tipo de encuesta que ya haremos. El umbral de referencia es de alrededor de 65% de experiencias positivas (lo confirmaremos al inscribirnos); nuestra meta interna es 70%. Pasos: medir en mayo, inscribirnos en julio, encuesta en agosto y resultado en septiembre.'),
    gap(),
    why('Sin una forma de aprender entre todos, cinco sucursales terminan trabajando de cinco maneras distintas. Un sello externo confirma que lo que decimos es cierto.', 'Si un gerente llega a su meta en 90 días en lugar de 120, cada sucursal gana ~$200,000 antes (OE-2). El sello ayuda a atraer y retener (OE-3).'),
  ];
}

function tablero(charts) {
  const key = (fill, text) => [new TextRun({ text: '■ ', color: fill, size: 16 }), new TextRun({ text: text + '     ', size: 15, color: C.ink2 })];
  return [
    h1('Cómo vamos a medir el avance'),
    p('Todos los números salen de lo que ya registramos (CRM, nómina, Classroom y la encuesta) y se juntan en un solo tablero que Dirección revisa cada mes.'),
    dataTable([12, 31, 10, 11, 14, 9, 13], ['Área', 'Qué medimos', 'Hoy', 'Meta', 'Cada cuándo', 'Objetivo', 'Responsable'], [
      [{ c: '**Capacitación**', rowSpan: 3 }, 'Gerentes certificados 60 días antes de abrir', '0%', '100%', 'Cada apertura', '{OE-1}', 'RH'],
      ['Días para que la sucursal venda 10 autos al mes', 'No se mide', '120 o menos', 'Día 120', '{OE-2}', 'Dir. Comercial'],
      ['Horas de formación por persona', '0', '40 al año', 'Cada año', '{OE-2}', 'RH'],
      [{ c: '**Clima**', rowSpan: 3 }, 'Clima favorable', '54%*', '70% o más', 'Cada 6 meses', '{OE-3}', 'RH'],
      ['¿Recomendarías trabajar aquí? (eNPS)', '+19*', '+30 o más', 'Cada 6 meses', '{OE-3}', 'RH'],
      ['Faltas no programadas', 'No se mide', '2% o menos', 'Cada mes', '{OE-3}', 'Gerentes'],
      [{ c: '**Desempeño**', rowSpan: 3 }, 'Gerentes en su meta a los 120 días', 'No se mide', '80% o más', 'Día 120', '{OE-2}', 'Dir. Comercial'],
      ['Evaluaciones hechas a tiempo', '0%', '100%', 'Cada 6 meses', '{OE-2}', 'RH'],
      ['Gerentes con nivel 3 o más en las 6 competencias', 'No se mide', '80% o más', 'Agosto 2027', '{OE-2}', 'Dir. Comercial'],
      [{ c: '**Permanencia**', rowSpan: 3 }, 'Gerentes que se van sin que queramos', 'No se mide', '10% o menos', 'Cada año', '{OE-3}', 'Dir. General'],
      ['Nuevos que siguen al año', 'No se mide', '85% o más', 'Por generación', '{OE-3}', 'RH'],
      ['Contratados por recomendación o canales propios', '0%', '40% o más', 'Cada año', '{OE-1}', 'RH y Marketing'],
      ['**Sucesión**', 'Puestos clave con reemplazo listo', '0%', '60%', 'Cada 6 meses', '{OE-1}', 'Dirección y RH'],
      ['**Inclusión**', 'Diferencia de sueldo en el mismo puesto', 'No se mide', '5% o menos', 'Cada 6 meses', '{OE-3}', 'Dir. General'],
      ['**Agilidad**', 'Días para cubrir una vacante de gerente', 'No se mide', '45 o menos', 'Cada vacante', '{OE-1}', 'RH'],
      ['**Innovación**', 'Dudas que el asistente resuelve solo', '—', '70% o más', 'Cada mes', '{OE-2}', 'Sistemas'],
    ], { pad: 22, size: 16 }),
    small('* Valores de la encuesta simulada de noviembre de 2026 (Módulo 2).', { before: 40 }),
    h2('Calendario de 12 meses · con las mismas fases de la Parte 1'),
    img(charts.gantt, CW / PX),
    new Paragraph({ spacing: { before: 20, after: 0 }, children: [...key('DBE5F3', 'Preparación'), ...key(C.navy, 'Trabajo'), ...key(C.accent, 'Fecha clave'), new TextRun({ text: '«Sueldos» = revisión semestral de diferencias de sueldo · «30-90» = evaluaciones de los gerentes nuevos', size: 15, color: C.ink2 })] }),
  ];
}

function cambios() {
  const people = [
    ['Si eres gerente de sucursal', ['Tendrás un padrino y la Academia desde antes de abrir tu sucursal.', 'Verás tus números todos los días en un tablero.', 'Sabrás cómo te evaluamos: a los 30, 90 y 120 días, y luego cada 6 meses.', 'Tendrás una ruta escrita para llegar a gerente regional.', 'Opinarás en precios y campañas en el consejo mensual de gerentes.']],
    ['Si eres asesor comercial', ['Sabrás exactamente qué clientes te tocan y cuáles son tus metas.', 'Trabajarás máximo 2 domingos al mes, con el rol publicado un mes antes.', 'Nadie te pedirá contestar fuera de tu guardia: la IA atiende y agenda.', 'Tu comisión aparecerá completa en tu recibo de nómina.', 'Puedes llegar a gerente en 24 meses; las vacantes se publican primero para ti.']],
    ['Si eres coordinador o trabajas en oficinas', ['Recibirás capacitación en expedientes y crédito.', 'Tendrás un día de trabajo en casa a la semana (oficinas).', 'Tu sueldo estará en la misma tabla de bandas que el de todos.', 'Tu área opinará en la evaluación de los gerentes.']],
    ['Si eres de la familia Ochoa', ['Tu puesto estará en la banda de sueldo que le corresponde, como el de cualquiera.', 'Los ascensos pasarán por el mismo comité, con alguien que no es de la familia.', 'La experiencia del fundador se volverá escuela para todos con las charlas grabadas.']],
  ];
  const pc = ([t, items]) => ({ kids: [p(t, { font: SERIF, bold: true, size: 21, color: C.navy, after: 60 }), ...bullets(items, { size: SZ.body, after: 30 })] });
  const next = [['Octubre 2026', 'Poner la base', ['Publicar bandas de sueldo y las 3 políticas', 'Arrancar los ciclos de 2 semanas', 'Elegir el sistema de nómina con checador']], ['Noviembre 2026', 'Escuchar y preparar', ['Aplicar la encuesta «Pulso Ochoa»', 'Checador en todas las sedes', 'Abrir el Nivel 1 de la Academia']], ['Diciembre 2026', 'Empezar a mejorar', ['Arranca el plan de mejora del clima', 'Formulario de evaluación 360° listo', 'Primeras charlas grabadas con el fundador']]];
  const faq = [['¿La evaluación es para correr gente?', 'No. Es para saber en qué apoyar a cada quien. Si alguien no llega, primero tiene un plan de mejora de 90 días con su padrino.'], ['¿Alguien va a ver lo que conteste en la encuesta?', 'No. Es anónima y solo mostramos grupos de 3 personas o más.'], ['¿Me van a bajar el sueldo con la jornada de 46 horas?', 'No. La ley lo prohíbe y nosotros tampoco lo haríamos.'], ['¿Los familiares van a tener privilegios?', 'No. Todos estamos en la misma tabla de sueldos y los ascensos pasan por el mismo comité.'], ['No tengo experiencia en autos, ¿puedo entrar o crecer?', 'Sí. Hay 3 meses de capacitación pagada y las vacantes se publican primero para el equipo.'], ['¿La inteligencia artificial me va a reemplazar?', 'No. Te ayuda a responder y a no olvidar clientes; las decisiones las tomas tú.']];
  const [fl, fr] = splitW([1, 1], 300);
  const faqCell = ([q, a], w) => cell(w, [p(q, { bold: true, after: 10 }), p(a, { color: C.ink2, after: 0, line: 245 })], { borders: { ...noCell, left: ln(C.navy, 18) }, margins: { top: 20, bottom: 20, left: 150, right: 60 } });
  const faqRows = [];
  for (let i = 0; i < faq.length; i += 2) {
    if (i) faqRows.push(new TableRow({ height: { value: 110, rule: HeightRule.ATLEAST }, children: [spacerCell(fl), spacerCell(300), spacerCell(fr)] }));
    faqRows.push(new TableRow({ cantSplit: true, children: [faqCell(faq[i], fl), spacerCell(300), faqCell(faq[i + 1], fr)] }));
  }
  return [
    h1('Qué cambia para cada quien'),
    p('Este plan no es solo para Recursos Humanos. Esto es lo que cada persona va a notar en su día a día a partir de octubre.', { size: 20, color: C.ink2, after: 120 }),
    cards([pc(people[0]), pc(people[1])], { gap: 160, split: true, margins: { top: 100, bottom: 100, left: 140, right: 120 } }),
    gap(),
    cards([pc(people[2]), pc(people[3])], { gap: 160, split: true, margins: { top: 100, bottom: 100, left: 140, right: 120 } }),
    h2('Próximos 90 días'),
    cards(next.map(([s, b, items]) => ({ kids: [p(s, { size: 13, caps: true, bold: true, color: C.muted, spacing: 8, after: 10 }), p(b, { bold: true, after: 30 }), ...bullets(items, { size: SZ.small, after: 20 })], borders: { ...noCell, top: ln(C.accent, 18) } })), { gap: 160, margins: { top: 60, bottom: 20, left: 0, right: 60 } }),
    h2('Preguntas que seguramente vas a tener'),
    tbl([fl, 300, fr], faqRows),
    gap(),
    callout('Lo que necesitamos de Dirección', '(1) Aprobar los $241,000 para el año de expansión; (2) publicar en octubre las bandas de sueldo y las tres políticas, antes de la siguiente contratación; (3) nombrar a la Coordinación de RH responsable del plan; (4) agendar entre noviembre y enero las 12 charlas grabadas con el fundador.'),
  ];
}

function conclusiones() {
  const ref = items => bullets(items, { size: 15, after: 15 });
  return [
    h1('Conclusiones'),
    ...[
      '**1. Todo gira alrededor del Gerente de Sucursal.** Cada acción sirve para una de tres cosas: tenerlo listo antes de abrir (OE-1), que su sucursal sea rentable en 4 meses (OE-2) y que se quede (OE-3). Ninguna acción está aquí por moda.',
      '**2. Mejorar el clima es mejorar las ventas.** Lo más bajo de la encuesta —no saber quién atiende a cada cliente— es la misma falla que hoy deja caer al 97% de los prospectos.',
      '**3. La tecnología ya la tenemos.** El CRM con inteligencia artificial, Classroom, Looker Studio y WhatsApp bastan. Trabajar en ciclos de dos semanas permite que una sola persona de RH saque adelante 26 contrataciones y 4 aperturas.',
      '**4. Las mismas reglas para todos.** Sueldos claros, contratar por lo que sabes hacer y tiempo para la familia son la mejor forma de atraer y retener talento en una empresa familiar que crece.',
      '**5. Lo que sabe el fundador, para todos.** El asistente con sus prácticas y la certificación Great Place to Work® convierten 25 años de experiencia en algo que cualquiera puede consultar, por menos de la mitad de lo que cuesta perder a un solo gerente.',
    ].map(t => p(t, { after: 90, line: 260 })),
    h2('Qué podría salir mal y cómo lo evitamos'),
    dataTable([31, 11, 49, 9], ['Riesgo', 'Probabilidad', 'Qué haremos', 'Módulos'], [
      ['Que RH no alcance con todo el trabajo', 'Alta', 'Priorizar por objetivo en ciclos de 2 semanas; sistemas y marketing ayudan; se avanza por apertura', '7'],
      ['Resistencia a publicar sueldos y aplicar las mismas reglas a la familia', 'Media', 'Decisión de Dirección desde octubre; el fundador lo explica en la «Junta Ochoa»', '2 · 5'],
      ['Que la gente no conteste con sinceridad por miedo', 'Media', 'Anonimato real (grupos de 3 o más) y resultados compartidos en 30 días', '2 · 4'],
      ['Que el asistente de IA se equivoque o se usen mal datos de clientes', 'Media', 'Prueba de 90 días, revisión humana en precios y créditos, datos sin nombres', '8'],
      ['Que se gane menos por auto y solo se abran 3 sucursales', 'Media', 'El presupuesto se ajusta por apertura; las acciones sirven igual con menos sucursales', 'Todos'],
    ], { pad: 26 }),
    h2('Fuentes y referencias', { before: 200 }),
    h3('De Autos Ochoa', { before: 0 }), ...ref(['Manual de Operaciones de RH y presentación del Sistema Integral de Gestión de Talento (SIGT), Parte 1, agosto de 2026.', 'Diagnóstico Estratégico y Cuestionario de Diagnóstico (ERP, CRM, ventas 2020–2026), agosto de 2026.', 'Plan de Crecimiento a 6 Meses, septiembre 2026 – febrero 2027.']),
        h3('Leyes y normas'), ...ref(['Ley Federal del Trabajo: arts. 2, 3, 39-A a 39-D, 59, 66 a 71, 86, 132 fr. XXXIV, 133, 153-A y siguientes (capacitación y DC-3), 170 y 330-A.', 'Decreto de reducción de la jornada laboral, DOF 1 de mayo de 2026.', 'NOM-035-STPS-2018, factores de riesgo psicosocial (Guía de Referencia II).', 'NMX-R-025-SCFI-2015, Igualdad Laboral y No Discriminación.', 'Ley Federal para Prevenir y Eliminar la Discriminación; Ley Federal de Protección de Datos Personales en Posesión de los Particulares.']),
    h3('Métodos'), ...ref(['Kirkpatrick, D. y Kirkpatrick, J. (2006). __Evaluating Training Programs: The Four Levels__ (3.ª ed.). Berrett-Koehler.', 'Lombardo, M. y Eichinger, R. (1996). __The Career Architect Development Planner__ (modelo 70-20-10). Lominger.', 'Reichheld, F. (2003). The One Number You Need to Grow. __Harvard Business Review__ (base del eNPS).', 'Schwaber, K. y Sutherland, J. (2020). __La Guía de Scrum__. Scrum.org.', 'Anderson, D. (2010). __Kanban__. Blue Hole Press.', 'Matriz de desempeño y potencial (9-Box), derivada de la matriz GE-McKinsey.', 'Global Reporting Initiative: GRI 401, 404, 405 y 406.', 'Great Place to Work® — Trust Index©. Naciones Unidas — Agenda 2030, ODS 3, 5, 8 y 10.', 'Deming, W. E. — ciclo PDCA; Imai, M. (1986). __Kaizen__. McGraw-Hill.']),
  ];
}

// ---------- Documento ----------
const charts = await captureCharts();
const doc = new Document({
  creator: 'Autos Ochoa · Recursos Humanos',
  title: 'Autos Ochoa · Plan de talento 2026–2027',
  description: 'Reto Final de Administración del Talento Humano: ocho módulos de desarrollo organizacional para el plan de expansión «Ochoa 5».',
  styles: {
    default: { document: { run: { font: SANS, size: SZ.body, color: C.ink }, paragraph: { spacing: { after: 80, lineRule: LineRuleType.AUTO, line: 252 } } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: SERIF, size: 34, bold: false, color: C.ink }, paragraph: { spacing: { before: 0, after: 120 }, keepNext: true, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: SERIF, size: 25, bold: true, color: C.navy }, paragraph: { spacing: { before: 140, after: 60 }, keepNext: true, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: SANS, size: 15, bold: true, allCaps: true, color: C.navy, characterSpacing: 16 }, paragraph: { spacing: { before: 120, after: 50 }, keepNext: true, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [{
      reference: 'bul',
      levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 240, hanging: 180 } }, run: { color: C.accent } } }],
    }],
  },
  sections: [
    cover(),
    section('En pocas palabras', resumen()),
    section('Módulo 1 · Capacitación y Desarrollo', modulo1()),
    section('Módulo 2 · Clima y Experiencia', modulo2(charts)),
    section('Módulo 3 · Marca Empleadora y Permanencia', modulo3()),
    section('Módulo 4 · Desempeño y Competencias', modulo4(charts)),
    section('Módulo 5 · Responsabilidad Social', modulo5()),
    section('Módulo 6 · Sucesión y Talento Clave', modulo6()),
    section('Módulo 7 · Digitalización y Agilidad', modulo7()),
    section('Módulo 8 · Innovación y Excelencia', modulo8()),
    section('Cómo mediremos y cuándo', tablero(charts)),
    section('Qué cambia para cada quien', cambios()),
    section('Conclusiones y fuentes', conclusiones()),
  ],
});

fs.mkdirSync(path.dirname(outPath), { recursive: true });
fs.writeFileSync(outPath, await Packer.toBuffer(doc));
console.log('Word listo:', path.relative(process.cwd(), outPath));
