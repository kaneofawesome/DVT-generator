#!/usr/bin/env node
// DVT generator — compact table format.
// Usage:
//   node build_dvt.js [data.json]
//   default: data.json = ./dvt_data.json
//
// Output filename is derived from the data file's meta block:
//   meta.outputFilename if set, else meta.headerLine (sanitized) + '.docx',
//   else <data-basename>.docx. The filename tracks the in-doc revision string
//   so a single edit in the meta bumps both together.
//
// All product-specific text (title, cover rows, header line, external-reference
// pages, matrix headers, appendix titles) lives in the `meta` block of the data
// JSON. Copy an existing data file and edit its meta + content to spin up a
// new DVT with the same look and feel.

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip'); // transitive dep of docx
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, PageOrientation,
  Bookmark, InternalHyperlink,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageBreak, PageNumber,
  SimpleField,
} = require('docx');

// ============ CLI ============
const INPUT = process.argv[2] || 'dvt_data.json';
if (!fs.existsSync(INPUT)) {
  console.error(`Input not found: ${INPUT}`);
  console.error('Usage: node build_dvt.js [data.json]');
  process.exit(1);
}
let data = JSON.parse(fs.readFileSync(INPUT, 'utf-8'));
data.meta = data.meta || {};

// ---------- Today's date default ----------
if (!data.meta.date) {
  const t = new Date();
  data.meta.date = `${t.getMonth() + 1}/${t.getDate()}/${t.getFullYear()}`;
}

// ---------- Section letter abstraction ----------
// External refs and sections claim letters in array order. Each section can
// have a `code` (stable symbolic name) and a `name` (sub-title). The script
// assigns `.letter` per section based on position. `id` is set to the letter
// for back-compat with bookmark names. `title` defaults to "Section X — name"
// if not explicitly set. Steps may use legacy `ref: "X.Y.Z"` or new
// `section: <code>` + `subref: "Y.Z"`; the script computes `ref` either way.
let letterCode = 'A'.charCodeAt(0);
const codeToLetter = {};
for (const ext of data.meta.externalRefs || []) {
  if (ext.sectionLetter) {
    letterCode = Math.max(letterCode, ext.sectionLetter.charCodeAt(0) + 1);
  } else {
    ext.sectionLetter = String.fromCharCode(letterCode++);
  }
  if (ext.code) codeToLetter[ext.code] = ext.sectionLetter;
  codeToLetter[ext.sectionLetter] = ext.sectionLetter; // letter resolves to itself
}
for (const sec of data.sections || []) {
  // If `id` is already a letter, honor it (back-compat). Otherwise assign by position.
  if (sec.id && /^[A-Z]$/.test(sec.id)) {
    sec.letter = sec.id;
    letterCode = Math.max(letterCode, sec.id.charCodeAt(0) + 1);
  } else if (!sec.letter) {
    sec.letter = String.fromCharCode(letterCode++);
  }
  sec.id = sec.letter; // bookmark target sec_<letter>
  if (sec.code) codeToLetter[sec.code] = sec.letter;
  codeToLetter[sec.letter] = sec.letter; // letter resolves to itself
  if (!sec.title) {
    sec.title = sec.name ? `Section ${sec.letter} — ${sec.name}` : `Section ${sec.letter}`;
  }
}
// Compute step refs from section + subref when legacy `ref` is absent
for (const step of data.steps || []) {
  if (step.ref) continue;
  if (step.section && step.subref != null) {
    const letter = codeToLetter[step.section];
    if (!letter) {
      console.error(`Step references unknown section code "${step.section}":`, step);
      process.exit(1);
    }
    step.ref = `${letter}.${step.subref}`;
  } else {
    console.error('Step is missing `ref` or `section`+`subref`:', step);
    process.exit(1);
  }
}

// ---------- Cross-reference expansion ----------
// In any string, [code] expands to the section letter (e.g. [fcat] → "D") and
// [code.subref] expands to "D.subref" (e.g. [fcat.1.1] → "D.1.1"). Letter
// codes work too: [D.1.1] → "D.1.1" — useful when migrating gradually.
// Unknown codes are left as-is so typos surface in the rendered doc.
function expandCrossRefs(s, map) {
  return s.replace(/\[([A-Za-z_]\w*)((?:\.[\w.-]+)?)\]/g, (m, code, suffix) => {
    if (!(code in map)) return m;
    return map[code] + (suffix || '');
  });
}
function deepCrossRef(obj, map) {
  if (typeof obj === 'string') return expandCrossRefs(obj, map);
  if (Array.isArray(obj)) return obj.map((x) => deepCrossRef(x, map));
  if (obj && typeof obj === 'object') {
    const out = {};
    for (const k of Object.keys(obj)) out[k] = deepCrossRef(obj[k], map);
    return out;
  }
  return obj;
}
data = deepCrossRef(data, codeToLetter);

// ---------- Meta interpolation ({key} → meta.key) ----------
function interpolateStr(s, vars) {
  return s.replace(/\{(\w+)\}/g, (m, k) => (k in vars ? String(vars[k]) : m));
}
function deepInterpolate(obj, vars) {
  if (typeof obj === 'string') return interpolateStr(obj, vars);
  if (Array.isArray(obj)) return obj.map((x) => deepInterpolate(x, vars));
  if (obj && typeof obj === 'object') {
    const out = {};
    for (const k of Object.keys(obj)) out[k] = deepInterpolate(obj[k], vars);
    return out;
  }
  return obj;
}
const vars = {};
for (const [k, v] of Object.entries(data.meta)) {
  if (typeof v === 'string' || typeof v === 'number') vars[k] = v;
}
data = deepInterpolate(data, vars);
const meta = data.meta || {};

// ---------- Lint ----------
// Warnings (non-fatal) and errors (fatal). All check the post-expansion data.
const allRefs = new Set((data.steps || []).map((s) => s.ref));
const lintWarnings = [];
const lintErrors = [];

function lintString(s, where) {
  // Unresolved [code...] tokens — typo or missing section code
  for (const m of s.matchAll(/\[([A-Za-z_]\w*)(?:\.[\w.-]+)?\]/g)) {
    lintWarnings.push(`Unresolved cross-ref ${m[0]} in ${where}`);
  }
  // Letter-style refs like "D.1.1" that don't match a known step
  for (const m of s.matchAll(/\b([A-Z])\.(\d+(?:\.\d+)*)\b/g)) {
    const ref = `${m[1]}.${m[2]}`;
    if (!allRefs.has(ref)) {
      lintWarnings.push(`Unknown ref "${ref}" in ${where} — stale cross-ref?`);
    }
  }
  // Unresolved {placeholder} (after meta interpolation any left over is a typo)
  for (const m of s.matchAll(/\{(\w+)\}/g)) {
    lintWarnings.push(`Unresolved placeholder ${m[0]} in ${where}`);
  }
}
function walkLint(obj, path) {
  if (typeof obj === 'string') lintString(obj, path);
  else if (Array.isArray(obj)) obj.forEach((x, i) => walkLint(x, `${path}[${i}]`));
  else if (obj && typeof obj === 'object') for (const k of Object.keys(obj)) walkLint(obj[k], `${path}.${k}`);
}
walkLint(data, 'data');

// Duplicate step refs — fatal
{
  const seen = new Set();
  for (const step of data.steps || []) {
    if (seen.has(step.ref)) lintErrors.push(`Duplicate step ref: ${step.ref}`);
    seen.add(step.ref);
  }
}

// Required meta — warn
for (const k of ['productName', 'headerLine']) {
  if (!meta[k]) lintWarnings.push(`meta.${k} is missing or empty`);
}

if (lintWarnings.length) {
  console.warn(`\nLint warnings (${lintWarnings.length}):`);
  for (const w of lintWarnings) console.warn('  ' + w);
  console.warn('');
}
if (lintErrors.length) {
  console.error(`Lint errors (${lintErrors.length}):`);
  for (const e of lintErrors) console.error('  ' + e);
  process.exit(1);
}

// Output filename comes from meta: meta.outputFilename, else derived from
// meta.headerLine (sanitized), else <data-basename>.docx.
// Output directory comes from meta.outputDir (absolute or resolved against
// the directory containing the data JSON, so paths like "../output" are
// portable regardless of where node is invoked from). Created if missing.
function sanitizeForFilename(s) {
  return String(s)
    .replace(/[—–]/g, '-')              // em/en-dash → hyphen
    .replace(/[\/\\:*?"<>|]/g, '_')     // reserved filename chars → underscore
    .replace(/\s+/g, ' ')               // collapse runs of whitespace
    .trim();
}
function resolveOutName() {
  if (meta.outputFilename) return meta.outputFilename;
  if (meta.headerLine) return sanitizeForFilename(meta.headerLine) + '.docx';
  return path.basename(INPUT, '.json') + '.docx';
}
function resolveOutDir() {
  // Default: alongside the data file (so the docx appears next to its source
  // regardless of where node was invoked from). Override with meta.outputDir,
  // which can be absolute or relative-to-the-data-file's directory.
  const inputDir = path.dirname(path.resolve(INPUT));
  if (!meta.outputDir) return inputDir;
  return path.isAbsolute(meta.outputDir) ? meta.outputDir : path.resolve(inputDir, meta.outputDir);
}
const OUT_NAME = resolveOutName();
const OUT_DIR = resolveOutDir();
const OUT = path.join(OUT_DIR, OUT_NAME);
if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });

// ============ Page / layout constants ============
const PAGE = { width: 12240, height: 15840, margin: 1008 }; // US Letter, 0.7"
const CONTENT_W = PAGE.width - 2 * PAGE.margin; // 10224 DXA

const COL_REF = 720, COL_HEADER = 1700, COL_BRIEF = 3200, COL_CRIT = 2400, COL_MEAS = 1500, COL_PF = 704;
const ROW_COLS = [COL_REF, COL_HEADER, COL_BRIEF, COL_CRIT, COL_MEAS, COL_PF];

const border = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 60, bottom: 60, left: 100, right: 100 };

const HEADER_FILL = 'D9E2F3';
const ALT_FILL = 'F5F7FB';
const SECTION_FILL = 'B4C7E7';

// ============ Bookmark IDs ============
const refToBkm = (ref) => 'test_' + String(ref).replace(/\./g, '_');
const refToPfBkm = (ref) => 'pf_' + String(ref).replace(/\./g, '_');

// ============ Helpers ============
const cell = (text, widthDxa, opts = {}) => {
  const children = Array.isArray(text)
    ? text
    : [new Paragraph({
        spacing: { before: 20, after: 20 },
        alignment: opts.align || AlignmentType.LEFT,
        children: [new TextRun({ text: String(text == null ? '' : text), bold: !!opts.bold, size: opts.size || 18 })],
      })];
  return new TableCell({
    borders,
    width: { size: widthDxa, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    children,
  });
};

const headerRow = (labels) => new TableRow({
  tableHeader: true,
  children: labels.map((label, i) => cell(label, ROW_COLS[i], { bold: true, fill: HEADER_FILL, align: AlignmentType.CENTER })),
});

// One test row. Bookmark wraps the Ref cell content (TOC hyperlink target +
// PAGEREF source) and the P/F cell ☐ glyph (REF mirror source).
const stepRow = (s, alt) => {
  const fill = alt ? ALT_FILL : undefined;
  const refPara = new Paragraph({
    spacing: { before: 20, after: 20 },
    alignment: AlignmentType.LEFT,
    children: [new Bookmark({
      id: refToBkm(s.ref),
      children: [new TextRun({ text: s.ref, bold: true, size: 18 })],
    })],
  });
  const pfPara = new Paragraph({
    spacing: { before: 20, after: 20 },
    alignment: AlignmentType.CENTER,
    children: [new Bookmark({
      id: refToPfBkm(s.ref),
      children: [new TextRun({ text: '☐', size: 22 })],
    })],
  });
  return new TableRow({
    cantSplit: true,
    children: [
      new TableCell({ borders, width: { size: COL_REF, type: WidthType.DXA }, shading: fill ? { fill, type: ShadingType.CLEAR } : undefined, margins: cellMargins, children: [refPara] }),
      cell(s.header, COL_HEADER, { fill, bold: true }),
      cell(s.brief, COL_BRIEF, { fill }),
      cell(s.criteria, COL_CRIT, { fill }),
      cell('', COL_MEAS, { fill }),
      new TableCell({ borders, width: { size: COL_PF, type: WidthType.DXA }, shading: fill ? { fill, type: ShadingType.CLEAR } : undefined, margins: cellMargins, children: [pfPara] }),
    ],
  });
};

const sectionTable = (steps) => new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: ROW_COLS,
  rows: [
    headerRow(['Ref', 'Header', 'Brief', 'Criteria', 'Measure', 'P/F']),
    ...steps.map((s, i) => stepRow(s, i % 2 === 1)),
  ],
});

// ============ Cover sheet ============
const COVER_L = 3600;
const COVER_R = CONTENT_W - COVER_L;
const coverCell = (label, value, isLabel) => new TableCell({
  borders,
  width: { size: isLabel ? COVER_L : COVER_R, type: WidthType.DXA },
  shading: isLabel ? { fill: 'F2F2F2', type: ShadingType.CLEAR } : undefined,
  margins: { top: 100, bottom: 100, left: 140, right: 140 },
  children: [new Paragraph({ children: [new TextRun({ text: isLabel ? label : value, bold: isLabel, size: 20 })] })],
});
const coverRow = (label, value) => new TableRow({ children: [coverCell(label, '', true), coverCell('', value || '', false)] });

const coverTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [COVER_L, COVER_R],
  rows: (meta.coverRows || []).map((r) => coverRow(r.label, r.value)),
});

const coverPara = [];
if (meta.title) coverPara.push(new Paragraph({ spacing: { before: 100, after: 100 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: meta.title, bold: true, size: 32 })] }));
if (meta.productName) coverPara.push(new Paragraph({ spacing: { before: 100, after: 60 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: meta.productName, bold: true, size: 28 })] }));
if (meta.subtitle) coverPara.push(new Paragraph({ spacing: { before: 60, after: 60 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: meta.subtitle, size: 22 })] }));
if (meta.descriptor) coverPara.push(new Paragraph({ spacing: { before: 60, after: 240 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: meta.descriptor, italics: true, size: 22 })] }));
coverPara.push(coverTable);

// ============ Known-issues, equipment, release-gate tables ============
const KI_W = [3800, 1500, 2300, 2624];
const knownIssuesTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: KI_W,
  rows: [
    new TableRow({ tableHeader: true, children: ['Defect', 'Jira', 'Suite test', 'DVT step'].map((label, i) => cell(label, KI_W[i], { bold: true, fill: HEADER_FILL })) }),
    ...(data.knownIssues || []).map((ki, i) => new TableRow({
      cantSplit: true,
      children: [
        cell(ki.defect, KI_W[0], { fill: i % 2 ? ALT_FILL : undefined }),
        cell(ki.jira, KI_W[1], { fill: i % 2 ? ALT_FILL : undefined }),
        cell(ki.suite, KI_W[2], { fill: i % 2 ? ALT_FILL : undefined }),
        cell(ki.dvt, KI_W[3], { fill: i % 2 ? ALT_FILL : undefined }),
      ],
    })),
  ],
});

const EQ_W = [3000, 4200, 3024];
const equipmentTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: EQ_W,
  rows: [
    new TableRow({ tableHeader: true, children: ['Equipment', 'Action / use', 'Spec'].map((label, i) => cell(label, EQ_W[i], { bold: true, fill: HEADER_FILL })) }),
    ...(data.equipment || []).map((eq, i) => new TableRow({
      cantSplit: true,
      children: [
        cell(eq.name, EQ_W[0], { fill: i % 2 ? ALT_FILL : undefined, bold: true }),
        cell(eq.action, EQ_W[1], { fill: i % 2 ? ALT_FILL : undefined }),
        cell(eq.spec, EQ_W[2], { fill: i % 2 ? ALT_FILL : undefined }),
      ],
    })),
  ],
});

const RG_CAT_W = [800, 4200, 5224];
const releaseCatTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: RG_CAT_W,
  rows: [
    new TableRow({ tableHeader: true, children: ['Cat', 'Change type', 'Required sections'].map((label, i) => cell(label, RG_CAT_W[i], { bold: true, fill: HEADER_FILL })) }),
    ...((data.releaseGate && data.releaseGate.categories) || []).map((c, i) => new TableRow({
      cantSplit: true,
      children: [
        cell(c.id, RG_CAT_W[0], { fill: i % 2 ? ALT_FILL : undefined, bold: true, align: AlignmentType.CENTER }),
        cell(c.name, RG_CAT_W[1], { fill: i % 2 ? ALT_FILL : undefined }),
        cell(c.required, RG_CAT_W[2], { fill: i % 2 ? ALT_FILL : undefined }),
      ],
    })),
  ],
});

// Generic compatibility matrix — headers + per-row cells driven from meta + data
const matrixHeaders = (meta.appendixB && meta.appendixB.matrixHeaders) || [];
const defaultMatrixWidths = () => {
  if (matrixHeaders.length === 0) return [];
  const w = Math.floor(CONTENT_W / matrixHeaders.length);
  const arr = Array(matrixHeaders.length).fill(w);
  arr[arr.length - 1] = CONTENT_W - w * (matrixHeaders.length - 1);
  return arr;
};
const RG_MATRIX_W = (meta.appendixB && meta.appendixB.matrixColumnWidths) || defaultMatrixWidths();
const matrixRows = ((data.releaseGate && data.releaseGate.matrix) || []);
const releaseMatrixTable = matrixHeaders.length > 0 ? new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: RG_MATRIX_W,
  rows: [
    new TableRow({ tableHeader: true, children: matrixHeaders.map((label, i) => cell(label, RG_MATRIX_W[i], { bold: true, fill: HEADER_FILL, align: AlignmentType.CENTER })) }),
    ...matrixRows.map((r, i) => new TableRow({
      cantSplit: true,
      children: (r.cells || []).map((c, j) => cell(c, RG_MATRIX_W[j] || RG_MATRIX_W[RG_MATRIX_W.length - 1], { fill: i % 2 ? ALT_FILL : undefined, bold: j === 0, align: (j > 0 && j < matrixHeaders.length - 1) ? AlignmentType.CENTER : AlignmentType.LEFT })),
    })),
  ],
}) : null;

// ============ TOC table ============
const TOC_W = [1000, 7700, 800, 724];
function tocSectionDividerRow(text, anchorId) {
  return new TableRow({
    cantSplit: true,
    children: [new TableCell({
      borders,
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnSpan: 4,
      shading: { fill: SECTION_FILL, type: ShadingType.CLEAR },
      margins: { top: 40, bottom: 40, left: 100, right: 100 },
      children: [new Paragraph({
        spacing: { before: 0, after: 0 },
        children: [new InternalHyperlink({
          anchor: anchorId,
          children: [new TextRun({ text, bold: true, size: 18, color: '1F3864' })],
        })],
      })],
    })],
  });
}
function tocTestRow(step) {
  const linkRun = (text) => new InternalHyperlink({
    anchor: refToBkm(step.ref),
    children: [new TextRun({ text, style: 'Hyperlink', color: '0066CC', size: 16 })],
  });
  return new TableRow({
    cantSplit: true,
    children: [
      new TableCell({ borders, width: { size: TOC_W[0], type: WidthType.DXA }, margins: { top: 20, bottom: 20, left: 80, right: 80 }, children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [linkRun(step.ref)] })] }),
      new TableCell({ borders, width: { size: TOC_W[1], type: WidthType.DXA }, margins: { top: 20, bottom: 20, left: 80, right: 80 }, children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [linkRun(step.header)] })] }),
      new TableCell({ borders, width: { size: TOC_W[2], type: WidthType.DXA }, margins: { top: 20, bottom: 20, left: 80, right: 80 }, children: [new Paragraph({ spacing: { before: 0, after: 0 }, alignment: AlignmentType.CENTER, children: [new SimpleField('PAGEREF ' + refToBkm(step.ref) + ' \\h')] })] }),
      new TableCell({ borders, width: { size: TOC_W[3], type: WidthType.DXA }, margins: { top: 20, bottom: 20, left: 80, right: 80 }, children: [new Paragraph({ spacing: { before: 0, after: 0 }, alignment: AlignmentType.CENTER, children: [new SimpleField('REF ' + refToPfBkm(step.ref) + ' \\h')] })] }),
    ],
  });
}
function tocRefPageRow(sectionLetter, sectionTitle) {
  return new TableRow({
    cantSplit: true,
    children: [
      new TableCell({ borders, width: { size: TOC_W[0], type: WidthType.DXA }, margins: { top: 20, bottom: 20, left: 80, right: 80 }, children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: 'Sec ' + sectionLetter, size: 16 })] })] }),
      new TableCell({ borders, width: { size: TOC_W[1], type: WidthType.DXA }, margins: { top: 20, bottom: 20, left: 80, right: 80 }, children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [new InternalHyperlink({ anchor: 'sec_' + sectionLetter, children: [new TextRun({ text: sectionTitle + ' (external reference)', style: 'Hyperlink', italics: true, color: '0066CC', size: 16 })] })] })] }),
      new TableCell({ borders, width: { size: TOC_W[2], type: WidthType.DXA }, margins: { top: 20, bottom: 20, left: 80, right: 80 }, children: [new Paragraph({ spacing: { before: 0, after: 0 }, alignment: AlignmentType.CENTER, children: [new SimpleField('PAGEREF sec_' + sectionLetter + ' \\h')] })] }),
      new TableCell({ borders, width: { size: TOC_W[3], type: WidthType.DXA }, margins: { top: 20, bottom: 20, left: 80, right: 80 }, children: [new Paragraph({ spacing: { before: 0, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '—', size: 16 })] })] }),
    ],
  });
}

const tocHeaderRow = new TableRow({
  tableHeader: true,
  children: ['Ref', 'Test header', 'Page', 'P/F'].map((label, i) => cell(label, TOC_W[i], { bold: true, fill: HEADER_FILL, align: AlignmentType.CENTER })),
});

const tocRows = [tocHeaderRow];
for (const ref of (meta.externalRefs || [])) {
  tocRows.push(tocRefPageRow(ref.sectionLetter, `Section ${ref.sectionLetter} — ${ref.title}`));
}
for (const section of (data.sections || [])) {
  tocRows.push(tocSectionDividerRow(section.title, 'sec_' + section.id));
  const sectionSteps = (data.steps || []).filter((s) => s.ref.startsWith(section.id + '.'));
  for (const step of sectionSteps) tocRows.push(tocTestRow(step));
}

const tocTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: TOC_W,
  rows: tocRows,
});

// ============ Compose document ============
const children = [];
children.push(...coverPara);

// TOC
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 0, after: 80 }, children: [new TextRun({ text: 'Contents', bold: true, size: 28 })] }));
children.push(new Paragraph({ spacing: { before: 40, after: 120 }, children: [new TextRun({ text: 'Click any row to jump to the test. Page numbers and P/F mirror update on field refresh — Word: Ctrl+A then F9; LibreOffice: Tools → Update → Fields.', italics: true, size: 16 })] }));
children.push(tocTable);

// Known issues
const ki = meta.knownIssues || {};
if ((data.knownIssues || []).length > 0 || ki.heading) {
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 0, after: 120 }, children: [new TextRun({ text: ki.heading || 'Known Issues', bold: true, size: 28 })] }));
  if (ki.intro) children.push(new Paragraph({ spacing: { before: 60, after: 120 }, children: [new TextRun({ text: ki.intro, italics: true, size: 18 })] }));
  children.push(knownIssuesTable);
}

// External-reference pages (Section A, B, ...) — one per entry in meta.externalRefs
const ATT_W = [Math.round(CONTENT_W * 0.55), 0];
ATT_W[1] = CONTENT_W - ATT_W[0];
function externalRefPage(opts) {
  const letter = opts.sectionLetter;
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 0, after: 120 },
    children: [new Bookmark({ id: 'sec_' + letter, children: [new TextRun({ text: `Section ${letter} — ${opts.title} (External Reference)`, bold: true, size: 28 })] })],
  }));
  if (opts.intro) children.push(new Paragraph({ spacing: { before: 60, after: 160 }, children: [new TextRun({ text: opts.intro, italics: true, size: 20 })] }));
  children.push(new Paragraph({ spacing: { before: 200, after: 80 }, children: [new TextRun({ text: 'Reference', bold: true, size: 22 })] }));
  children.push(new Paragraph({
    spacing: { before: 40, after: 40 },
    children: [
      new TextRun({ text: `${opts.title} — latest released revision: `, size: 20 }),
      new TextRun({ text: opts.linkPlaceholder || '<Link>', bold: true, color: '0066CC', size: 20 }),
    ],
  }));
  if (opts.passCriteria) {
    children.push(new Paragraph({ spacing: { before: 200, after: 80 }, children: [new TextRun({ text: `Pass criteria for Section ${letter}`, bold: true, size: 22 })] }));
    children.push(new Paragraph({ spacing: { before: 40, after: 40 }, children: [new TextRun({ text: opts.passCriteria, size: 20 })] }));
  }
  if (opts.attestHeaders && opts.attestHeaders.length === 2) {
    children.push(new Paragraph({ spacing: { before: 200, after: 80 }, children: [new TextRun({ text: 'Tester attestation', bold: true, size: 22 })] }));
    const tbl = new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: ATT_W,
      rows: [
        new TableRow({ tableHeader: true, children: opts.attestHeaders.map((label, i) => cell(label, ATT_W[i], { bold: true, fill: HEADER_FILL })) }),
        new TableRow({ cantSplit: true, children: [cell('', ATT_W[0]), cell('', ATT_W[1])] }),
      ],
    });
    children.push(tbl);
  }
}
for (const ref of (meta.externalRefs || [])) externalRefPage(ref);

// Section tables
for (const section of (data.sections || [])) {
  const sectionSteps = (data.steps || []).filter((s) => s.ref.startsWith(section.id + '.'));
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 0, after: 120 }, children: [new Bookmark({ id: 'sec_' + section.id, children: [new TextRun({ text: section.title, bold: true, size: 28 })] })] }));
  if (section.intro) children.push(new Paragraph({ spacing: { before: 60, after: 160 }, children: [new TextRun({ text: section.intro, italics: true, size: 18 })] }));
  children.push(sectionTable(sectionSteps));
}

// Appendix A — Equipment
const appA = meta.appendixA || {};
if ((data.equipment || []).length > 0) {
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 0, after: 120 }, children: [new TextRun({ text: appA.heading || 'Appendix A — Equipment', bold: true, size: 28 })] }));
  if (appA.intro) children.push(new Paragraph({ spacing: { before: 60, after: 160 }, children: [new TextRun({ text: appA.intro, italics: true, size: 18 })] }));
  children.push(equipmentTable);
}

// Appendix B — Release-gate matrix
const appB = meta.appendixB || {};
if (((data.releaseGate && data.releaseGate.categories) || []).length > 0 || matrixRows.length > 0) {
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 0, after: 120 }, children: [new TextRun({ text: appB.heading || 'Appendix B — Release-Gate Matrix', bold: true, size: 28 })] }));
  if (appB.intro) children.push(new Paragraph({ spacing: { before: 60, after: 120 }, children: [new TextRun({ text: appB.intro, italics: true, size: 18 })] }));
  if (((data.releaseGate && data.releaseGate.categories) || []).length > 0) {
    children.push(new Paragraph({ spacing: { before: 120, after: 80 }, children: [new TextRun({ text: appB.categoriesTitle || 'B.1 Change-severity categories', bold: true, size: 22 })] }));
    children.push(releaseCatTable);
  }
  if (releaseMatrixTable && matrixRows.length > 0) {
    children.push(new Paragraph({ spacing: { before: 200, after: 80 }, children: [new TextRun({ text: appB.matrixTitle || 'B.2 Compatibility matrix', bold: true, size: 22 })] }));
    children.push(releaseMatrixTable);
  }
}

// Appendix C — Elaborations
const appC = meta.appendixC || {};
if ((data.elaborations || []).length > 0) {
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 0, after: 120 }, children: [new TextRun({ text: appC.heading || 'Appendix C — Elaborations and Rationale', bold: true, size: 28 })] }));
  if (appC.intro) children.push(new Paragraph({ spacing: { before: 60, after: 160 }, children: [new TextRun({ text: appC.intro, italics: true, size: 18 })] }));
  for (const e of data.elaborations) {
    children.push(new Paragraph({ spacing: { before: 140, after: 40 }, children: [new TextRun({ text: e.ref, bold: true, size: 20 })] }));
    children.push(new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: e.text, size: 18 })] }));
  }
}

// ============ Document config ============
const doc = new Document({
  creator: meta.creator || 'DVT Generator',
  styles: {
    default: { document: { run: { font: 'Arial', size: 20 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 28, bold: true, font: 'Arial' }, paragraph: { spacing: { before: 120, after: 120 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 22, bold: true, font: 'Arial' }, paragraph: { spacing: { before: 100, after: 100 }, outlineLevel: 1 } },
    ],
  },
  sections: [{
    properties: { page: { size: { width: PAGE.width, height: PAGE.height, orientation: PageOrientation.PORTRAIT }, margin: { top: PAGE.margin, right: PAGE.margin, bottom: PAGE.margin, left: PAGE.margin } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: meta.headerLine || (meta.productName ? `${meta.productName} DVT` : 'DVT'), color: '666666', size: 16 })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Page ', size: 16, color: '666666' }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: '666666' }), new TextRun({ text: ' of ', size: 16, color: '666666' }), new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: '666666' })] })] }) },
    children,
  }],
});

// ============ Post-process: inject content-control checkboxes ============
async function injectCheckboxControls(buf) {
  const zip = await JSZip.loadAsync(buf);
  let docXml = await zip.file('word/document.xml').async('string');

  const w14Ns = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
  if (!docXml.includes(w14Ns)) {
    docXml = docXml.replace(/<w:document\b/, `<w:document ${w14Ns}`);
  }

  const pfBookmarks = (docXml.match(/w:name="pf_[^"]+"/g) || []).length;
  const ballotBoxLit = (docXml.match(/☐/g) || []).length;
  const ballotBoxEnt = (docXml.match(/&#x?2610;/gi) || []).length;
  console.log(`Diagnostic: pf_ bookmarks=${pfBookmarks}, ☐ literal=${ballotBoxLit}, &#x2610;-style entities=${ballotBoxEnt}`);

  const bkmRangeRe = /(<w:bookmarkStart\b[^>]*\bw:name="pf_[^"]+"[^>]*\/>)([\s\S]*?)(<w:bookmarkEnd\b[^>]*\/>)/g;
  const runWithBoxRe = /<w:r\b[\s\S]*?<w:t[^>]*>(?:☐|&#9744;|&#x2610;|&#X2610;)<\/w:t>[\s\S]*?<\/w:r>/;
  let sdtId = 50000;
  let replaceCount = 0;
  docXml = docXml.replace(bkmRangeRe, (full, bkmStart, inner, bkmEnd) => {
    const m = inner.match(runWithBoxRe);
    if (!m) return full;
    const runXml = m[0];
    const before = inner.slice(0, m.index);
    const after = inner.slice(m.index + runXml.length);
    replaceCount++;
    const id = sdtId++;
    const sdt =
      '<w:sdt>' +
        '<w:sdtPr>' +
          '<w:rPr/>' +
          `<w:id w:val="${id}"/>` +
          '<w14:checkbox>' +
            '<w14:checked w:val="0"/>' +
            '<w14:checkedState w:val="2612" w:font="MS Gothic"/>' +
            '<w14:uncheckedState w:val="2610" w:font="MS Gothic"/>' +
          '</w14:checkbox>' +
        '</w:sdtPr>' +
        `<w:sdtContent>${runXml}</w:sdtContent>` +
      '</w:sdt>';
    return `${bkmStart}${before}${sdt}${after}${bkmEnd}`;
  });

  zip.file('word/document.xml', docXml);
  const out = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  return { buf: out, count: replaceCount };
}

Packer.toBuffer(doc).then(async (buf) => {
  const { buf: finalBuf, count } = await injectCheckboxControls(buf);
  fs.writeFileSync(OUT, finalBuf);
  console.log(`Wrote ${OUT} (${finalBuf.length} bytes); injected ${count} content-control checkboxes.`);
}).catch((err) => {
  console.error(err);
  process.exit(1);
});
