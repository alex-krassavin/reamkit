// Legacy `.doc` reader (DOC-1..11) — the DocumentReader that sniffs a Word
// 97–2003 binary file (an OLE2/CFB container with a `WordDocument` stream) and
// reads its text, run/paragraph formatting, tables, inline images,
// headers/footers and list items into the same FlowDoc the OOXML docx reader
// produces, so the whole render pipeline (projection → PDF/SVG/HTML, re-write to
// .docx) works on a legacy `.doc` the way it already does for `.xls`. The shared
// CFB container reader (`src/core/ole`) is the keystone both legacy formats reuse.
//
// Scope: paragraphs of formatted runs — bold / italic / underline / font size
// (CHPX) on each run, alignment / indentation / spacing (PAPX) on each paragraph,
// tables (in-table paragraphs grouped into a row/cell grid, with per-column widths
// from sprmTDefTable), inline images (the picture char's CHPX → a PICF in the Data
// stream), fields (resolved to their cached result), the section's headers/footers
// (the PlcfHdd stories), list items (sprmPIlfo/sprmPIlvl → the LST/LVL number
// format or bullet — DOC-10) and table cell borders + vertical merges (the row's
// TC80 array in sprmTDefTable — DOC-11).

import type {
  Alignment,
  BodyElement,
  Border,
  CellBorders,
  CellMerge,
  HeaderFooterReference,
  HeaderFooterType,
  ImageBlock,
  ParagraphProperties,
  RunProperties,
  Table,
  TableCell,
  TableRow,
  UnderlineStyle,
} from '@/core/document-model';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir';
import type {
  DocBorder,
  DocCharProps,
  DocContent,
  DocListTables,
  DocLvl,
  DocParaProps,
  DocParagraph,
  DocPicture,
  DocTc,
} from '@/word/doc/doc-text';

import { FEATURES, ResourceStore, pt, twipsToPt } from '@/core/ir';
import { isCfb, openCfb } from '@/core/ole/cfb';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';
import { extractDocContent } from '@/word/doc/doc-text';

const ZERO_WIDTH_SPACE = '​';

// US Letter (612pt) minus one-inch margins each side — the table's column band.
const CONTENT_WIDTH_PT = 468;
// List items render with a generic bullet per level + a quarter-inch indent per
// level (the exact numbered format is not read from the LST/LVL tables).
const LIST_BULLETS = ['•', '◦', '▪'];
const LIST_INDENT_PT = 18;
const TABLE_BORDER: Border = { style: 'single', width: pt(0.5), colorHex: '000000' };
const TABLE_BORDERS: CellBorders = {
  top: TABLE_BORDER,
  right: TABLE_BORDER,
  bottom: TABLE_BORDER,
  left: TABLE_BORDER,
  insideH: TABLE_BORDER,
  insideV: TABLE_BORDER,
};

// A list marker resolver (DOC-10): maps a list paragraph's (ilfo, ilvl) to its
// rendered marker text — a real number ("1.", "a)", "iii.") or a bullet glyph —
// keeping per-(lsid, level) running counters across the body. Returns undefined
// when the list tables can't resolve the level (the caller uses a generic bullet).
function createListMarkers(tables: DocListTables | undefined) {
  const counters = new Map<number, Array<number | undefined>>();
  const resolve = (ilfo: number, ilvl: number): { lsid: number; lvl: DocLvl } | undefined => {
    if (!tables || ilfo < 1 || ilfo > tables.lfoLsids.length) return undefined;
    const lsid = tables.lfoLsids[ilfo - 1]!;
    const list = tables.lists.get(lsid);
    const lvl = list?.levels[list.simple ? 0 : ilvl];
    return list && lvl ? { lsid, lvl } : undefined;
  };
  return {
    marker(ilfo: number, ilvl: number): string | undefined {
      const r = resolve(ilfo, ilvl);
      if (!r) return undefined;
      const { lsid, lvl } = r;
      // A bullet (nfc 23) / none (0xFF) / placeholder-less template renders literally.
      if (lvl.nfc === 23 || lvl.nfc === 0xff || !lvl.xst.some((cu) => cu <= 8)) {
        return lvl.xst.map((cu) => String.fromCharCode(cu)).join('');
      }
      // Advance this level's counter; reset the deeper levels (they re-seed later).
      const col = counters.get(lsid) ?? [];
      col[ilvl] = (col[ilvl] ?? lvl.iStartAt - 1) + 1;
      for (let l = ilvl + 1; l < col.length; l++) col[l] = undefined;
      counters.set(lsid, col);
      // Each placeholder → the number of that level, formatted by that level's nfc.
      let out = '';
      for (const cu of lvl.xst) {
        if (cu <= 8) {
          const other = resolve(ilfo, cu);
          out += formatListNumber(other?.lvl.nfc ?? 0, col[cu] ?? other?.lvl.iStartAt ?? 1);
        } else {
          out += String.fromCharCode(cu);
        }
      }
      return out;
    },
  };
}
type ListMarkers = ReturnType<typeof createListMarkers>;

// An MSONFC number format + value → the marker text for one level's number.
function formatListNumber(nfc: number, n: number): string {
  switch (nfc) {
    case 1:
      return toRoman(n);
    case 2:
      return toRoman(n).toLowerCase();
    case 3:
      return toLetters(n);
    case 4:
      return toLetters(n).toLowerCase();
    case 22:
      return n < 10 ? `0${n}` : String(n); // arabicLZ (decimal with leading zero)
    default:
      return String(n); // arabic + anything not specially numbered
  }
}
function toLetters(n: number): string {
  let s = '';
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}
const ROMAN: ReadonlyArray<readonly [number, string]> = [
  [1000, 'M'],
  [900, 'CM'],
  [500, 'D'],
  [400, 'CD'],
  [100, 'C'],
  [90, 'XC'],
  [50, 'L'],
  [40, 'XL'],
  [10, 'X'],
  [9, 'IX'],
  [5, 'V'],
  [4, 'IV'],
  [1, 'I'],
];
function toRoman(n: number): string {
  if (n <= 0) return String(n);
  let out = '';
  for (const [v, sym] of ROMAN) while (n >= v) ((out += sym), (n -= v));
  return out;
}

// A legacy `.doc` is an OLE2 container with a `WordDocument` stream. The check
// also keeps a `.xls`/`.ppt` (or an encrypted OOXML, also a CFB) from
// mis-routing here.
function looksLikeDoc(bytes: Uint8Array): boolean {
  if (!isCfb(bytes)) return false;
  try {
    return openCfb(bytes).hasStream('WordDocument');
  } catch {
    return false;
  }
}

// The text + run + paragraph formatting is read; the higher-level structure is
// not. Reported (degraded, keyed on text) so a caller's loss report is honest.
const DOC_TEXT_LOSS: Loss = {
  severity: 'degraded',
  feature: FEATURES.text,
  detail:
    "legacy .doc: the document text, run formatting (bold/italic/underline/size), paragraph formatting (alignment/indent/spacing), tables (with column widths, cell borders, vertical merges and cell background shading), inline images, fields, the section's headers/footers and list items (with their number format / bullet) are read; drawing shapes / text boxes and comments are not (re-save as .docx for full fidelity)",
};

const DOC_ENCRYPTED_LOSS: Loss = {
  severity: 'dropped',
  feature: FEATURES.text,
  detail: 'legacy .doc is encrypted/obfuscated — its text cannot be read',
};

export function readDoc(bytes: Uint8Array): ReadResult<FlowDoc> {
  const content = extractDocContent(bytes);
  const resources = new ResourceStore();
  const body = buildBody(content, resources);

  // Header/footer stories → a named-body map + the section's references.
  const headersFooters = new Map<string, ReadonlyArray<BodyElement>>();
  const headers: Array<HeaderFooterReference> = [];
  const footers: Array<HeaderFooterReference> = [];
  const addStory = (
    story: ReadonlyArray<DocParagraph> | undefined,
    kind: 'header' | 'footer',
    type: HeaderFooterType,
  ): void => {
    if (!story) return;
    const id = `${kind}-${type}`;
    const els = resolveBodyStyles(
      story.flatMap((p) => mapParagraph(p, resources)),
      EMPTY_STYLE_SHEET,
    );
    headersFooters.set(id, els);
    (kind === 'header' ? headers : footers).push({ type, relationshipId: id });
  };
  const hf = content.headerFooters;
  if (hf) {
    addStory(hf.defaultHeader, 'header', 'default');
    addStory(hf.firstHeader, 'header', 'first');
    addStory(hf.evenHeader, 'header', 'even');
    addStory(hf.defaultFooter, 'footer', 'default');
    addStory(hf.firstFooter, 'footer', 'first');
    addStory(hf.evenFooter, 'footer', 'even');
  }

  // Word's default page: US Letter with one-inch margins.
  const doc: FlowDoc = {
    kind: 'flow',
    body: resolveBodyStyles(body, EMPTY_STYLE_SHEET),
    sections: [],
    section: {
      pageSize: { width: pt(612), height: pt(792) },
      margins: { top: pt(72), right: pt(72), bottom: pt(72), left: pt(72) },
      headers,
      footers,
    },
    styles: EMPTY_STYLE_SHEET,
    resources,
    ...(headersFooters.size > 0 ? { headersFooters } : {}),
  };
  return { doc, losses: [content.encrypted ? DOC_ENCRYPTED_LOSS : DOC_TEXT_LOSS] };
}

// Group the extracted paragraphs into the FlowDoc body: runs of in-table
// paragraphs become Tables; everything else is a paragraph. doc-text has already
// split paragraphs at the CR / cell mark and cleaned the control characters.
function buildBody(content: DocContent, resources: ResourceStore): Array<BodyElement> {
  const markers = createListMarkers(content.listTables);
  const body = groupTables(content.paragraphs, resources, markers);
  // A document ends with a trailing paragraph mark — drop the empty tail it adds.
  while (body.length > 1 && isEmptyParagraph(body[body.length - 1]!)) body.pop();
  if (body.length === 0) body.push(emptyParagraph());
  return body;
}

// Consecutive in-table paragraphs (sprmPFInTable) collapse into one Table; the
// rest map straight to paragraphs (and image blocks).
function groupTables(
  paras: ReadonlyArray<DocParagraph>,
  resources: ResourceStore,
  markers: ListMarkers,
): Array<BodyElement> {
  const out: Array<BodyElement> = [];
  let i = 0;
  while (i < paras.length) {
    if (paras[i]!.props.inTable) {
      const group: Array<DocParagraph> = [];
      while (i < paras.length && paras[i]!.props.inTable) group.push(paras[i++]!);
      out.push({ kind: 'table', table: buildTable(group, resources, markers) });
    } else {
      out.push(...mapParagraph(paras[i++]!, resources, markers));
    }
  }
  return out;
}

// Split a run of in-table paragraphs into rows at each TTP (rowEnd); within a row
// each paragraph is a cell (an empty terminating paragraph is dropped). Column
// widths + per-cell borders come from the TTP's sprmTDefTable (DOC-7/DOC-11);
// vertical-merge markers are resolved across rows into CellMerge roles (DOC-11). A
// default thin border keeps an undescribed grid visible.
function buildTable(
  group: ReadonlyArray<DocParagraph>,
  resources: ResourceStore,
  markers: ListMarkers,
): Table {
  const edges = group.find((p) => p.props.cellEdgesTwips)?.props.cellEdgesTwips;
  const rawRows: Array<Array<RawCell>> = [];
  let rowParas: Array<DocParagraph> = [];
  for (const p of group) {
    rowParas.push(p);
    if (p.props.rowEnd) {
      rawRows.push(buildRawRow(rowParas, resources, edges, markers));
      rowParas = [];
    }
  }
  if (rowParas.length > 0) rawRows.push(buildRawRow(rowParas, resources, edges, markers)); // no TTP
  resolveVerticalMerges(rawRows);

  const rows: Array<TableRow> = rawRows.map((cells) => ({
    properties: {},
    cells: cells.map(toTableCell),
  }));
  return {
    properties: { borders: TABLE_BORDERS, layout: 'fixed' },
    grid: columnWidths(edges, rows),
    rows,
  };
}

// A cell before its vertical-merge role is resolved.
interface RawCell {
  readonly content: Array<BodyElement>;
  readonly width: ReturnType<typeof pt> | undefined;
  readonly borders: CellBorders | undefined;
  readonly shading: string | undefined; // background-fill hex from sprmTDefTableShd
  readonly vMergeRaw: 'restart' | 'continue' | undefined;
  merge?: CellMerge;
}

// Per-column widths from the cell boundaries (twips → points); without them the
// columns share the content width evenly.
function columnWidths(edges: ReadonlyArray<number> | undefined, rows: ReadonlyArray<TableRow>) {
  if (edges && edges.length >= 2) {
    const widths = [];
    for (let i = 0; i + 1 < edges.length; i++)
      widths.push(twipsToPt(Math.max(0, edges[i + 1]! - edges[i]!)));
    return widths;
  }
  const cols = Math.max(1, ...rows.map((r) => r.cells.length));
  const even = pt(CONTENT_WIDTH_PT / cols);
  return Array.from({ length: cols }, () => even);
}

function buildRawRow(
  rowParas: ReadonlyArray<DocParagraph>,
  resources: ResourceStore,
  edges: ReadonlyArray<number> | undefined,
  markers: ListMarkers,
): Array<RawCell> {
  // The row's TC80 descriptors (borders + vertical merge) + per-cell shading ride
  // on its TTP.
  const descriptors = rowParas.find((p) => p.props.cellDescriptors)?.props.cellDescriptors;
  const shadings = rowParas.find((p) => p.props.cellShadings)?.props.cellShadings;
  const cells: Array<RawCell> = [];
  let col = 0;
  for (const p of rowParas) {
    // The TTP terminates the row; when it carries no text it is not itself a cell.
    if (p.props.rowEnd && isBlank(p)) continue;
    const tc = descriptors?.[col];
    cells.push({
      content: mapParagraph(p, resources, markers),
      width: cellWidth(edges, col),
      borders: tc?.borders ? toCellBorders(tc.borders) : undefined,
      shading: shadings?.[col],
      vMergeRaw: tc?.vMerge,
    });
    col++;
  }
  if (cells.length === 0) {
    cells.push({
      content: [emptyParagraph()],
      width: undefined,
      borders: undefined,
      shading: undefined,
      vMergeRaw: undefined,
    });
  }
  return cells;
}

// Resolve the per-cell vertical-merge markers (DOC-11) into CellMerge roles: in
// each column, a `restart` followed by ≥1 `continue` becomes start / middle… / end
// (the renderer then spans the start cell down). A lone restart stays a normal cell.
function resolveVerticalMerges(rows: ReadonlyArray<ReadonlyArray<RawCell>>): void {
  const maxCols = Math.max(0, ...rows.map((r) => r.length));
  for (let c = 0; c < maxCols; c++) {
    let span: Array<RawCell> = [];
    const close = (): void => {
      if (span.length >= 2) {
        span[0]!.merge = 'start';
        span[span.length - 1]!.merge = 'end';
        for (let k = 1; k < span.length - 1; k++) span[k]!.merge = 'middle';
      }
      span = [];
    };
    for (const row of rows) {
      const cell = row[c];
      if (cell?.vMergeRaw === 'restart') {
        close();
        span = [cell];
      } else if (cell?.vMergeRaw === 'continue' && span.length > 0) {
        span.push(cell);
      } else {
        close();
      }
    }
    close();
  }
}

function toTableCell(raw: RawCell): TableCell {
  return {
    properties: {
      ...(raw.width ? { width: raw.width } : {}),
      ...(raw.borders ? { borders: raw.borders } : {}),
      ...(raw.shading ? { shading: { colorHex: raw.shading } } : {}),
      ...(raw.merge ? { merge: raw.merge } : {}),
    },
    content: raw.content,
  };
}

// One TC80's borders → the model's CellBorders (only the edges that have a border).
function toCellBorders(b: NonNullable<DocTc['borders']>): CellBorders {
  const edge = (db: DocBorder | undefined): Border | undefined =>
    db
      ? { style: db.double ? 'double' : 'single', width: pt(db.widthPt), colorHex: db.colorHex }
      : undefined;
  const top = edge(b.top);
  const left = edge(b.left);
  const bottom = edge(b.bottom);
  const right = edge(b.right);
  return {
    ...(top ? { top } : {}),
    ...(left ? { left } : {}),
    ...(bottom ? { bottom } : {}),
    ...(right ? { right } : {}),
  };
}

function cellWidth(edges: ReadonlyArray<number> | undefined, col: number) {
  if (edges && col + 1 < edges.length) {
    const w = edges[col + 1]! - edges[col]!;
    if (w > 0) return twipsToPt(w);
  }
  return undefined;
}

function isBlank(p: DocParagraph): boolean {
  return p.runs.every((r) => r.text.length === 0 && r.picture === undefined);
}

// A paragraph → its text paragraph (if it has text) followed by an image block per
// embedded picture. A picture occupies its own run, so a picture-only paragraph
// yields just the image(s).
function mapParagraph(
  p: DocParagraph,
  resources: ResourceStore,
  markers?: ListMarkers,
): Array<BodyElement> {
  const out: Array<BodyElement> = [];
  const textRuns = p.runs.filter((r) => r.picture === undefined);
  const ilvl = p.props.listIlvl ?? 0;
  const isList = (p.props.listIlfo ?? 0) > 0;
  if (textRuns.some((r) => r.text.length > 0) || isList) {
    const runs = textRuns.map((r) => ({ text: r.text, properties: toRunProperties(r.props) }));
    let properties = toParaProperties(p.props);
    if (isList) {
      // The LST/LVL number format (DOC-10) → a real "1." / "a)" / bullet marker;
      // an unresolved list falls back to a generic per-level bullet. + a per-level indent.
      const marker =
        markers?.marker(p.props.listIlfo!, ilvl) ?? LIST_BULLETS[ilvl % LIST_BULLETS.length];
      runs.unshift({ text: `${marker} `, properties: {} });
      properties = { ...properties, indentLeft: pt((ilvl + 1) * LIST_INDENT_PT) };
    }
    out.push({
      kind: 'paragraph',
      paragraph: {
        properties,
        runs: runs.length > 0 ? runs : [{ text: ZERO_WIDTH_SPACE, properties: {} }],
      },
    });
  }
  for (const r of p.runs) {
    if (r.picture) out.push(imageBlock(r.picture, resources, toParaProperties(p.props)));
  }
  if (out.length === 0) out.push(emptyParagraph()); // preserve the (empty) line
  return out;
}

function imageBlock(
  pic: DocPicture,
  resources: ResourceStore,
  paragraphProperties: ParagraphProperties,
): BodyElement {
  const image: ImageBlock = {
    resource: resources.put(pic.bytes),
    width: twipsToPt(pic.widthTwips),
    height: twipsToPt(pic.heightTwips),
    paragraphProperties,
  };
  return { kind: 'image', image };
}

function emptyParagraph(): BodyElement {
  return {
    kind: 'paragraph',
    paragraph: { properties: {}, runs: [{ text: ZERO_WIDTH_SPACE, properties: {} }] },
  };
}

function isEmptyParagraph(el: BodyElement): boolean {
  return (
    el.kind === 'paragraph' &&
    el.paragraph.runs.length === 1 &&
    el.paragraph.runs[0]!.text === ZERO_WIDTH_SPACE
  );
}

function toRunProperties(p: DocCharProps): RunProperties {
  const underline = p.underlineKul ? underlineStyle(p.underlineKul) : undefined;
  const size = p.sizeHalfPts && p.sizeHalfPts > 0 ? pt(p.sizeHalfPts / 2) : undefined;
  return {
    ...(p.bold ? { bold: true } : {}),
    ...(p.italic ? { italic: true } : {}),
    ...(underline ? { underline } : {}),
    ...(size ? { fontSizePt: size } : {}),
  };
}

function toParaProperties(p: DocParaProps): ParagraphProperties {
  const alignment = p.jc !== undefined ? alignmentFromJc(p.jc) : undefined;
  // Twips → points (20 twips per point); indents may be negative (hanging).
  return {
    ...(alignment ? { alignment } : {}),
    ...(p.indentLeftTwips !== undefined ? { indentLeft: pt(p.indentLeftTwips / 20) } : {}),
    ...(p.indentRightTwips !== undefined ? { indentRight: pt(p.indentRightTwips / 20) } : {}),
    ...(p.indentFirstTwips !== undefined ? { indentFirstLine: pt(p.indentFirstTwips / 20) } : {}),
    ...(p.spaceBeforeTwips !== undefined ? { spacingBefore: pt(p.spaceBeforeTwips / 20) } : {}),
    ...(p.spaceAfterTwips !== undefined ? { spacingAfter: pt(p.spaceAfterTwips / 20) } : {}),
  };
}

// Word's `kul` underline code → the document-model underline style.
function underlineStyle(kul: number): UnderlineStyle | undefined {
  switch (kul) {
    case 0:
      return undefined;
    case 3:
      return 'double';
    case 4:
      return 'dotted';
    case 7:
      return 'dash';
    default:
      return 'single';
  }
}

// Word's `jc` justification code → the document-model alignment (0 = left default).
function alignmentFromJc(jc: number): Alignment | undefined {
  switch (jc) {
    case 1:
      return 'center';
    case 2:
      return 'right';
    case 3:
      return 'both';
    case 4:
      return 'distribute';
    default:
      return undefined;
  }
}

export const docReader: DocumentReader<FlowDoc> = {
  id: 'doc',
  produces: 'flow',
  supports: new Set([FEATURES.text]),
  sniff: looksLikeDoc,
  read: (bytes) => readDoc(bytes),
};
