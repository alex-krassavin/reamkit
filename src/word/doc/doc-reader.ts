// Legacy `.doc` reader (DOC-1..5) — the DocumentReader that sniffs a Word
// 97–2003 binary file (an OLE2/CFB container with a `WordDocument` stream) and
// reads its text, run/paragraph formatting, tables and inline images into the
// same FlowDoc the OOXML docx reader produces, so the whole render pipeline
// (projection → PDF/SVG/HTML, re-write to .docx) works on a legacy `.doc` the way
// it already does for `.xls`. The shared CFB container reader (`src/core/ole`) is
// the keystone both legacy formats reuse.
//
// Scope: paragraphs of formatted runs — bold / italic / underline / font size
// (CHPX) on each run, alignment / indentation / spacing (PAPX) on each paragraph,
// tables (in-table paragraphs grouped into a row/cell grid) and inline images
// (the picture char's CHPX → a PICF in the Data stream). Headers/footers, lists
// and fields are not read yet — recorded as a loss.

import type {
  Alignment,
  BodyElement,
  Border,
  CellBorders,
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
  DocCharProps,
  DocContent,
  DocParaProps,
  DocParagraph,
  DocPicture,
} from '@/word/doc/doc-text';

import { FEATURES, ResourceStore, pt, twipsToPt } from '@/core/ir';
import { isCfb, openCfb } from '@/core/ole/cfb';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';
import { extractDocContent } from '@/word/doc/doc-text';

const ZERO_WIDTH_SPACE = '​';

// US Letter (612pt) minus one-inch margins each side — the table's column band.
const CONTENT_WIDTH_PT = 468;
const TABLE_BORDER: Border = { style: 'single', width: pt(0.5), colorHex: '000000' };
const TABLE_BORDERS: CellBorders = {
  top: TABLE_BORDER,
  right: TABLE_BORDER,
  bottom: TABLE_BORDER,
  left: TABLE_BORDER,
  insideH: TABLE_BORDER,
  insideV: TABLE_BORDER,
};

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
    'legacy .doc: the document text, run formatting (bold/italic/underline/size), paragraph formatting (alignment/indent/spacing), tables and inline images are read; headers/footers, lists, fields and table cell widths/borders/merges are not (re-save as .docx for full fidelity)',
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

  // Word's default page: US Letter with one-inch margins.
  const doc: FlowDoc = {
    kind: 'flow',
    body: resolveBodyStyles(body, EMPTY_STYLE_SHEET),
    sections: [],
    section: {
      pageSize: { width: pt(612), height: pt(792) },
      margins: { top: pt(72), right: pt(72), bottom: pt(72), left: pt(72) },
      headers: [],
      footers: [],
    },
    styles: EMPTY_STYLE_SHEET,
    resources,
  };
  return { doc, losses: [content.encrypted ? DOC_ENCRYPTED_LOSS : DOC_TEXT_LOSS] };
}

// Group the extracted paragraphs into the FlowDoc body: runs of in-table
// paragraphs become Tables; everything else is a paragraph. doc-text has already
// split paragraphs at the CR / cell mark and cleaned the control characters.
function buildBody(content: DocContent, resources: ResourceStore): Array<BodyElement> {
  const body = groupTables(content.paragraphs, resources);
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
): Array<BodyElement> {
  const out: Array<BodyElement> = [];
  let i = 0;
  while (i < paras.length) {
    if (paras[i]!.props.inTable) {
      const group: Array<DocParagraph> = [];
      while (i < paras.length && paras[i]!.props.inTable) group.push(paras[i++]!);
      out.push({ kind: 'table', table: buildTable(group, resources) });
    } else {
      out.push(...mapParagraph(paras[i++]!, resources));
    }
  }
  return out;
}

// Split a run of in-table paragraphs into rows at each TTP (rowEnd); within a row
// each paragraph is a cell (an empty terminating paragraph is dropped). Cell
// widths are not in the PAPX we read, so columns share the content width evenly
// and a default thin border makes the grid visible.
function buildTable(group: ReadonlyArray<DocParagraph>, resources: ResourceStore): Table {
  const rows: Array<TableRow> = [];
  let rowParas: Array<DocParagraph> = [];
  for (const p of group) {
    rowParas.push(p);
    if (p.props.rowEnd) {
      rows.push(buildRow(rowParas, resources));
      rowParas = [];
    }
  }
  if (rowParas.length > 0) rows.push(buildRow(rowParas, resources)); // a row without a TTP

  const cols = Math.max(1, ...rows.map((r) => r.cells.length));
  const colWidth = pt(CONTENT_WIDTH_PT / cols);
  const grid = Array.from({ length: cols }, () => colWidth);
  return { properties: { borders: TABLE_BORDERS, layout: 'fixed' }, grid, rows };
}

function buildRow(rowParas: ReadonlyArray<DocParagraph>, resources: ResourceStore): TableRow {
  const cells: Array<TableCell> = [];
  for (const p of rowParas) {
    // The TTP terminates the row; when it carries no text it is not itself a cell.
    if (p.props.rowEnd && isBlank(p)) continue;
    cells.push({ properties: {}, content: mapParagraph(p, resources) });
  }
  if (cells.length === 0) cells.push({ properties: {}, content: [emptyParagraph()] });
  return { properties: {}, cells };
}

function isBlank(p: DocParagraph): boolean {
  return p.runs.every((r) => r.text.length === 0 && r.picture === undefined);
}

// A paragraph → its text paragraph (if it has text) followed by an image block per
// embedded picture. A picture occupies its own run, so a picture-only paragraph
// yields just the image(s).
function mapParagraph(p: DocParagraph, resources: ResourceStore): Array<BodyElement> {
  const out: Array<BodyElement> = [];
  const textRuns = p.runs.filter((r) => r.picture === undefined);
  if (textRuns.some((r) => r.text.length > 0)) {
    out.push({
      kind: 'paragraph',
      paragraph: {
        properties: toParaProperties(p.props),
        runs: textRuns.map((r) => ({ text: r.text, properties: toRunProperties(r.props) })),
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
