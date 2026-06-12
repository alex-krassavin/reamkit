// docx writer (E-DOCX): FlowDoc → WordprocessingML package — the inverse of
// the docx reader, and the fifth adapter overall. A flow medium with zero
// layout and zero I/O, like the HTML writer.
//
// v1 contract (epics.md, variant A): the writer emits a DENORMALIZED but
// valid document. FlowDoc's body carries RESOLVED properties (the stage-6
// cascade is already collapsed), so what we write is direct formatting — no
// named styles. The round-trip guarantee is therefore semantic, not textual:
// readDocx(writeDocx(flow)) yields an equivalent FlowDoc, never the original
// bytes. Anything the writer does not serialize yet is reported as a loss,
// exactly like the other writers.

import type {
  BodyElement,
  CellBorders,
  CellMargins,
  CellProperties,
  FontFamilyMap,
  Numbering,
  NumberingLevel,
  Paragraph,
  ParagraphProperties,
  Run,
  RunProperties,
  SectionProperties,
  Table,
  TableCell,
  TableProperties,
  TableRow,
} from '@/core/document-model';
import type { ResolvedParagraphProperties, ResolvedRunProperties } from '@/core/style-cascade';
import type { DocumentWriter, WriteResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir';
import type { Relationship } from '@/core/opc';

import { FEATURES } from '@/core/ir';
import { buildOpcPackage } from '@/core/opc';
import {
  EMPTY_STYLE_SHEET,
  resolveParagraphProperties,
  resolveRunProperties,
} from '@/core/style-cascade';

const encoder = new TextEncoder();

// The reader stored RESOLVED properties back onto each run/paragraph (stage
// 6). The defaults are what the same resolver yields for empty input over the
// empty sheet — a field equal to these is implicit and is NOT serialized, so a
// re-read materializes the same value. This delta keeps the emitted rPr/pPr
// minimal and the round-trip an IR identity.
const DEFAULT_RUN = resolveRunProperties({}, {}, EMPTY_STYLE_SHEET);
const DEFAULT_PARA = resolveParagraphProperties({}, EMPTY_STYLE_SHEET);

const DOC_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const NUMBERING_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
const REL_OFFICE_DOCUMENT =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const REL_NUMBERING =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';
const REL_HYPERLINK =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const NUMBERING_PART = 'word/numbering.xml';

// Mutable per-document state threaded through the block/paragraph emitters:
// the document.xml relationships accumulator (numbering + each external
// hyperlink share one rId space) and a document-wide bookmark id counter.
interface WriteState {
  readonly docRels: Array<Relationship>;
  relSeq: number;
  bookmarkSeq: number;
}

export function writeDocx(flow: FlowDoc): WriteResult {
  const losses: Array<Loss> = [];
  const body: Array<string> = [];
  const state: WriteState = { docRels: [], relSeq: 0, bookmarkSeq: 0 };

  for (const el of flow.body) emitBlock(body, el, losses, state);

  const section = flow.sections[0]?.properties ?? flow.section;
  if (section) body.push(sectPrXml(section));

  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<w:body>${body.join('')}</w:body>` +
    '</w:document>';

  // §17.9 numbering: re-emit the raw definitions whenever a paragraph carries
  // a list reference (the markers were stripped above — re-read regenerates
  // them). Lives at the fixed word/numbering.xml path the reader expects, with
  // a content-type Override and a relationship from the main document.
  const usesNumbering = flow.body.some(
    (el) => el.kind === 'paragraph' && el.paragraph.properties.numbering !== undefined,
  );
  const numberingPart =
    usesNumbering && flow.numbering
      ? {
          path: NUMBERING_PART,
          data: encoder.encode(numberingXml(flow.numbering)),
          contentType: NUMBERING_CONTENT_TYPE,
        }
      : undefined;
  if (numberingPart) {
    state.docRels.push({
      id: `rId${++state.relSeq}`,
      type: REL_NUMBERING,
      target: 'numbering.xml',
      targetMode: 'Internal',
    });
  }

  const bytes = buildOpcPackage({
    parts: [
      {
        path: 'word/document.xml',
        data: encoder.encode(documentXml),
        contentType: DOC_CONTENT_TYPE,
      },
      ...(numberingPart ? [numberingPart] : []),
    ],
    rootRelationships: [
      {
        id: 'rId1',
        type: REL_OFFICE_DOCUMENT,
        target: 'word/document.xml',
        targetMode: 'Internal',
      },
    ],
    ...(state.docRels.length > 0
      ? {
          partRelationships: [{ sourcePart: 'word/document.xml', relationships: state.docRels }],
        }
      : {}),
  });

  return { bytes, losses };
}

export const docxWriter: DocumentWriter<FlowDoc> = {
  id: 'docx',
  consumes: 'flow',
  supports: new Set([FEATURES.text]),
  write: (doc) => writeDocx(doc),
};

function emitBlock(
  out: Array<string>,
  el: BodyElement,
  losses: Array<Loss>,
  state: WriteState,
): void {
  if (el.kind === 'paragraph') {
    out.push(paragraphXml(el.paragraph, state));
    return;
  }
  if (el.kind === 'table') {
    out.push(tableXml(el.table, losses, state));
    return;
  }
  // Images, charts, shapes: D5 of the epic. Reported, not dropped silently.
  const feature =
    el.kind === 'image' ? FEATURES.images : el.kind === 'chart' ? FEATURES.charts : FEATURES.shapes;
  losses.push({
    severity: 'dropped',
    feature,
    detail: `${el.kind} not written by the docx writer (v0)`,
  });
}

// §17.4 — w:tbl: properties, the column grid, then rows. Cell content recurses
// through emitBlock, so nested tables and per-cell paragraphs round-trip.
function tableXml(table: Table, losses: Array<Loss>, state: WriteState): string {
  const grid = table.grid.map((w) => `<w:gridCol w:w="${twips(w)}"/>`).join('');
  const rows = table.rows.map((row) => rowXml(row, losses, state)).join('');
  return `<w:tbl>${tblPrXml(table.properties)}<w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`;
}

function tblPrXml(p: TableProperties): string {
  const out: Array<string> = [];
  if (p.widthType !== undefined) {
    const w =
      p.widthType === 'pct'
        ? Math.round((p.widthFraction ?? 0) * 5000)
        : p.widthPt !== undefined
          ? twips(p.widthPt)
          : 0;
    out.push(`<w:tblW w:w="${w}" w:type="${p.widthType}"/>`);
  }
  if (p.alignment && p.alignment !== 'left') out.push(`<w:jc w:val="${p.alignment}"/>`);
  const borders = bordersXml('w:tblBorders', p.borders);
  if (borders) out.push(borders);
  const margins = cellMarginsXml('w:tblCellMar', p.defaultCellMargins);
  if (margins) out.push(margins);
  return out.length > 0 ? `<w:tblPr>${out.join('')}</w:tblPr>` : '';
}

function rowXml(row: TableRow, losses: Array<Loss>, state: WriteState): string {
  const trPr: Array<string> = [];
  if (row.properties.height !== undefined) {
    const rule = row.properties.heightRule ?? 'atLeast';
    trPr.push(`<w:trHeight w:val="${twips(row.properties.height)}" w:hRule="${rule}"/>`);
  }
  if (row.properties.isHeader) trPr.push('<w:tblHeader/>');
  if (row.properties.cantSplit) trPr.push('<w:cantSplit/>');
  const trPrXml = trPr.length > 0 ? `<w:trPr>${trPr.join('')}</w:trPr>` : '';
  const cells = row.cells.map((cell) => cellXml(cell, losses, state)).join('');
  return `<w:tr>${trPrXml}${cells}</w:tr>`;
}

function cellXml(cell: TableCell, losses: Array<Loss>, state: WriteState): string {
  const content: Array<string> = [];
  for (const child of cell.content) emitBlock(content, child, losses, state);
  // §17.4.66 — a w:tc must contain at least one block, ending in a paragraph.
  if (content.length === 0) content.push('<w:p/>');
  return `<w:tc>${tcPrXml(cell.properties)}${content.join('')}</w:tc>`;
}

function tcPrXml(p: CellProperties): string {
  const out: Array<string> = [];
  if (p.width !== undefined) out.push(`<w:tcW w:w="${twips(p.width)}" w:type="dxa"/>`);
  if (p.colSpan !== undefined && p.colSpan > 1) {
    out.push(`<w:gridSpan w:val="${p.colSpan}"/>`);
  }
  if (p.merge !== undefined) {
    // §17.4.85 — 'start' restarts a vertical merge; 'middle'/'end' continue it.
    out.push(p.merge === 'start' ? '<w:vMerge w:val="restart"/>' : '<w:vMerge w:val="continue"/>');
  }
  const borders = bordersXml('w:tcBorders', p.borders);
  if (borders) out.push(borders);
  const margins = cellMarginsXml('w:tcMar', p.margins);
  if (margins) out.push(margins);
  if (p.shading) out.push(`<w:shd w:val="clear" w:color="auto" w:fill="${p.shading.colorHex}"/>`);
  return out.length > 0 ? `<w:tcPr>${out.join('')}</w:tcPr>` : '';
}

const BORDER_SIDES: ReadonlyArray<[keyof CellBorders, string]> = [
  ['top', 'w:top'],
  ['left', 'w:left'],
  ['bottom', 'w:bottom'],
  ['right', 'w:right'],
  ['insideH', 'w:insideH'],
  ['insideV', 'w:insideV'],
];

function bordersXml(tag: 'w:tblBorders' | 'w:tcBorders', borders: CellBorders | undefined): string {
  if (!borders) return '';
  const sides = BORDER_SIDES.map(([key, el]) => {
    const b = borders[key];
    if (!b) return '';
    // §17.4.x — w:sz in eighths of a point; the reader divides by 8.
    const sz = b.width !== undefined ? ` w:sz="${Math.round(b.width * 8)}"` : '';
    const color = b.colorHex !== undefined ? ` w:color="${b.colorHex}"` : '';
    return `<${el} w:val="${b.style}"${sz}${color}/>`;
  }).join('');
  return sides ? `<${tag}>${sides}</${tag}>` : '';
}

function cellMarginsXml(tag: 'w:tblCellMar' | 'w:tcMar', margins: CellMargins | undefined): string {
  if (!margins) return '';
  const sides: Array<[keyof CellMargins, string]> = [
    ['top', 'w:top'],
    ['left', 'w:left'],
    ['bottom', 'w:bottom'],
    ['right', 'w:right'],
  ];
  const inner = sides
    .map(([key, el]) => {
      const v = margins[key];
      return v !== undefined ? `<${el} w:w="${twips(v)}" w:type="dxa"/>` : '';
    })
    .join('');
  return inner ? `<${tag}>${inner}</${tag}>` : '';
}

function paragraphXml(p: Paragraph, state: WriteState): string {
  // The runs the reader actually kept: list markers re-materialize from
  // numbering.xml, and math/inline-image runs are not written yet.
  const visible = p.runs.filter(
    (run) =>
      !run.listMarker && run.text !== '' && run.math === undefined && run.inlineImage === undefined,
  );

  // §17.16.22 — group adjacent runs sharing a hyperlink target back into one
  // w:hyperlink container (the reader stamped href/anchor onto every run
  // inside it; this is the inverse).
  const inner: Array<string> = [];
  let i = 0;
  while (i < visible.length) {
    const run = visible[i]!;
    const key = run.href ?? run.anchor;
    if (key === undefined) {
      inner.push(runXml(run));
      i++;
      continue;
    }
    let j = i + 1;
    while (j < visible.length && (visible[j]!.href ?? visible[j]!.anchor) === key) j++;
    const group = visible.slice(i, j).map(runXml).join('');
    inner.push(hyperlinkXml(run, group, state));
    i = j;
  }

  // §17.13.6.2 bookmarks: opened at the paragraph (start + end with a unique
  // id each); the reader reads the start, the end keeps the markup valid.
  const bookmarks = (p.bookmarks ?? [])
    .map((name) => {
      const id = state.bookmarkSeq++;
      return `<w:bookmarkStart w:id="${id}" w:name="${escapeAttr(name)}"/><w:bookmarkEnd w:id="${id}"/>`;
    })
    .join('');

  return `<w:p>${pPrXml(p.properties as ResolvedParagraphProperties)}${bookmarks}${inner.join('')}</w:p>`;
}

// w:hyperlink container: @r:id for an external target (allocates a rel),
// @w:anchor for an internal bookmark reference.
function hyperlinkXml(run: Run, inner: string, state: WriteState): string {
  if (run.href !== undefined) {
    const id = `rId${++state.relSeq}`;
    state.docRels.push({
      id,
      type: REL_HYPERLINK,
      target: run.href,
      targetMode: 'External',
    });
    return `<w:hyperlink r:id="${id}">${inner}</w:hyperlink>`;
  }
  return `<w:hyperlink w:anchor="${escapeAttr(run.anchor!)}">${inner}</w:hyperlink>`;
}

function runXml(run: Run): string {
  const rPr = rPrXml(run.properties as ResolvedRunProperties);
  return `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(run.text)}</w:t></w:r>`;
}

// §17.3.2 — run properties as a delta from the resolved defaults.
function rPrXml(r: ResolvedRunProperties): string {
  const out: Array<string> = [];
  if (r.bold !== DEFAULT_RUN.bold) out.push(toggle('w:b', r.bold));
  if (r.italic !== DEFAULT_RUN.italic) out.push(toggle('w:i', r.italic));
  if (r.strike !== DEFAULT_RUN.strike) out.push(toggle('w:strike', r.strike));
  if (r.underline !== DEFAULT_RUN.underline) out.push(`<w:u w:val="${r.underline}"/>`);
  const fonts = rFontsXml(r.fontFamily);
  if (fonts) out.push(fonts);
  if (r.fontSizePt !== DEFAULT_RUN.fontSizePt) {
    // §17.3.2.38 w:sz — half-points.
    out.push(`<w:sz w:val="${Math.round(r.fontSizePt * 2)}"/>`);
  }
  if (r.colorHex !== DEFAULT_RUN.colorHex) out.push(`<w:color w:val="${r.colorHex}"/>`);
  if (r.verticalAlign !== DEFAULT_RUN.verticalAlign) {
    out.push(`<w:vertAlign w:val="${r.verticalAlign}"/>`);
  }
  if (r.rtl !== DEFAULT_RUN.rtl) out.push(toggle('w:rtl', r.rtl));
  if (r.lang !== undefined) out.push(`<w:lang w:val="${escapeAttr(r.lang)}"/>`);
  return out.length > 0 ? `<w:rPr>${out.join('')}</w:rPr>` : '';
}

// §17.3.1 — paragraph properties as a delta from the resolved defaults.
function pPrXml(p: ResolvedParagraphProperties): string {
  const out: Array<string> = [];
  if (p.numbering) {
    // §17.3.1.19 — list membership; the marker itself comes from numbering.xml.
    out.push(
      `<w:numPr><w:ilvl w:val="${p.numbering.ilvl}"/><w:numId w:val="${escapeAttr(p.numbering.numId)}"/></w:numPr>`,
    );
  }
  if (p.outlineLevel !== undefined) out.push(`<w:outlineLvl w:val="${p.outlineLevel}"/>`);
  if (p.pageBreakBefore) out.push('<w:pageBreakBefore/>');
  if (p.bidi !== DEFAULT_PARA.bidi) out.push(toggle('w:bidi', p.bidi));
  const ind = indXml(p);
  if (ind) out.push(ind);
  const spacing = spacingXml(p);
  if (spacing) out.push(spacing);
  if (p.alignment !== DEFAULT_PARA.alignment) out.push(`<w:jc w:val="${p.alignment}"/>`);
  return out.length > 0 ? `<w:pPr>${out.join('')}</w:pPr>` : '';
}

function indXml(p: ResolvedParagraphProperties): string {
  const attrs: Array<string> = [];
  if (p.indentLeft !== DEFAULT_PARA.indentLeft) attrs.push(`w:left="${twips(p.indentLeft)}"`);
  if (p.indentRight !== DEFAULT_PARA.indentRight) attrs.push(`w:right="${twips(p.indentRight)}"`);
  if (p.indentFirstLine !== DEFAULT_PARA.indentFirstLine) {
    // A negative first-line indent is a hanging indent (§17.3.1.12).
    if (p.indentFirstLine < 0) attrs.push(`w:hanging="${twips(-p.indentFirstLine)}"`);
    else attrs.push(`w:firstLine="${twips(p.indentFirstLine)}"`);
  }
  return attrs.length > 0 ? `<w:ind ${attrs.join(' ')}/>` : '';
}

function spacingXml(p: ResolvedParagraphProperties): string {
  const attrs: Array<string> = [];
  if (p.spacingBefore !== DEFAULT_PARA.spacingBefore) {
    attrs.push(`w:before="${twips(p.spacingBefore)}"`);
  }
  if (p.spacingAfter !== DEFAULT_PARA.spacingAfter) {
    attrs.push(`w:after="${twips(p.spacingAfter)}"`);
  }
  if (
    (p.spacingLineRule !== DEFAULT_PARA.spacingLineRule ||
      p.spacingLine !== DEFAULT_PARA.spacingLine) &&
    p.spacingLine > 0
  ) {
    // §17.3.1.33: 'auto' line spacing is in 240ths (line units); exact/atLeast
    // in twips. The reader stores spacingLine in points either way.
    const lineVal =
      p.spacingLineRule === 'auto' ? Math.round(p.spacingLine * 12) : twips(p.spacingLine);
    attrs.push(`w:line="${lineVal}"`, `w:lineRule="${p.spacingLineRule}"`);
  }
  return attrs.length > 0 ? `<w:spacing ${attrs.join(' ')}/>` : '';
}

// §17.3.2.26 w:rFonts — only the slots that differ from the resolved default.
function rFontsXml(fonts: FontFamilyMap): string {
  const d = DEFAULT_RUN.fontFamily;
  const attrs: Array<string> = [];
  if (fonts.ascii && fonts.ascii !== d.ascii) attrs.push(`w:ascii="${escapeAttr(fonts.ascii)}"`);
  if (fonts.hAnsi && fonts.hAnsi !== d.hAnsi) attrs.push(`w:hAnsi="${escapeAttr(fonts.hAnsi)}"`);
  if (fonts.cs && fonts.cs !== d.cs) attrs.push(`w:cs="${escapeAttr(fonts.cs)}"`);
  return attrs.length > 0 ? `<w:rFonts ${attrs.join(' ')}/>` : '';
}

// A boolean toggle property (§17.3.2.x): present-true is bare; present-false is
// w:val="false" (overrides an inherited true — exact on re-read here).
function toggle(tag: string, on: boolean): string {
  return on ? `<${tag}/>` : `<${tag} w:val="false"/>`;
}

// §17.9.1 numbering.xml: every abstractNum (levels with start/numFmt/lvlText
// and the level's raw pPr/rPr), then the num instances binding numId →
// abstractNumId. Re-emitted from the FlowDoc's raw `numbering` round-trip
// material, so re-read regenerates identical markers.
function numberingXml(numbering: Numbering): string {
  const out: Array<string> = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
  ];
  for (const abstractNum of numbering.abstractNums.values()) {
    out.push(`<w:abstractNum w:abstractNumId="${escapeAttr(abstractNum.id)}">`);
    for (const level of [...abstractNum.levels.values()].sort((a, b) => a.ilvl - b.ilvl)) {
      out.push(levelXml(level));
    }
    out.push('</w:abstractNum>');
  }
  for (const inst of numbering.numInstances.values()) {
    out.push(
      `<w:num w:numId="${escapeAttr(inst.numId)}">` +
        `<w:abstractNumId w:val="${escapeAttr(inst.abstractNumId)}"/></w:num>`,
    );
  }
  out.push('</w:numbering>');
  return out.join('');
}

function levelXml(level: NumberingLevel): string {
  const inner: Array<string> = [
    `<w:start w:val="${level.start}"/>`,
    `<w:numFmt w:val="${level.format}"/>`,
    `<w:lvlText w:val="${escapeAttr(level.lvlText)}"/>`,
  ];
  const pPr = rawParaPrXml(level.paragraphProperties);
  if (pPr) inner.push(pPr);
  const rPr = rawRunPrXml(level.runProperties);
  if (rPr) inner.push(rPr);
  return `<w:lvl w:ilvl="${level.ilvl}">${inner.join('')}</w:lvl>`;
}

// A numbering level's RAW (sparse) paragraph props — present fields only,
// no delta-against-defaults (unlike the resolved-body serializer above).
function rawParaPrXml(p: ParagraphProperties): string {
  const attrs: Array<string> = [];
  if (p.indentLeft !== undefined) attrs.push(`w:left="${twips(p.indentLeft)}"`);
  if (p.indentRight !== undefined) attrs.push(`w:right="${twips(p.indentRight)}"`);
  if (p.indentFirstLine !== undefined) {
    if (p.indentFirstLine < 0) attrs.push(`w:hanging="${twips(-p.indentFirstLine)}"`);
    else attrs.push(`w:firstLine="${twips(p.indentFirstLine)}"`);
  }
  const ind = attrs.length > 0 ? `<w:ind ${attrs.join(' ')}/>` : '';
  const jc = p.alignment !== undefined ? `<w:jc w:val="${p.alignment}"/>` : '';
  return ind || jc ? `<w:pPr>${jc}${ind}</w:pPr>` : '';
}

function rawRunPrXml(r: RunProperties): string {
  const out: Array<string> = [];
  if (r.bold !== undefined) out.push(toggle('w:b', r.bold));
  if (r.italic !== undefined) out.push(toggle('w:i', r.italic));
  if (r.strike !== undefined) out.push(toggle('w:strike', r.strike));
  if (r.underline !== undefined) out.push(`<w:u w:val="${r.underline}"/>`);
  const fonts = r.fontFamily ? rawRFontsXml(r.fontFamily) : '';
  if (fonts) out.push(fonts);
  if (r.fontSizePt !== undefined) out.push(`<w:sz w:val="${Math.round(r.fontSizePt * 2)}"/>`);
  if (r.colorHex !== undefined) out.push(`<w:color w:val="${r.colorHex}"/>`);
  if (r.verticalAlign !== undefined) out.push(`<w:vertAlign w:val="${r.verticalAlign}"/>`);
  return out.length > 0 ? `<w:rPr>${out.join('')}</w:rPr>` : '';
}

function rawRFontsXml(fonts: FontFamilyMap): string {
  const attrs: Array<string> = [];
  if (fonts.ascii) attrs.push(`w:ascii="${escapeAttr(fonts.ascii)}"`);
  if (fonts.hAnsi) attrs.push(`w:hAnsi="${escapeAttr(fonts.hAnsi)}"`);
  if (fonts.cs) attrs.push(`w:cs="${escapeAttr(fonts.cs)}"`);
  return attrs.length > 0 ? `<w:rFonts ${attrs.join(' ')}/>` : '';
}

// §17.6.17 — page size and margins in twentieths of a point.
function sectPrXml(s: SectionProperties): string {
  const parts: Array<string> = [];
  if (s.pageSize) {
    const orient = s.pageSize.orientation === 'landscape' ? ' w:orient="landscape"' : '';
    parts.push(
      `<w:pgSz w:w="${twips(s.pageSize.width)}" w:h="${twips(s.pageSize.height)}"${orient}/>`,
    );
  }
  if (s.margins) {
    const m = s.margins;
    const header = m.header !== undefined ? ` w:header="${twips(m.header)}"` : '';
    const footer = m.footer !== undefined ? ` w:footer="${twips(m.footer)}"` : '';
    parts.push(
      `<w:pgMar w:top="${twips(m.top)}" w:right="${twips(m.right)}"` +
        ` w:bottom="${twips(m.bottom)}" w:left="${twips(m.left)}"${header}${footer}/>`,
    );
  }
  if (parts.length === 0) return '';
  return `<w:sectPr>${parts.join('')}</w:sectPr>`;
}

const twips = (pt: number): number => Math.round(pt * 20);

function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeAttr(s: string): string {
  return escapeXml(s).replace(/"/g, '&quot;');
}
