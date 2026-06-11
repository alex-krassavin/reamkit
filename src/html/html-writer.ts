// HTML writer (ir-design §7): the fourth adapter — and the first one written
// against the FLOW tree. Where the SVG writer crash-tested the PageDoc
// contract, this one crash-tests FlowDoc: a semantic document in a flow
// medium, so there is no pagination, no layout engine and no fonts to embed —
// `Ream.parse(bytes).convert('html')` performs zero I/O.
//
// Stage-6 contract at work: `flow.body` already carries FINAL effective
// properties (the readers materialize list markers and resolve the style
// cascade), so this writer maps values straight to CSS. It still routes every
// property read through the cascade resolver over the EMPTY sheet — on reader
// output that is a memoized identity (the fixpoint registered at parse time),
// and on a hand-built raw tree it degrades to actually resolving, exactly
// like the PDF layout does.
//
// Deliberately out of scope (reported as losses): chart and shape geometry
// (text inside shapes IS emitted), inline math, headers/footers (no fixed
// pages to band them onto). Page-break hints map to CSS `break-before` so
// printing the HTML approximates the source pagination.

import type {
  BodyElement,
  Border,
  ImageBlock,
  Paragraph,
  Run,
  Table,
  TableCell,
} from '@/core/document-model';
import type { ResolvedParagraphProperties, ResolvedRunProperties } from '@/core/style-cascade';
import type { DocumentWriter, WriteResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss, ResourceId, ResourceStore } from '@/core/ir';

import { toBase64 } from '@/core/bytes';
import { detectImageFormat } from '@/core/images';
import { sanitizeHref } from '@/core/links';
import { FEATURES } from '@/core/ir';
import {
  EMPTY_STYLE_SHEET,
  resolveParagraphProperties,
  resolveRunProperties,
} from '@/core/style-cascade';

export function writeHtml(flow: FlowDoc): WriteResult {
  const losses: Array<Loss> = [];
  const out: Array<string> = [];
  const ctx: EmitCtx = { resources: flow.resources, losses };

  const lang = flow.language ?? 'en-US';
  const title = flow.info?.title ?? 'Document';
  out.push('<!DOCTYPE html>');
  out.push(`<html lang="${escapeAttr(lang)}">`);
  out.push('<head>');
  out.push('<meta charset="utf-8"/>');
  out.push('<meta name="viewport" content="width=device-width, initial-scale=1"/>');
  out.push(`<title>${escapeText(title)}</title>`);
  out.push(`<style>${BASE_CSS}</style>`);
  out.push('</head>');
  out.push('<body>');
  // A document-shaped column: the first section's content width, like the
  // page the source was authored for (A4 + 1" margins when unspecified —
  // the same fallback the layout engine uses).
  out.push(`<article style="max-width: ${fmt(contentWidthPt(flow))}pt">`);

  if (flow.headersFooters && flow.headersFooters.size > 0) {
    losses.push({
      severity: 'dropped',
      feature: FEATURES.headersFooters,
      detail: 'headers/footers have no fixed pages to band onto in flowed HTML',
    });
  }

  for (const el of flow.body) emitBlock(out, el, ctx);

  out.push('</article>');
  out.push('</body>');
  out.push('</html>');
  return { bytes: new TextEncoder().encode(out.join('\n')), losses };
}

export const htmlWriter: DocumentWriter<FlowDoc> = {
  id: 'html',
  consumes: 'flow',
  supports: new Set([
    FEATURES.text,
    FEATURES.tables,
    FEATURES.tablesNested,
    FEATURES.lists,
    FEATURES.images,
    FEATURES.hyperlinks,
    FEATURES.rtl,
    FEATURES.trackedChanges,
  ]),
  write: (doc) => writeHtml(doc),
};

// Paragraph margins are zeroed in the base CSS because the document's own
// resolved spacing is emitted per element; `.tab` renders the tab the list
// markers carry ("1.\t…") as a fixed gap.
const BASE_CSS = [
  'body{margin:24pt;color:#000;background:#fff}',
  'article{margin:0 auto}',
  'p,h1,h2,h3,h4,h5,h6,figure{margin:0}',
  'table{border-collapse:collapse}',
  'td,th{vertical-align:top}',
  'th{text-align:inherit}',
  '.tab{display:inline-block;min-width:18pt}',
  'img{vertical-align:baseline}',
].join('');

interface EmitCtx {
  readonly resources: ResourceStore;
  readonly losses: Array<Loss>;
}

function emitBlock(out: Array<string>, el: BodyElement, ctx: EmitCtx): void {
  if (el.kind === 'paragraph') {
    emitParagraph(out, el.paragraph, ctx);
  } else if (el.kind === 'table') {
    emitTable(out, el.table, ctx);
  } else if (el.kind === 'image') {
    emitImageBlock(out, el.image, ctx);
  } else if (el.kind === 'chart') {
    ctx.losses.push({
      severity: 'dropped',
      feature: FEATURES.charts,
      detail: 'chart not rendered by the HTML writer (v0)',
    });
  } else {
    // Shape: the geometry is out of scope, but a text box's content is real
    // document text — emit it as a plain block.
    ctx.losses.push({
      severity: 'dropped',
      feature: FEATURES.shapes,
      detail: 'shape geometry not rendered by the HTML writer (v0)',
    });
    if (el.shape.text) {
      out.push('<div>');
      for (const child of el.shape.text.content) emitBlock(out, child, ctx);
      out.push('</div>');
    }
  }
}

// ---------------------------------------------------------------------------
// Paragraphs
// ---------------------------------------------------------------------------

function emitParagraph(out: Array<string>, p: Paragraph, ctx: EmitCtx): void {
  const resolved = resolveParagraphProperties(p.properties, EMPTY_STYLE_SHEET);
  // Same mapping as the tagged-PDF structure pass: outline level 0–8 → H1–H6.
  const lvl = resolved.outlineLevel;
  const tag = lvl !== undefined && lvl >= 0 && lvl <= 8 ? `h${Math.min(lvl, 5) + 1}` : 'p';

  const style = paragraphCss(resolved);
  const dir = resolved.bidi ? ' dir="rtl"' : '';
  const open = `<${tag}${dir}${style ? ` style="${style}"` : ''}>`;

  const parts: Array<string> = [];
  for (const run of p.runs) parts.push(runHtml(run, p, ctx));
  // An empty paragraph still occupies a line, like in the source document.
  out.push(`${open}${parts.join('') || '&nbsp;'}</${tag}>`);
}

function paragraphCss(r: ResolvedParagraphProperties): string {
  const css: Array<string> = [];
  if (r.alignment === 'center' || r.alignment === 'right') css.push(`text-align:${r.alignment}`);
  else if (r.alignment === 'both' || r.alignment === 'distribute') css.push('text-align:justify');
  if (r.spacingBefore > 0) css.push(`margin-top:${fmt(r.spacingBefore)}pt`);
  if (r.spacingAfter > 0) css.push(`margin-bottom:${fmt(r.spacingAfter)}pt`);
  if (r.indentLeft !== 0) css.push(`margin-left:${fmt(r.indentLeft)}pt`);
  if (r.indentRight !== 0) css.push(`margin-right:${fmt(r.indentRight)}pt`);
  if (r.indentFirstLine !== 0) css.push(`text-indent:${fmt(r.indentFirstLine)}pt`);
  if (r.spacingLine > 0) {
    // §17.3.1.33 w:spacing — rule 'auto' means a multiple of single spacing
    // (240 twips = 12pt = 1.0); exact/atLeast are absolute.
    if (r.spacingLineRule === 'auto') css.push(`line-height:${fmt(r.spacingLine / 12)}`);
    else css.push(`line-height:${fmt(r.spacingLine)}pt`);
  }
  if (r.pageBreakBefore) css.push('break-before:page');
  return css.join(';');
}

// ---------------------------------------------------------------------------
// Runs
// ---------------------------------------------------------------------------

function runHtml(run: Run, p: Paragraph, ctx: EmitCtx): string {
  if (run.math !== undefined) {
    ctx.losses.push({
      severity: 'dropped',
      feature: FEATURES.math,
      detail: 'inline math not rendered by the HTML writer (v0)',
    });
    return '';
  }
  if (run.inlineImage) {
    const img = run.inlineImage;
    const src = dataUri(img.resource, ctx.resources);
    if (!src) return '';
    return `<img src="${src}" style="width:${fmt(img.width)}pt;height:${fmt(img.height)}pt" alt=""/>`;
  }
  if (run.text.length === 0) return '';

  const resolved = resolveRunProperties(run.properties, p.properties, EMPTY_STYLE_SHEET);
  const style = runCss(resolved);
  const dir = resolved.rtl ? ' dir="rtl"' : '';
  let html = `<span${dir}${style ? ` style="${style}"` : ''}>${textHtml(run.text)}</span>`;
  if (resolved.verticalAlign === 'superscript') html = `<sup>${html}</sup>`;
  else if (resolved.verticalAlign === 'subscript') html = `<sub>${html}</sub>`;
  if (run.href !== undefined) {
    // Untrusted input: only allowlisted schemes become clickable (core/links).
    const safe = sanitizeHref(run.href);
    if (safe !== undefined) {
      html = `<a href="${escapeAttr(safe)}">${html}</a>`;
    } else {
      ctx.losses.push({
        severity: 'degraded',
        feature: FEATURES.hyperlinks,
        detail: `hyperlink target with a disallowed scheme rendered as plain text`,
      });
    }
  }
  return html;
}

function runCss(r: ResolvedRunProperties): string {
  const css: Array<string> = [];
  const family = r.fontFamily.ascii;
  if (family) css.push(`font-family:${JSON.stringify(family)}`);
  css.push(`font-size:${fmt(r.fontSizePt)}pt`);
  if (r.bold) css.push('font-weight:700');
  if (r.italic) css.push('font-style:italic');
  const deco: Array<string> = [];
  const underlined = r.underline !== 'none';
  if (underlined) deco.push('underline');
  if (r.strike) deco.push('line-through');
  if (deco.length > 0) {
    css.push(`text-decoration-line:${deco.join(' ')}`);
    const style = underlined ? decorationStyle(r.underline) : undefined;
    if (style) css.push(`text-decoration-style:${style}`);
  }
  if (r.colorHex !== '000000') css.push(`color:#${r.colorHex}`);
  return css.join(';');
}

function decorationStyle(u: ResolvedRunProperties['underline'] & string): string | undefined {
  switch (u) {
    case 'double':
      return 'double';
    case 'dotted':
    case 'dottedHeavy':
      return 'dotted';
    case 'dash':
    case 'dashHeavy':
      return 'dashed';
    case 'wave':
      return 'wavy';
    default:
      return undefined; // single/thick → solid (the default)
  }
}

// Escape, then map the control characters the model uses: '\t' (list markers,
// w:tab) → a fixed gap, '\n' (w:br) → a line break.
function textHtml(text: string): string {
  return escapeText(text).replaceAll('\t', '<span class="tab"></span>').replaceAll('\n', '<br/>');
}

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

function emitTable(out: Array<string>, table: Table, ctx: EmitCtx): void {
  const css: Array<string> = [];
  if (table.properties.widthPt !== undefined) css.push(`width:${fmt(table.properties.widthPt)}pt`);
  else if (table.properties.widthFraction !== undefined) {
    css.push(`width:${fmt(table.properties.widthFraction * 100)}%`);
  }
  if (table.properties.alignment === 'center') css.push('margin-left:auto;margin-right:auto');
  else if (table.properties.alignment === 'right') css.push('margin-left:auto');
  out.push(`<table${css.length > 0 ? ` style="${css.join(';')}"` : ''}>`);

  if (table.grid.length > 0) {
    out.push('<colgroup>');
    for (const w of table.grid) out.push(`<col style="width:${fmt(w)}pt"/>`);
    out.push('</colgroup>');
  }

  // Grid-column starts per cell (cells carry no explicit position; vertical
  // merges are resolved by column like the layout's rowSpan derivation).
  const colStarts = table.rows.map((row) => {
    let col = 0;
    return row.cells.map((cell) => {
      const start = col;
      col += cell.properties.colSpan ?? 1;
      return start;
    });
  });

  for (let ri = 0; ri < table.rows.length; ri++) {
    const row = table.rows[ri]!;
    out.push('<tr>');
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci]!;
      const merge = cell.properties.merge;
      // Continuation cells are covered by the start cell's rowspan.
      if (merge === 'middle' || merge === 'end') continue;
      const rowSpan =
        merge === 'start' ? mergeRowSpan(table, colStarts, ri, colStarts[ri]![ci]!) : 1;
      emitCell(out, cell, table, ctx, {
        isHeader: row.properties.isHeader === true,
        firstRow: ri === 0,
        lastRow: ri === table.rows.length - 1,
        firstCol: colStarts[ri]![ci]! === 0,
        lastCol: colStarts[ri]![ci]! + (cell.properties.colSpan ?? 1) >= table.grid.length,
        rowSpan,
      });
    }
    out.push('</tr>');
  }
  out.push('</table>');
}

function mergeRowSpan(
  table: Table,
  colStarts: ReadonlyArray<ReadonlyArray<number>>,
  startRow: number,
  colStart: number,
): number {
  let span = 1;
  for (let r = startRow + 1; r < table.rows.length; r++) {
    const row = table.rows[r]!;
    const idx = colStarts[r]!.findIndex((c) => c === colStart);
    const cell = idx >= 0 ? row.cells[idx] : undefined;
    const merge = cell?.properties.merge;
    if (merge !== 'middle' && merge !== 'end') break;
    span++;
    if (merge === 'end') break;
  }
  return span;
}

interface CellPos {
  readonly isHeader: boolean;
  readonly firstRow: boolean;
  readonly lastRow: boolean;
  readonly firstCol: boolean;
  readonly lastCol: boolean;
  readonly rowSpan: number;
}

function emitCell(
  out: Array<string>,
  cell: TableCell,
  table: Table,
  ctx: EmitCtx,
  pos: CellPos,
): void {
  const tag = pos.isHeader ? 'th' : 'td';
  const attrs: Array<string> = [];
  const colSpan = cell.properties.colSpan ?? 1;
  if (colSpan > 1) attrs.push(`colspan="${colSpan}"`);
  if (pos.rowSpan > 1) attrs.push(`rowspan="${pos.rowSpan}"`);

  const css: Array<string> = [];
  // Each cell takes its own border, inheriting the table default (outer side
  // on edge cells, insideH/insideV between cells). The layout's §17.4
  // heavier-border conflict pass is page-rendering territory; with
  // border-collapse the browser performs its own conflict resolution.
  const t = table.properties.borders;
  const c = cell.properties.borders;
  pushBorder(css, 'top', c?.top ?? (pos.firstRow ? t?.top : t?.insideH));
  pushBorder(css, 'bottom', c?.bottom ?? (pos.lastRow ? t?.bottom : t?.insideH));
  pushBorder(css, 'left', c?.left ?? (pos.firstCol ? t?.left : t?.insideV));
  pushBorder(css, 'right', c?.right ?? (pos.lastCol ? t?.right : t?.insideV));
  if (cell.properties.shading) css.push(`background-color:#${cell.properties.shading.colorHex}`);
  const margins = cell.properties.margins ?? table.properties.defaultCellMargins;
  // Word's default cell padding (108 twips = 5.4pt left/right) keeps text off
  // the rules even when the document does not specify margins.
  const padTop = margins?.top ?? 0;
  const padRight = margins?.right ?? 5.4;
  const padBottom = margins?.bottom ?? 0;
  const padLeft = margins?.left ?? 5.4;
  css.push(`padding:${fmt(padTop)}pt ${fmt(padRight)}pt ${fmt(padBottom)}pt ${fmt(padLeft)}pt`);

  out.push(`<${tag}${attrs.length > 0 ? ` ${attrs.join(' ')}` : ''} style="${css.join(';')}">`);
  for (const child of cell.content) emitBlock(out, child, ctx);
  out.push(`</${tag}>`);
}

function pushBorder(css: Array<string>, side: string, border: Border | undefined): void {
  if (!border || border.style === 'none') return;
  const width = border.width ?? 0.5;
  const style =
    border.style === 'double'
      ? 'double'
      : border.style === 'dotted'
        ? 'dotted'
        : border.style === 'dashed'
          ? 'dashed'
          : 'solid'; // single + thick
  css.push(`border-${side}:${fmt(width)}pt ${style} #${border.colorHex ?? '000000'}`);
}

// ---------------------------------------------------------------------------
// Images
// ---------------------------------------------------------------------------

function emitImageBlock(out: Array<string>, image: ImageBlock, ctx: EmitCtx): void {
  const src = dataUri(image.resource, ctx.resources);
  if (!src) return; // unresolved/unsupported resource: the block has no content
  const pp = image.paragraphProperties;
  const align = pp.alignment === 'center' || pp.alignment === 'right' ? pp.alignment : undefined;
  const css: Array<string> = [];
  if (align) css.push(`text-align:${align}`);
  if (pp.spacingBefore !== undefined && pp.spacingBefore > 0) {
    css.push(`margin-top:${fmt(pp.spacingBefore)}pt`);
  }
  if (pp.spacingAfter !== undefined && pp.spacingAfter > 0) {
    css.push(`margin-bottom:${fmt(pp.spacingAfter)}pt`);
  }
  const alt = image.altText ?? '';
  out.push(
    `<figure${css.length > 0 ? ` style="${css.join(';')}"` : ''}>` +
      `<img src="${src}" style="width:${fmt(image.width)}pt;height:${fmt(image.height)}pt" alt="${escapeAttr(alt)}"/>` +
      `</figure>`,
  );
}

const MIME_BY_FORMAT = {
  jpeg: 'image/jpeg',
  png: 'image/png',
  jpeg2000: 'image/jp2',
} as const;

function dataUri(resource: ResourceId | undefined, store: ResourceStore): string | undefined {
  if (resource === undefined) return undefined;
  const bytes = store.get(resource);
  if (!bytes) return undefined;
  const format = detectImageFormat(bytes);
  if (!format) return undefined;
  return `data:${MIME_BY_FORMAT[format]};base64,${toBase64(bytes)}`;
}

// ---------------------------------------------------------------------------
// Plumbing
// ---------------------------------------------------------------------------

// First section's content width — the column the document was authored for.
// Falls back to the layout engine's A4 + 1" margins.
function contentWidthPt(flow: FlowDoc): number {
  const props = flow.sections[0]?.properties ?? flow.section;
  const pageWidth = props?.pageSize?.width ?? 595;
  const left = props?.margins?.left ?? 72;
  const right = props?.margins?.right ?? 72;
  return Math.max(72, pageWidth - left - right);
}

function escapeText(s: string): string {
  return s.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}

function escapeAttr(s: string): string {
  return escapeText(s).replaceAll('"', '&quot;');
}

function fmt(n: number): string {
  const r = Math.round(n * 100) / 100;
  return Object.is(r, -0) ? '0' : String(r);
}
