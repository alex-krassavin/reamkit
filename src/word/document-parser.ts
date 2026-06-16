// ECMA-376 Part 1 §17 — WordprocessingML document parser.
//
// Produces BodyElement[] preserving the original interleaving of paragraphs
// and tables in the body. Uses fast-xml-parser in preserveOrder mode for
// traversal; for per-element property extraction we adapt PO subtrees to the
// flat shape the rPr/pPr parsers consume (po-to-flat).

import { XMLParser } from 'fast-xml-parser';

import type {
  BodyElement,
  HeaderFooterReference,
  HeaderFooterType,
  InlineImage,
  PageMargins,
  PageSize,
  Paragraph,
  Run,
  Section,
  SectionColumns,
  SectionProperties,
} from '@/core/document-model';

import type { ColorResolver } from '@/core/drawingml/colors';
import type { Pt, ResourceId } from '@/core/ir';
import type { PoNode } from '@/core/po-helpers';
import { emuToPt, twipsToPt } from '@/core/ir';
import { parseOMath } from '@/word/omml-parser';
import { defaultColorResolver } from '@/core/drawingml/colors';
import { expandMcChildren, parseDrawing, parseVmlPicture } from '@/word/drawing-parser';
import { diagramTransform, parseDiagramDrawing } from '@/pptx/slide-parser';
import { parseParagraphProperties } from '@/word/paragraph-properties';
import {
  poAttr,
  poChildren,
  poFindByPath,
  poFindDescendant,
  poIntAttr,
  poIs,
  poText,
} from '@/core/po-helpers';
import { poElementToFlat } from '@/word/po-to-flat';
import { parseRunProperties } from '@/word/run-properties';
import { parseTable } from '@/word/table-parser';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
});

const RUN_CONTAINER_TAGS = new Set([
  'w:hyperlink',
  'w:sdt',
  'w:sdtContent',
  'w:smartTag',
  'w:fldSimple',
  // Tracked changes — "accept all / final document" semantics: descend into
  // inserted (§17.13.5.18) and moved-in (§17.13.5.31) runs so their text
  // renders. w:del / w:moveFrom are deliberately NOT here, so deleted /
  // moved-out runs are dropped (their text lives in w:delText, not w:t).
  'w:ins',
  'w:moveTo',
]);

// Resolves a drawing relationship id to a content-addressed ResourceId —
// supplied by the converter (which owns the OPC package and the ResourceStore).
export type ImageResolver = (relId: string) => ResourceId | undefined;
export type HyperlinkResolver = (relId: string) => string | undefined;

// Document-wide resolvers every nested parser needs — one context object
// instead of threading a parameter pair through ten signatures (oop-design §8).
export interface ParseContext {
  readonly resolveColor: ColorResolver;
  readonly resolveImage?: ImageResolver;
  // §17.16.22 w:hyperlink r:id → external target URL from the owning part's
  // rels (TargetMode="External" only). Absent ⇒ links unwrap to plain text.
  readonly resolveHyperlink?: HyperlinkResolver;
  // SmartArt: a data relationship id (dgm:relIds @r:dm) → the diagram's
  // pre-rendered drawing override (its dsp:spTree), or undefined when the file
  // ships none (E-SMARTART SA2).
  readonly resolveDiagram?: (relId: string) => PoNode | undefined;
}

export const DEFAULT_PARSE_CONTEXT: ParseContext = { resolveColor: defaultColorResolver };

export function parseDocument(
  documentXml: Uint8Array,
  ctx: ParseContext = DEFAULT_PARSE_CONTEXT,
): Array<BodyElement> {
  const xml = decoder.decode(documentXml);
  const tree = parser.parse(xml) as Array<PoNode>;
  const body = poFindByPath(tree, ['w:document', 'w:body']);
  if (!body) return [];
  return parseBodyElements(poChildren(body), ctx);
}

const HF_TYPES = new Set<HeaderFooterType>(['default', 'first', 'even']);

export const EMPTY_SECTION: SectionProperties = {
  headers: [],
  footers: [],
};

// ECMA-376 Part 1 §17.6.17 — sectPr lives as the last child of w:body and
// describes the (final) section of the document. Returned as a single
// SectionProperties for backward compatibility; for multi-section documents
// the renderer reads parseSections instead.
export function parseSection(documentXml: Uint8Array): SectionProperties {
  const sections = parseSections(documentXml);
  if (sections.length === 0) return EMPTY_SECTION;
  // Use the final section as the document-wide fallback; it usually carries
  // pgSz/pgMar even when intermediate sections only override headers.
  return sections[sections.length - 1]!.properties;
}

// ECMA-376 §17.6 — collect every sectPr (one per intermediate paragraph plus
// the body-final one). Each Section carries the endIndex (exclusive) into the
// body element list: section i applies to body[sections[i-1].endIndex..endIndex).
// A document with no sectPr at all returns a single empty section spanning
// the whole body.
export function parseSections(documentXml: Uint8Array): Array<Section> {
  const xml = decoder.decode(documentXml);
  const tree = parser.parse(xml) as Array<PoNode>;
  const body = poFindByPath(tree, ['w:document', 'w:body']);
  if (!body) return [];

  const children = poChildren(body);
  const sections: Array<Section> = [];
  let bodyIdx = 0;

  for (const child of children) {
    if (poIs(child, 'w:sectPr')) {
      // Final body-level sectPr: applies to remaining body elements.
      sections.push({ properties: parseSectPrNode(child), endIndex: bodyIdx });
      continue;
    }
    if (poIs(child, 'w:p')) {
      // Mid-document section break: sectPr inside pPr ends the section at the
      // *end* of this paragraph (paragraph belongs to the closing section).
      const pPrNode = poChildren(child).find((c) => poIs(c, 'w:pPr'));
      const sectPrInPPr = pPrNode
        ? poChildren(pPrNode).find((c) => poIs(c, 'w:sectPr'))
        : undefined;
      // tryExtractImageFromParagraph and parseTable count toward bodyIdx
      // identically — we mirror parseBodyElements' "one BodyElement per w:p
      // or w:tbl" semantics here.
      bodyIdx++;
      if (sectPrInPPr) {
        sections.push({ properties: parseSectPrNode(sectPrInPPr), endIndex: bodyIdx });
      }
    } else if (poIs(child, 'w:tbl')) {
      bodyIdx++;
    }
  }

  if (sections.length === 0 || sections[sections.length - 1]!.endIndex < bodyIdx) {
    sections.push({ properties: EMPTY_SECTION, endIndex: bodyIdx });
  }
  return sections;
}

function parseSectPrNode(sectPr: PoNode): SectionProperties {
  let pageSize: PageSize | undefined;
  let margins: PageMargins | undefined;
  let titlePg = false;
  let columns: SectionColumns | undefined;
  const headers: Array<HeaderFooterReference> = [];
  const footers: Array<HeaderFooterReference> = [];

  for (const child of poChildren(sectPr)) {
    if (poIs(child, 'w:pgSz')) {
      const w = poIntAttr(child, 'w');
      const h = poIntAttr(child, 'h');
      const orientRaw = poAttr(child, 'orient');
      if (w !== undefined && h !== undefined) {
        pageSize = {
          width: twipsToPt(w),
          height: twipsToPt(h),
          ...(orientRaw === 'portrait' || orientRaw === 'landscape'
            ? { orientation: orientRaw }
            : {}),
        };
      }
    } else if (poIs(child, 'w:pgMar')) {
      const top = poIntAttr(child, 'top');
      const right = poIntAttr(child, 'right');
      const bottom = poIntAttr(child, 'bottom');
      const left = poIntAttr(child, 'left');
      const header = poIntAttr(child, 'header');
      const footer = poIntAttr(child, 'footer');
      margins = {
        top: twipsToPt(top ?? 1440),
        right: twipsToPt(right ?? 1440),
        bottom: twipsToPt(bottom ?? 1440),
        left: twipsToPt(left ?? 1440),
        ...(header !== undefined ? { header: twipsToPt(header) } : {}),
        ...(footer !== undefined ? { footer: twipsToPt(footer) } : {}),
      };
    } else if (poIs(child, 'w:headerReference')) {
      pushHeaderFooter(child, headers);
    } else if (poIs(child, 'w:footerReference')) {
      pushHeaderFooter(child, footers);
    } else if (poIs(child, 'w:titlePg')) {
      const val = poAttr(child, 'val');
      titlePg = val === undefined || val === '' || (val !== '0' && val !== 'false');
    } else if (poIs(child, 'w:cols')) {
      columns = parseColumns(child);
    }
  }

  return {
    ...(pageSize ? { pageSize } : {}),
    ...(margins ? { margins } : {}),
    headers,
    footers,
    ...(titlePg ? { titlePg: true } : {}),
    ...(columns ? { columns } : {}),
  };
}

// §17.6.4 w:cols: @w:num equal-width columns separated by @w:space, OR
// explicit w:col children each with their own width/trailing space.
function parseColumns(cols: PoNode): SectionColumns | undefined {
  const explicit: Array<{ widthPt: number; spacePt: number }> = [];
  for (const col of poChildren(cols)) {
    if (!poIs(col, 'w:col')) continue;
    const w = poIntAttr(col, 'w');
    if (w === undefined) continue;
    explicit.push({ widthPt: twipsToPt(w), spacePt: twipsToPt(poIntAttr(col, 'space') ?? 0) });
  }
  const num = poIntAttr(cols, 'num');
  const count = explicit.length > 0 ? explicit.length : (num ?? 1);
  if (count <= 1) return undefined;
  return {
    count,
    spacePt: twipsToPt(poIntAttr(cols, 'space') ?? 720),
    ...(explicit.length > 0 ? { explicit } : {}),
  };
}

function pushHeaderFooter(node: PoNode, list: Array<HeaderFooterReference>): void {
  const rId = poAttr(node, 'id');
  if (!rId) return;
  const typeRaw = poAttr(node, 'type') ?? 'default';
  const type: HeaderFooterType = HF_TYPES.has(typeRaw as HeaderFooterType)
    ? (typeRaw as HeaderFooterType)
    : 'default';
  list.push({ type, relationshipId: rId });
}

// Parses word/header*.xml or word/footer*.xml. The root is w:hdr or w:ftr,
// and its children are the same body-element shape as the main document.
export function parseHeaderFooter(
  xml: Uint8Array,
  ctx: ParseContext = DEFAULT_PARSE_CONTEXT,
): Array<BodyElement> {
  const tree = parser.parse(decoder.decode(xml)) as Array<PoNode>;
  const root = tree.find((n) => poIs(n, 'w:hdr') || poIs(n, 'w:ftr'));
  if (!root) return [];
  return parseBodyElements(poChildren(root), ctx);
}

export function parseBodyElements(
  children: ReadonlyArray<PoNode>,
  ctx: ParseContext = DEFAULT_PARSE_CONTEXT,
): Array<BodyElement> {
  const out: Array<BodyElement> = [];
  // Body-level w:bookmarkStart (between block elements) anchors to the NEXT
  // paragraph.
  let pendingBookmarks: Array<string> | undefined;
  for (const child of children) {
    if (poIs(child, 'w:bookmarkStart')) {
      const bookmarkName = poAttr(child, 'name');
      if (bookmarkName !== undefined && bookmarkName !== '' && bookmarkName !== '_GoBack') {
        (pendingBookmarks ??= []).push(bookmarkName);
      }
    } else if (poIs(child, 'w:p')) {
      const drawings = tryExtractDrawingFromParagraph(child, ctx);
      if (drawings) {
        out.push(...drawings);
      } else {
        out.push({ kind: 'paragraph', paragraph: parseParagraph(child, ctx, pendingBookmarks) });
        pendingBookmarks = undefined;
      }
    } else if (poIs(child, 'w:tbl')) {
      out.push({ kind: 'table', table: parseTable(child, ctx) });
    } else if (poIs(child, 'w:sdt')) {
      // §17.5.2 block-level structured document tag (content control): the
      // wrapper is chrome — its sdtContent children are ordinary body flow.
      const content = poChildren(child).find((c) => poIs(c, 'w:sdtContent'));
      if (content) out.push(...parseBodyElements(poChildren(content), ctx));
    }
  }
  return out;
}

// ECMA-376 Part 1 §17.3.3 (drawing) + §20 (DrawingML).
// A paragraph containing ONLY a w:drawing (no text alongside it) collapses to
// a standalone block: an ImageBlock for an embedded picture, or a ShapeBlock
// for a wps:wsp shape. Mixed text+drawing paragraphs keep a picture on the run
// via Run.inlineImage (shapes in mixed runs are dropped in M5) and are emitted
// as paragraphs by parseBodyElements.
// Run wrappers a lone picture can hide inside that the writer flattens away:
// tracked-change insals/moves (§17.13.5) and content controls / smart tags
// (§17.5.2). A re-read of the writer's output sees the bare <w:drawing>, so to
// keep paragraph counts stable across the round-trip these must be transparent
// to the standalone-image collapse too. w:hyperlink and w:fldSimple are
// deliberately excluded — the writer preserves a link / field on an inline
// image, so such a paragraph must stay a paragraph, not collapse to a block.
const COLLAPSE_TRANSPARENT_TAGS = new Set([
  'w:ins',
  'w:moveTo',
  'w:sdt',
  'w:sdtContent',
  'w:smartTag',
]);

interface LoneDrawingScan {
  drawing?: PoNode;
  vml?: PoNode;
  hasOther: boolean;
}

// Scan a paragraph (descending through the transparent wrappers above) for a
// single drawing / VML picture with no sibling content. hasOther trips on real
// text, math, fields or hyperlinks — anything the writer would keep, which
// means the paragraph must not collapse to a standalone image block.
function scanForLoneDrawing(container: PoNode, acc: LoneDrawingScan): void {
  for (const child of expandMcChildren(poChildren(container))) {
    if (poIs(child, 'w:pPr')) continue;
    if (poIs(child, 'w:r')) {
      for (const rc of expandMcChildren(poChildren(child))) {
        if (poIs(rc, 'w:drawing')) {
          if (!acc.drawing) acc.drawing = rc;
        } else if (poIs(rc, 'w:pict') || poIs(rc, 'w:object')) {
          // Only a VML node that actually bears a picture (<v:imagedata>) is a
          // candidate — an empty frame or a bare ActiveX/OLE control object is
          // ignored, so a paragraph that pairs a picture run with a control run
          // still collapses on the picture.
          if (!acc.vml && poFindDescendant(rc, 'v:imagedata') !== undefined) acc.vml = rc;
        } else if (poIs(rc, 'w:t') && poText(rc).length > 0) {
          acc.hasOther = true;
        } else if (poIs(rc, 'w:tab') || poIs(rc, 'w:br') || poIs(rc, 'w:noBreakHyphen')) {
          acc.hasOther = true;
        }
      }
      continue;
    }
    if (poIs(child, 'm:oMath') || poIs(child, 'm:oMathPara')) {
      acc.hasOther = true;
      continue;
    }
    const tag = elementTag(child);
    if (tag && COLLAPSE_TRANSPARENT_TAGS.has(tag)) {
      scanForLoneDrawing(child, acc);
    } else if (tag && RUN_CONTAINER_TAGS.has(tag)) {
      // A hyperlink / field wrapping the picture: keep the paragraph (the
      // writer preserves the link or field on the inline image).
      acc.hasOther = true;
    }
  }
}

function tryExtractDrawingFromParagraph(p: PoNode, ctx: ParseContext): Array<BodyElement> | null {
  const scan: LoneDrawingScan = { hasOther: false };
  scanForLoneDrawing(p, scan);
  const { drawing, vml } = scan;
  if (scan.hasOther || (!drawing && !vml)) return null;

  // Inject parseBodyElements (bound to this context) so a shape's text box is
  // parsed without a module cycle. A modern <w:drawing> takes precedence over a
  // legacy <w:pict>/<w:object> VML picture (§14). Collapsing a lone picture to
  // a standalone block here is what keeps the round-trip symmetric: the writer
  // re-emits a block image as its own lone-drawing paragraph, which a re-read
  // collapses again — so the FIRST read must collapse too, or paragraph counts
  // drift by one on every standalone VML image.
  const fromVml = drawing === undefined;
  const parseBody = (children: ReadonlyArray<PoNode>): Array<BodyElement> =>
    parseBodyElements(children, ctx);
  const content = fromVml
    ? parseVmlPicture(vml!)
    : parseDrawing(drawing, ctx.resolveColor, parseBody);
  if (!content) return null;
  // A dangling VML <v:imagedata r:id> (referenced media absent from the
  // package) carries nothing to render; skip it so the paragraph stays empty
  // on both read passes rather than materialising an un-writable phantom.
  if (fromVml && content.kind === 'image' && ctx.resolveImage?.(content.imageId) === undefined) {
    return null;
  }

  const pPrNode = poChildren(p).find((c) => poIs(c, 'w:pPr'));
  const paragraphProperties = pPrNode ? parseParagraphProperties(poElementToFlat(pPrNode)) : {};

  if (content.kind === 'image') {
    const resource = ctx.resolveImage?.(content.imageId);
    return [
      {
        kind: 'image',
        image: {
          ...(resource ? { resource } : {}),
          width: content.width,
          height: content.height,
          paragraphProperties,
          ...(content.altText ? { altText: content.altText } : {}),
          ...(content.float ? { float: content.float } : {}),
        },
      },
    ];
  }
  if (content.kind === 'chart') {
    return [
      {
        kind: 'chart',
        chart: {
          chartRelId: content.chartRelId,
          width: content.width,
          height: content.height,
          paragraphProperties,
          ...(content.altText ? { altText: content.altText } : {}),
          ...(content.float ? { float: content.float } : {}),
        },
      },
    ];
  }
  if (content.kind === 'diagram') {
    // SmartArt: resolve the drawing override and render its nodes as floating
    // shapes anchored to the paragraph's column origin (E-SMARTART SA2). No
    // override ⇒ keep the (empty) paragraph, byte-stable.
    const spTree = ctx.resolveDiagram?.(content.dmRelId);
    if (!spTree) return null;
    const frame = { x: 0, y: 0, cx: content.widthEmu, cy: content.heightEmu };
    const shapes = parseDiagramDrawing(
      spTree,
      diagramTransform(spTree, frame),
      (box) => ({
        wrap: 'none',
        posH: { relativeFrom: 'column', offsetPt: emuToPt(box.x) },
        posV: { relativeFrom: 'paragraph', offsetPt: emuToPt(box.y) },
      }),
      ctx.resolveColor,
      undefined,
    );
    return shapes.length > 0 ? shapes.map((shape) => ({ kind: 'shape', shape })) : null;
  }
  return [
    {
      kind: 'shape',
      shape: {
        ...content.data,
        paragraphProperties,
        ...(content.altText ? { altText: content.altText } : {}),
        ...(content.float ? { float: content.float } : {}),
      },
    },
  ];
}

function parseParagraph(p: PoNode, ctx: ParseContext, extraBookmarks?: Array<string>): Paragraph {
  // §17.13.6.2 — bookmarks opening in this paragraph (plus any the caller
  // carried over from between-paragraph positions). The hidden _GoBack
  // edit-cursor bookmark is noise in every Word save — skipped.
  const bookmarks: Array<string> = [...(extraBookmarks ?? [])];
  for (const child of poChildren(p)) {
    if (!poIs(child, 'w:bookmarkStart')) continue;
    const bookmarkName = poAttr(child, 'name');
    if (bookmarkName !== undefined && bookmarkName !== '' && bookmarkName !== '_GoBack') {
      bookmarks.push(bookmarkName);
    }
  }
  const pPr = poChildren(p).find((c) => poIs(c, 'w:pPr'));
  let properties = parseParagraphProperties(pPr ? poElementToFlat(pPr) : undefined);
  // A display equation (m:oMathPara) centres its paragraph by default
  // (m:oMathParaPr/m:jc may override). Only applied when the paragraph has no
  // explicit alignment of its own.
  const mathPara = poChildren(p).find((c) => poIs(c, 'm:oMathPara'));
  if (mathPara && properties.alignment === undefined) {
    const paraPr = poChildren(mathPara).find((c) => poIs(c, 'm:oMathParaPr'));
    const jcNode = paraPr ? poChildren(paraPr).find((c) => poIs(c, 'm:jc')) : undefined;
    const jcVal = poAttr(jcNode, 'val');
    const alignment = jcVal === 'left' ? 'left' : jcVal === 'right' ? 'right' : 'center';
    properties = { ...properties, alignment };
  }
  const collected: Array<CollectedRun> = [];
  collectRuns(p, collected, ctx);
  return {
    properties,
    runs: applyFieldFsm(collected),
    ...(bookmarks.length > 0 ? { bookmarks } : {}),
  };
}

// A parsed run plus the complex-field markers the FSM consumes (§17.16.18
// w:fldChar / w:instrText). Internal to run collection.
interface CollectedRun {
  readonly run: Run;
  readonly fldChar?: 'begin' | 'separate' | 'end';
  readonly instrText?: string;
}

// §17.16.5.35 PAGE / §17.16.5.33 NUMPAGES: the instruction's first keyword;
// switches (\* MERGEFORMAT …) are ignored. Anything else stays a cached
// result (REF, TOC, DATE, … render their stored text exactly as before).
function parseFieldInstr(instr: string | undefined): 'PAGE' | 'NUMPAGES' | undefined {
  if (!instr) return undefined;
  const m = /^\s*([A-Za-z]+)/.exec(instr);
  const kw = m?.[1]?.toUpperCase();
  return kw === 'PAGE' ? 'PAGE' : kw === 'NUMPAGES' ? 'NUMPAGES' : undefined;
}

// Fold a recognized field's cached-result runs into ONE field run: the cached
// text concatenated (the per-page substitution replaces it wholesale), the
// first result run's formatting, any hyperlink carried along.
function synthesizeFieldRun(
  result: ReadonlyArray<Run>,
  field: 'PAGE' | 'NUMPAGES',
  href?: string,
): Run {
  const first = result[0];
  const linked = href ?? result.find((r) => r.href !== undefined)?.href;
  return {
    text: result.map((r) => r.text).join(''),
    properties: first?.properties ?? {},
    field,
    ...(linked !== undefined ? { href: linked } : {}),
  };
}

// §17.16.18 complex fields: begin → instrText* → separate → cached result →
// end, spread across sibling runs. Recognized PAGE/NUMPAGES collapse to one
// field run; everything else keeps its cached result exactly as before (the
// zero-glyph marker runs were never rendered, so dropping them is inert).
function applyFieldFsm(collected: ReadonlyArray<CollectedRun>): Array<Run> {
  const out: Array<Run> = [];
  let st: { phase: 'instr' | 'result'; instr: string; result: Array<Run>; depth: number } | null =
    null;
  for (const c of collected) {
    if (c.fldChar === 'begin') {
      if (st) {
        if (st.phase === 'instr') st.depth++;
        else {
          // A new field opening inside a result: flush what we have and track
          // the new one (nested result fields are rare; keep it simple).
          out.push(...st.result);
          st = { phase: 'instr', instr: '', result: [], depth: 0 };
        }
      } else {
        st = { phase: 'instr', instr: '', result: [], depth: 0 };
      }
      continue;
    }
    if (!st) {
      out.push(c.run);
      continue;
    }
    if (c.fldChar === 'separate') {
      if (st.depth === 0) st.phase = 'result';
      continue;
    }
    if (c.fldChar === 'end') {
      if (st.depth > 0) {
        st.depth--;
        continue;
      }
      const field = parseFieldInstr(st.instr);
      if (field) out.push(synthesizeFieldRun(st.result, field));
      else out.push(...st.result);
      st = null;
      continue;
    }
    if (st.phase === 'instr') {
      if (c.instrText !== undefined) st.instr += c.instrText;
      continue;
    }
    st.result.push(c.run);
  }
  if (st) out.push(...st.result); // unterminated field: keep the visible part
  return out;
}

function collectRuns(
  container: PoNode,
  out: Array<CollectedRun>,
  ctx: ParseContext,
  href?: string,
  anchor?: string,
): void {
  for (const child of poChildren(container)) {
    if (poIs(child, 'w:pPr')) continue;
    if (poIs(child, 'w:r')) {
      const parsed = parseRun(child, ctx);
      const run =
        href !== undefined || anchor !== undefined
          ? {
              ...parsed.run,
              ...(href !== undefined ? { href } : {}),
              ...(anchor !== undefined ? { anchor } : {}),
            }
          : parsed.run;
      out.push({
        run,
        ...(parsed.fldChar ? { fldChar: parsed.fldChar } : {}),
        ...(parsed.instrText !== undefined ? { instrText: parsed.instrText } : {}),
      });
      continue;
    }
    // OfficeMath: an inline equation (m:oMath) or a display paragraph
    // (m:oMathPara, holding one or more m:oMath) → math runs.
    if (poIs(child, 'm:oMath')) {
      out.push({ run: { text: '', properties: {}, math: parseOMath(child) } });
      continue;
    }
    if (poIs(child, 'm:oMathPara')) {
      for (const om of poChildren(child)) {
        if (poIs(om, 'm:oMath')) {
          out.push({ run: { text: '', properties: {}, math: parseOMath(om) } });
        }
      }
      continue;
    }
    const tag = elementTag(child);
    if (tag === 'w:fldSimple') {
      // §17.16.19 — the instruction is an attribute, the children are the
      // cached result. PAGE/NUMPAGES collapse to one field run; anything else
      // keeps its cached runs (the old unwrap behavior).
      const field = parseFieldInstr(poAttr(child, 'instr'));
      if (field) {
        const inner: Array<CollectedRun> = [];
        collectRuns(child, inner, ctx, href, anchor);
        out.push({ run: synthesizeFieldRun(applyFieldFsm(inner), field, href) });
        continue;
      }
      collectRuns(child, out, ctx, href, anchor);
      continue;
    }
    if (tag && RUN_CONTAINER_TAGS.has(tag)) {
      // A hyperlink container stamps its target onto every run inside (nested
      // containers inherit the outer link): @r:id resolves to an external URL,
      // @w:anchor names a bookmark in this document (§17.16.22).
      let childHref = href;
      let childAnchor = anchor;
      if (tag === 'w:hyperlink') {
        const rId = poAttr(child, 'id');
        const resolved = rId ? ctx.resolveHyperlink?.(rId) : undefined;
        if (resolved !== undefined) {
          childHref = resolved;
        } else {
          const bookmark = poAttr(child, 'anchor');
          if (bookmark !== undefined && bookmark !== '') childAnchor = bookmark;
        }
      }
      collectRuns(child, out, ctx, childHref, childAnchor);
    }
  }
}

function parseRun(
  r: PoNode,
  ctx: ParseContext,
): { run: Run; fldChar?: 'begin' | 'separate' | 'end'; instrText?: string } {
  const rPr = poChildren(r).find((c) => poIs(c, 'w:rPr'));
  const properties = parseRunProperties(rPr ? poElementToFlat(rPr) : undefined);
  let text = '';
  let pageBreak = false;
  let inlineImage: InlineImage | undefined;
  let fldChar: 'begin' | 'separate' | 'end' | undefined;
  let instrText: string | undefined;
  let footnoteRef: string | undefined;
  let endnoteRef: string | undefined;
  let noteNumber = false;
  for (const child of expandMcChildren(poChildren(r))) {
    if (poIs(child, 'w:rPr')) continue;
    if (poIs(child, 'w:fldChar')) {
      const t = poAttr(child, 'fldCharType');
      if (t === 'begin' || t === 'separate' || t === 'end') fldChar = t;
      continue;
    }
    if (poIs(child, 'w:instrText')) {
      instrText = (instrText ?? '') + poText(child);
      continue;
    }
    if (poIs(child, 'w:footnoteReference')) {
      const id = poAttr(child, 'id');
      if (id !== undefined) footnoteRef = id;
      continue;
    }
    if (poIs(child, 'w:endnoteReference')) {
      const id = poAttr(child, 'id');
      if (id !== undefined) endnoteRef = id;
      continue;
    }
    if (poIs(child, 'w:footnoteRef') || poIs(child, 'w:endnoteRef')) {
      noteNumber = true;
      continue;
    }
    if (poIs(child, 'w:t')) {
      text += poText(child);
    } else if (poIs(child, 'w:tab')) {
      text += '\t';
    } else if (poIs(child, 'w:br')) {
      // §17.3.3.1 — w:type="page" forces a page break; any other break type
      // (textWrapping/column/none) is a soft line break within the text flow.
      if (poAttr(child, 'type') === 'page') pageBreak = true;
      else text += '\n';
    } else if (poIs(child, 'w:noBreakHyphen')) {
      text += '‑';
    } else if (poIs(child, 'w:softHyphen')) {
      text += '­';
    } else if (poIs(child, 'w:drawing')) {
      // Only pictures render inline in M5; a wps shape in a mixed run is
      // dropped (its text is preserved). Colour is irrelevant for pictures,
      // so this deliberately does NOT take ctx.resolveColor (byte-parity with
      // the pre-context code; revisit if inline shapes ever render).
      const content = parseDrawing(child, defaultColorResolver);
      if (content && content.kind === 'image') {
        const resource = ctx.resolveImage?.(content.imageId);
        inlineImage = {
          ...(resource ? { resource } : {}),
          width: content.width,
          height: content.height,
        };
      }
    } else if (poIs(child, 'w:pict') || poIs(child, 'w:object')) {
      // §14 legacy VML picture (and OLE-object image previews). Modern
      // <w:drawing> wins under MC resolution, so this fires only for pure-VML
      // content — common in headers and older files. Unlike a DrawingML blip,
      // a VML image is materialised only when its part actually resolves to
      // bytes: a dangling <v:imagedata r:id> (the referenced media stripped
      // from the package, as some corpus files have) carries nothing to render,
      // so we skip the phantom rather than emit an empty picture.
      const content = parseVmlPicture(child);
      if (content && content.kind === 'image') {
        const resource = ctx.resolveImage?.(content.imageId);
        if (resource) {
          inlineImage = { resource, width: content.width, height: content.height };
        }
      }
    }
  }
  return {
    run: {
      text,
      properties,
      ...(inlineImage ? { inlineImage } : {}),
      ...(pageBreak ? { pageBreak: true } : {}),
      ...(footnoteRef !== undefined ? { footnoteRef } : {}),
      ...(endnoteRef !== undefined ? { endnoteRef } : {}),
      ...(noteNumber ? { noteNumber: true } : {}),
    },
    ...(fldChar ? { fldChar } : {}),
    ...(instrText !== undefined ? { instrText } : {}),
  };
}

function elementTag(node: PoNode): string | undefined {
  for (const key of Object.keys(node)) {
    if (key !== ':@' && key !== '#text') return key;
  }
  return undefined;
}

// §17.11 — footnotes.xml / endnotes.xml. Returns content by id; the
// separator / continuationSeparator / continuationNotice stubs (negative ids
// or an explicit w:type) are skipped — the layout draws its own separator.
export function parseNotes(
  notesXml: Uint8Array,
  rootTag: 'w:footnotes' | 'w:endnotes',
  noteTag: 'w:footnote' | 'w:endnote',
  ctx: ParseContext = DEFAULT_PARSE_CONTEXT,
): Map<string, Array<BodyElement>> {
  const xml = decoder.decode(notesXml);
  const tree = parser.parse(xml) as Array<PoNode>;
  const root = poFindByPath(tree, [rootTag]);
  const out = new Map<string, Array<BodyElement>>();
  if (!root) return out;
  for (const note of poChildren(root)) {
    if (!poIs(note, noteTag)) continue;
    const type = poAttr(note, 'type');
    if (type !== undefined && type !== 'normal') continue;
    const id = poAttr(note, 'id');
    if (id === undefined) continue;
    out.set(id, parseBodyElements(poChildren(note), ctx));
  }
  return out;
}
