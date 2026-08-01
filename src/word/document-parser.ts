// ECMA-376 Part 1 §17 — WordprocessingML document parser.
//
// Produces BodyElement[] preserving the original interleaving of paragraphs
// and tables in the body. Uses fast-xml-parser in preserveOrder mode for
// traversal; for per-element property extraction we adapt PO subtrees to the
// flat shape the rPr/pPr parsers consume (po-to-flat).

import { XMLParser } from 'fast-xml-parser';

import type {
  BodyElement,
  Comment,
  HeaderFooterReference,
  HeaderFooterType,
  InlineImage,
  PageMargins,
  PageSize,
  Paragraph,
  ParagraphProperties,
  Run,
  Section,
  SectionColumns,
  SectionProperties,
  TabStop,
} from '@/core/document-model';

import type { ColorResolver } from '@/core/drawingml/colors';
import type { Loss, ResourceId } from '@/core/ir';
import type { PoNode } from '@/core/po-helpers';
import type { DrawingContent } from '@/word/drawing-parser';
import { resolveInternalEntities } from '@/core/opc/xml-entities';
import { emuToPt, pt, twipsToPt } from '@/core/ir';
import { parseOMath } from '@/word/omml-parser';
import { defaultColorResolver } from '@/core/drawingml/colors';
import { expandMcChildren, parseDrawing, parseVmlPicture } from '@/word/drawing-parser';
import { diagramTransform, noDiagramOverrideLoss, parseDiagramNodes } from '@/pptx/slide-parser';
import { parseParagraphProperties } from '@/word/paragraph-properties';
import {
  poAttr,
  poAttrLocal,
  poChildren,
  poFindByPath,
  poFindDescendant,
  poIntAttr,
  poIs,
  poIsLocal,
  poText,
} from '@/core/po-helpers';
import { poElementToFlat } from '@/word/po-to-flat';
import { parseRunProperties } from '@/word/run-properties';
import { parseBorders, parseTable } from '@/word/table-parser';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  // §4.1 of XML 1.0: a numeric character reference is not an entity — `&#10;`
  // IS a line feed and every parser must decode it. fast-xml-parser gates that
  // on `htmlEntities`, which defaults to false, so `&#10;` reached the page as
  // five literal characters (formats.xlsx writes "Hello,&#10;Calc!"). Named
  // HTML entities come along with the switch; in XML they are undefined anyway,
  // and reading `&nbsp;` as a space beats drawing it. Nested DOCTYPE entities
  // stay unexpanded either way — the parser never registers them (54764-2.xlsx).
  htmlEntities: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
  // The parser's own guard against runaway nesting defaults to 100 tags, which
  // real documents reach: deep-table-cell.docx nests tables past it and the
  // whole file went unread over a limit meant for pathological input. Deep
  // enough for any document a word processor can produce, and still bounded.
  maxNestedTags: 1000,
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

/**
 * Resolves a drawing relationship id to a content-addressed `ResourceId` —
 * supplied by the converter (which owns the OPC package and the `ResourceStore`).
 */
export type ImageResolver = (relId: string) => ResourceId | undefined;
/** Resolves a `w:hyperlink` `r:id` to its external target URL. */
export type HyperlinkResolver = (relId: string) => string | undefined;

/**
 * A SmartArt diagram's pre-rendered drawing override: its `dsp:spTree`, plus a
 * resolver for the picture fills its nodes name — those relationships belong to
 * the drawing part, not to the part that references the diagram.
 */
export interface ResolvedDiagram {
  readonly spTree: PoNode;
  readonly resolveImage?: ImageResolver;
}

/**
 * Document-wide resolvers every nested parser needs — one context object instead
 * of threading a parameter pair through ten signatures (oop-design §8).
 */
export interface ParseContext {
  /** Resolver for theme/scheme/auto colours. */
  readonly resolveColor: ColorResolver;
  /** Resolver for a drawing relationship id to a stored image. */
  readonly resolveImage?: ImageResolver;
  /**
   * §17.16.22 `w:hyperlink` `r:id` → external target URL from the owning part's
   * rels (`TargetMode="External"` only). Absent ⇒ links unwrap to plain text.
   */
  readonly resolveHyperlink?: HyperlinkResolver;
  /**
   * SmartArt: a data relationship id (`dgm:relIds` `@r:dm`) → the diagram's
   * pre-rendered drawing override, or `undefined` when the file ships none
   * (E-SMARTART SA2).
   */
  readonly resolveDiagram?: (relId: string) => ResolvedDiagram | undefined;
  /**
   * §21.2 a `c:chart` `@r:id` → the chart part's path, the key the reader files
   * parsed charts under. Relationship ids are scoped to their owning part, so a
   * footer's `rId1` and the body's `rId1` are different charts. Absent ⇒ the
   * raw rel id is kept.
   */
  readonly resolveChartPart?: (relId: string) => string | undefined;
  /**
   * §20.1.4.2.19 — the theme's `a:lnStyleLst` widths in points, which a
   * gallery-styled shape's `a:lnRef idx` indexes for its outline weight.
   */
  readonly themeLineWidths?: ReadonlyArray<number>;
  /**
   * Sink for graceful-degradation notices (E-SMARTART SA3): a SmartArt with no
   * drawing override records a dropped-feature {@link Loss} rather than vanishing.
   */
  readonly onLoss?: (loss: Loss) => void;
  /**
   * §17.13.4 comment ranges currently open as the body is read — a mutable set
   * the run collector stamps onto each run (`commentRangeRefs`). A comment range
   * can span paragraphs, so the state is document-level, not per-paragraph. The
   * reference is readonly; its contents mutate during the walk (CM2c).
   */
  readonly openCommentRanges?: Set<string>;
}

/** The default {@link ParseContext} — just the default colour resolver. */
export const DEFAULT_PARSE_CONTEXT: ParseContext = { resolveColor: defaultColorResolver };

/**
 * Parse `word/document.xml` (ECMA-376 Part 1 §17) into a flat list of
 * {@link BodyElement}, preserving the original interleaving of paragraphs and
 * tables.
 *
 * @param documentXml The raw `document.xml` bytes.
 * @param ctx         The document-wide parse context.
 * @returns The body elements; empty when the `w:body` is absent.
 */
export function parseDocument(
  documentXml: Uint8Array,
  ctx: ParseContext = DEFAULT_PARSE_CONTEXT,
  blocks?: BlockCounter,
): Array<BodyElement> {
  // A DTD is forbidden in OPC and fast-xml-parser refuses one outright when it
  // declares an EXTERNAL entity — which is the right instinct and the wrong
  // outcome: ExternalEntityInText.docx names http://poi.apache.org/ and the
  // whole document went unread over it. The subset is resolved and stripped
  // here, and a reference to something we will not fetch simply drops.
  const xml = resolveInternalEntities(decoder.decode(documentXml));
  const tree = parser.parse(xml) as Array<PoNode>;
  const body = poFindByPath(tree, ['w:document', 'w:body']);
  if (!body) return [];
  return parseBodyElements(poChildren(body), ctx, blocks);
}

/**
 * Lines a parsed body up with anything counted in SOURCE blocks. `index[i]` is
 * the ordinal of the `w:p`/`w:tbl` that produced body element `i`; `next` is the
 * running count. A paragraph carrying anchored drawings produces several
 * elements, so the two counts diverge — and {@link parseSections} counts blocks.
 */
export interface BlockCounter {
  readonly index: Array<number>;
  next: number;
}

/** A fresh {@link BlockCounter}. */
export function newBlockCounter(): BlockCounter {
  return { index: [], next: 0 };
}

/**
 * Translate a section boundary counted in source blocks into one counted in body
 * elements: the number of elements that came from a block before `blockEnd`.
 *
 * @param blocks   The counter filled while parsing the body.
 * @param blockEnd The section's `endIndex`, in source blocks.
 * @returns The matching body-element index.
 */
export function bodyIndexForBlock(blocks: BlockCounter, blockEnd: number): number {
  let n = 0;
  while (n < blocks.index.length && blocks.index[n]! < blockEnd) n++;
  return n;
}

const HF_TYPES = new Set<HeaderFooterType>(['default', 'first', 'even']);

/** The empty {@link SectionProperties} fallback (no headers/footers). */
export const EMPTY_SECTION: SectionProperties = {
  headers: [],
  footers: [],
};

/**
 * Parse the document's final section properties (ECMA-376 Part 1 §17.6.17 — the
 * body-level `sectPr` is the last child of `w:body` and describes the final
 * section). Returned as a single {@link SectionProperties} for backward
 * compatibility; multi-section documents use {@link parseSections} instead.
 *
 * @param documentXml The raw `document.xml` bytes.
 * @returns The final section's properties, or {@link EMPTY_SECTION} when there is none.
 */
export function parseSection(documentXml: Uint8Array): SectionProperties {
  const sections = parseSections(documentXml);
  if (sections.length === 0) return EMPTY_SECTION;
  // Use the final section as the document-wide fallback; it usually carries
  // pgSz/pgMar even when intermediate sections only override headers.
  return sections[sections.length - 1]!.properties;
}

/**
 * Collect every section in the document (ECMA-376 §17.6): one `sectPr` per
 * intermediate paragraph plus the body-final one. Each {@link Section} carries
 * the exclusive `endIndex` into the body-element list, so section `i` applies to
 * `body[sections[i-1].endIndex .. endIndex)`. A document with no `sectPr` at all
 * returns a single empty section spanning the whole body.
 *
 * @param documentXml The raw `document.xml` bytes.
 * @returns The sections in document order; empty when the `w:body` is absent.
 */
export function parseSections(documentXml: Uint8Array): Array<Section> {
  const xml = resolveInternalEntities(decoder.decode(documentXml));
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
    } else if (poIs(child, 'w:sdt')) {
      // A block-level content control is chrome: its children are body flow,
      // and parseBodyElements counts them, so this must too.
      const content = poChildren(child).find((c) => poIs(c, 'w:sdtContent'));
      for (const inner of content ? poChildren(content) : []) {
        if (poIs(inner, 'w:p') || poIs(inner, 'w:tbl')) bodyIdx++;
      }
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
  let pageNumberStart: number | undefined;
  let lineNumbering: SectionProperties['lineNumbering'];
  let columns: SectionColumns | undefined;
  let sectionStart: 'continuous' | 'nextPage' | undefined;
  let pageBorders: SectionProperties['pageBorders'];
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
    } else if (poIs(child, 'w:lnNumType')) {
      // §17.6.8 — line numbers in the margin. fdo66543.docx counts by three and
      // we printed none of them.
      const countBy = poIntAttr(child, 'countBy') ?? 1;
      const start = poIntAttr(child, 'start') ?? 1;
      const distance = poIntAttr(child, 'distance');
      const restartRaw = poAttr(child, 'restart');
      lineNumbering = {
        countBy: countBy > 0 ? countBy : 1,
        start,
        ...(distance !== undefined ? { distancePt: twipsToPt(distance) } : {}),
        restart:
          restartRaw === 'newSection' || restartRaw === 'continuous' ? restartRaw : 'newPage',
      };
    } else if (poIs(child, 'w:pgNumType')) {
      // §17.6.12 — the number this section's first page carries.
      // fdo44689_start_page_0.docx asks for 0 and its footer printed 1.
      const start = poIntAttr(child, 'start');
      if (start !== undefined) pageNumberStart = start;
    } else if (poIs(child, 'w:cols')) {
      columns = parseColumns(child);
    } else if (poIs(child, 'w:pgBorders')) {
      // §17.6.10 — the rules around the page, spelled exactly as a cell's are.
      // `@w:offsetFrom` defaults to the text margin (§17.18.65 ST_PageBorderOffset).
      const borders = parseBorders(child);
      if (borders) {
        pageBorders = {
          borders,
          offsetFrom: poAttr(child, 'offsetFrom') === 'page' ? 'page' : 'text',
        };
      }
    } else if (poIs(child, 'w:type')) {
      // §17.6.22 ST_SectionMark. Only `continuous` keeps the page; the
      // odd/even/column starts all begin a new one, which is what we do for
      // every section anyway.
      sectionStart = poAttr(child, 'val') === 'continuous' ? 'continuous' : 'nextPage';
    }
  }

  return {
    ...(pageSize ? { pageSize } : {}),
    ...(margins ? { margins } : {}),
    headers,
    footers,
    ...(titlePg ? { titlePg: true } : {}),
    ...(pageNumberStart !== undefined ? { pageNumberStart } : {}),
    ...(lineNumbering ? { lineNumbering } : {}),
    ...(columns ? { columns } : {}),
    ...(sectionStart ? { sectionStart } : {}),
    ...(pageBorders ? { pageBorders } : {}),
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

/**
 * Parse `word/header*.xml` or `word/footer*.xml`. The root is `w:hdr` or
 * `w:ftr`, whose children are the same body-element shape as the main document.
 *
 * @param xml The raw header/footer part bytes.
 * @param ctx The document-wide parse context.
 * @returns The body elements; empty when neither root is found.
 */
export function parseHeaderFooter(
  xml: Uint8Array,
  ctx: ParseContext = DEFAULT_PARSE_CONTEXT,
): Array<BodyElement> {
  const tree = parser.parse(resolveInternalEntities(decoder.decode(xml))) as Array<PoNode>;
  const root = tree.find((n) => poIs(n, 'w:hdr') || poIs(n, 'w:ftr'));
  if (!root) return [];
  return parseBodyElements(poChildren(root), ctx);
}

/**
 * Parse a sequence of body-level children (`w:p`, `w:tbl`, `w:sdt`,
 * `w:bookmarkStart`) into {@link BodyElement}s, preserving order. A lone-drawing
 * paragraph collapses to a standalone image/shape/chart block; a block-level SDT
 * unwraps to its content; a body-level bookmark anchors onto the next paragraph.
 *
 * @param children The body-level child nodes.
 * @param ctx      The document-wide parse context.
 * @returns The parsed body elements.
 */
export function parseBodyElements(
  children: ReadonlyArray<PoNode>,
  ctx: ParseContext = DEFAULT_PARSE_CONTEXT,
  // Out-param: for each emitted element, the ordinal of the `w:p`/`w:tbl` it
  // came from. A paragraph carrying anchored drawings emits several elements,
  // so this is what lets a caller line the body up with something counted in
  // source blocks — {@link parseSections}' endIndex, for one.
  blocks?: BlockCounter,
): Array<BodyElement> {
  const out: Array<BodyElement> = [];
  const mark = (n: number): void => {
    if (!blocks) return;
    for (let i = 0; i < n; i++) blocks.index.push(blocks.next);
  };
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
        mark(drawings.length);
      } else {
        const anchored: Array<BodyElement> = [];
        const paragraph = parseParagraph(child, ctx, pendingBookmarks, anchored);
        // The floats first: each places itself against the paragraph it hangs
        // off, and one emitted after it would hang off whatever follows.
        out.push(...anchored, { kind: 'paragraph', paragraph });
        mark(anchored.length + 1);
        pendingBookmarks = undefined;
      }
      if (blocks) blocks.next++;
    } else if (poIs(child, 'w:tbl')) {
      out.push({ kind: 'table', table: parseTable(child, ctx) });
      mark(1);
      if (blocks) blocks.next++;
    } else if (poIs(child, 'w:sdt')) {
      // §17.5.2 block-level structured document tag (content control): the
      // wrapper is chrome — its sdtContent children are ordinary body flow.
      const content = poChildren(child).find((c) => poIs(c, 'w:sdtContent'));
      if (content) out.push(...parseBodyElements(poChildren(content), ctx, blocks));
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
  /**
   * Every drawing in the paragraph, in order. A paragraph may anchor several —
   * LineStyle_DashType.docx hangs all seven of its rectangles off one — and
   * keeping only the first dropped the other six on the floor.
   */
  drawings: Array<PoNode>;
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
          acc.drawings.push(rc);
        } else if (poIs(rc, 'w:pict') || poIs(rc, 'w:object')) {
          // Only a VML node that actually bears something to draw is a
          // candidate — a picture (`v:imagedata`) or §14.1.2.22 WordArt
          // (`v:textpath`). An empty frame or a bare ActiveX/OLE control object
          // is ignored, so a paragraph that pairs a picture run with a control
          // run still collapses on the picture.
          if (
            !acc.vml &&
            (poFindDescendant(rc, 'v:imagedata') ?? poFindDescendant(rc, 'v:textpath')) !==
              undefined
          ) {
            acc.vml = rc;
          }
        } else if (poIs(rc, 'w:t') && poText(rc).length > 0) {
          acc.hasOther = true;
        } else if (
          poIs(rc, 'w:tab') ||
          poIs(rc, 'w:ptab') ||
          poIs(rc, 'w:br') ||
          poIs(rc, 'w:noBreakHyphen')
        ) {
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
  const scan: LoneDrawingScan = { hasOther: false, drawings: [] };
  scanForLoneDrawing(p, scan);
  const { drawings, vml } = scan;
  if (scan.hasOther || (drawings.length === 0 && !vml)) return null;

  // Inject parseBodyElements (bound to this context) so a shape's text box is
  // parsed without a module cycle. A modern <w:drawing> takes precedence over a
  // legacy <w:pict>/<w:object> VML picture (§14). Collapsing a lone picture to
  // a standalone block here is what keeps the round-trip symmetric: the writer
  // re-emits a block image as its own lone-drawing paragraph, which a re-read
  // collapses again — so the FIRST read must collapse too, or paragraph counts
  // drift by one on every standalone VML image.
  const parseBody = (children: ReadonlyArray<PoNode>): Array<BodyElement> =>
    parseBodyElements(children, ctx);
  const pPrNode = poChildren(p).find((c) => poIs(c, 'w:pPr'));
  const paragraphProperties = pPrNode ? parseParagraphProperties(poElementToFlat(pPrNode)) : {};

  if (drawings.length === 0) {
    const content = parseVmlPicture(vml!, parseBody);
    if (!content) return null;
    // A dangling VML <v:imagedata r:id> (referenced media absent from the
    // package) carries nothing to render; skip it so the paragraph stays empty
    // on both read passes rather than materialising an un-writable phantom.
    if (content.kind === 'image' && ctx.resolveImage?.(content.imageId) === undefined) return null;
    return blocksForDrawing(content, paragraphProperties, ctx);
  }

  // Every drawing the paragraph holds, in order — not just the first. They are
  // anchored floats, so each places itself, and a paragraph carrying one
  // behaves exactly as it did.
  const out: Array<BodyElement> = [];
  for (const d of drawings) {
    const content = parseDrawing(
      d,
      ctx.resolveColor,
      parseBody,
      ctx.resolveImage,
      ctx.resolveChartPart,
      ctx.themeLineWidths,
    );
    if (!content) continue;
    out.push(...(blocksForDrawing(content, paragraphProperties, ctx) ?? []));
  }
  return out.length > 0 ? out : null;
}

/**
 * One parsed drawing as the body elements it becomes: a picture, a chart, a
 * shape, or the several shapes a SmartArt diagram draws.
 *
 * @param content              The parsed drawing.
 * @param paragraphProperties  The owning paragraph's properties, which a block
 *                             carries for its spacing and alignment.
 * @param ctx                  The document-wide parse context.
 * @returns The blocks, or `null` when the drawing renders nothing.
 */
function blocksForDrawing(
  content: DrawingContent,
  paragraphProperties: ParagraphProperties,
  ctx: ParseContext,
): Array<BodyElement> | null {
  if (content.kind === 'image') {
    const resource = ctx.resolveImage?.(content.imageId);
    return [
      {
        kind: 'image',
        image: {
          ...(resource ? { resource } : {}),
          width: content.width,
          height: content.height,
          ...(content.crop ? { crop: content.crop } : {}),
          ...(content.rotation60k ? { rotation60k: content.rotation60k } : {}),
          ...(content.relativeSize ? { relativeSize: content.relativeSize } : {}),
          ...(content.effectExtent ? { effectExtent: content.effectExtent } : {}),
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
    // SmartArt: resolve the drawing override and render its nodes (E-SMARTART
    // SA2). No override ⇒ keep the (empty) paragraph, byte-stable.
    const diagram = ctx.resolveDiagram?.(content.dmRelId);
    if (!diagram) {
      // No drawing override shipped — record a graceful loss and keep the
      // (empty) paragraph, byte-stable (SA3).
      ctx.onLoss?.(noDiagramOverrideLoss());
      return null;
    }
    const frame = { x: 0, y: 0, cx: content.widthEmu, cy: content.heightEmu };
    const nodes = parseDiagramNodes(
      diagram.spTree,
      diagramTransform(diagram.spTree, frame),
      ctx.resolveColor,
      undefined,
      diagram.resolveImage,
    );
    // One box holding the nodes at their offsets inside it, rather than a
    // scatter of floats each anchored to the paragraph: a diagram is a drawing
    // like any other, and an INLINE one has to reserve its height the way the
    // reference does. fdo87488 stacks two full-page diagrams — the second of
    // them empty — and we drew one page where the reference draws two.
    return [
      {
        kind: 'shape',
        shape: {
          width: emuToPt(content.widthEmu),
          height: emuToPt(content.heightEmu),
          children: nodes.map(({ box, shape }) => ({
            shape,
            xPt: emuToPt(box.x),
            yPt: emuToPt(box.y),
          })),
          geometry: { kind: 'preset', preset: 'rect', adjust: new Map() },
          fill: { kind: 'none' },
          paragraphProperties,
          ...(content.altText ? { altText: content.altText } : {}),
          ...(content.float ? { float: content.float } : {}),
        },
      },
    ];
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

function parseParagraph(
  p: PoNode,
  ctx: ParseContext,
  extraBookmarks?: Array<string>,
  // §20.4.2.3 — out-param: the anchored drawings the paragraph's runs carry.
  // They hang off the paragraph rather than sitting in its lines, so the caller
  // emits them as blocks beside it.
  anchored?: Array<BodyElement>,
): Paragraph {
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
  // §17.6.17 — a `w:sectPr` in the paragraph mark makes this paragraph the last
  // of its section, and the mark itself the break. An otherwise empty one is
  // therefore not a blank line: fdo73596_RunInStyle brackets its index with two
  // of them and we opened a 20pt hole on either side.
  if (pPr && poChildren(pPr).some((c) => poIs(c, 'w:sectPr'))) {
    properties = { ...properties, sectionBreak: true };
  }
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
  // §17.3.3.15 — every `w:ptab` in the paragraph becomes a stop of its own, in
  // the order they appear: the tab that reaches it is an ordinary one, and the
  // resolver takes the next stop past where the line has run. Read nowhere,
  // SimpleHeadThreeColFoot.docx printed "Footer LeftFooter MiddleFooter Right".
  const pTabs = collectPositionTabs(p);
  if (pTabs.length > 0) {
    properties = { ...properties, tabs: [...(properties.tabs ?? []), ...pTabs] };
  }
  const collected: Array<CollectedRun> = [];
  collectRuns(p, collected, ctx, undefined, undefined, anchored);
  return {
    properties,
    runs: applyFieldFsm(collected),
    ...(bookmarks.length > 0 ? { bookmarks } : {}),
  };
}

/**
 * §17.3.3.15 `w:ptab` — the absolute-position tabs a paragraph holds, as stops.
 *
 * `@w:alignment` says which edge of the text column the tab reaches for; the
 * `left` alignment reaches for no edge at all (it goes to where the text
 * already is) and states no stop. `@w:relativeTo` distinguishes the margin from
 * the paragraph's indent, which for a stop measured from the text margin is the
 * same place.
 *
 * @param p The `w:p` element.
 * @returns One stop per position tab, in document order.
 */
function collectPositionTabs(p: PoNode): Array<TabStop> {
  const out: Array<TabStop> = [];
  const visit = (node: PoNode): void => {
    for (const child of poChildren(node)) {
      if (poIs(child, 'w:ptab')) {
        const alignment = poAttr(child, 'alignment');
        if (alignment === 'center' || alignment === 'right') {
          out.push({ positionPt: pt(0), alignment, relativeTo: alignment });
        }
        continue;
      }
      if (poIs(child, 'w:r') || poIs(child, 'w:hyperlink') || poIs(child, 'w:ins')) visit(child);
    }
  };
  visit(p);
  return out;
}

// A parsed run plus the complex-field markers the FSM consumes (§17.16.18
// w:fldChar / w:instrText). Internal to run collection.
interface CollectedRun {
  readonly run: Run;
  readonly fldChar?: 'begin' | 'separate' | 'end';
  readonly instrText?: string;
}

/**
 * §17.16.5.44 `MACROBUTTON macroName displayText` — the display text, which is
 * everything after the macro's name and is all a reader ever sees of the field.
 *
 * @param instr The field instruction.
 * @returns The display text, or `undefined` when this is not a MACROBUTTON.
 */
function macroButtonText(instr: string): string | undefined {
  const m = /^\s*MACROBUTTON\s+\S+\s(.*)$/su.exec(instr);
  const shown = m?.[1]?.replace(/\s+$/u, '');
  return shown !== undefined && shown !== '' ? shown : undefined;
}

// §17.16.5.35 PAGE / §17.16.5.33 NUMPAGES: the instruction's first keyword;
// switches (\* MERGEFORMAT …) are ignored. Anything else stays a cached
// result (REF, TOC, DATE, … render their stored text exactly as before).
function parseFieldInstr(instr: string | undefined): 'PAGE' | 'NUMPAGES' | undefined {
  const kw = fieldKeyword(instr);
  return kw === 'PAGE' ? 'PAGE' : kw === 'NUMPAGES' ? 'NUMPAGES' : undefined;
}

function fieldKeyword(instr: string | undefined): string | undefined {
  if (instr === undefined) return undefined;
  return /^\s*([A-Za-z]+)/.exec(instr)?.[1]?.toUpperCase();
}

// The fields §17.16.5 defines as having NO result: they do their work — bind a
// bookmark, prompt, record an index or contents entry — and display nothing.
// Word caches a result for them all the same, and printing it put "Praun et
// al. 20012001" in the middle of fdo76163's bibliography, where both
// references print nothing.
const SILENT_FIELDS = new Set([
  'SET', // §17.16.5.53 — binds a bookmark to a value
  'ASK', // §17.16.5.2 — prompts, then binds
  'XE', // §17.16.5.71 — index entry
  'TC', // §17.16.5.63 — table-of-contents entry
  'RD', // §17.16.5.50 — referenced document for an index/TOC
  'PRIVATE', // §17.16.5.48 — data another format left behind
]);

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
      const silent = SILENT_FIELDS.has(fieldKeyword(st.instr) ?? '');
      if (silent) {
        // The field shows nothing; its cached result is bookkeeping.
      } else if (field) out.push(synthesizeFieldRun(st.result, field));
      else if (st.result.length > 0) out.push(...st.result);
      else {
        // §17.16.5.44 MACROBUTTON — the words a reader SEES are in the
        // instruction, after the macro's name, and the field caches no result
        // of its own. Unsupportedtextfields.docx asks for "contacts  ssss" and
        // we printed a blank line where LibreOffice prints them.
        const shown = macroButtonText(st.instr);
        if (shown !== undefined) out.push({ text: shown, properties: {} });
      }
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
  // §20.4.2.3 — out-param threaded to parseRun: the anchored drawings the
  // paragraph's runs carry, which are blocks of their own rather than glyphs.
  anchored?: Array<BodyElement>,
): void {
  for (const child of poChildren(container)) {
    if (poIs(child, 'w:pPr')) continue;
    // §17.13.4.3/4 — comment range bounds (siblings of w:r). Track the open set
    // so runs in between carry commentRangeRefs (the highlighted span).
    if (poIs(child, 'w:commentRangeStart')) {
      const id = poAttr(child, 'id');
      if (id !== undefined) ctx.openCommentRanges?.add(id);
      continue;
    }
    if (poIs(child, 'w:commentRangeEnd')) {
      const id = poAttr(child, 'id');
      if (id !== undefined) ctx.openCommentRanges?.delete(id);
      continue;
    }
    if (poIs(child, 'w:r')) {
      const parsed = parseRun(child, ctx, anchored);
      const ranges =
        ctx.openCommentRanges && ctx.openCommentRanges.size > 0
          ? [...ctx.openCommentRanges]
          : undefined;
      const run =
        href !== undefined || anchor !== undefined || ranges !== undefined
          ? {
              ...parsed.run,
              ...(href !== undefined ? { href } : {}),
              ...(anchor !== undefined ? { anchor } : {}),
              ...(ranges !== undefined ? { commentRangeRefs: ranges } : {}),
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
        collectRuns(child, inner, ctx, href, anchor, anchored);
        out.push({ run: synthesizeFieldRun(applyFieldFsm(inner), field, href) });
        continue;
      }
      collectRuns(child, out, ctx, href, anchor, anchored);
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
      collectRuns(child, out, ctx, childHref, childAnchor, anchored);
    }
  }
}

function parseRun(
  r: PoNode,
  ctx: ParseContext,
  // §20.4.2.3 — out-param: the ANCHORED drawings this run carries. They are not
  // part of the line at all; the caller emits them as blocks of their own.
  anchored?: Array<BodyElement>,
): { run: Run; fldChar?: 'begin' | 'separate' | 'end'; instrText?: string } {
  const rPr = poChildren(r).find((c) => poIs(c, 'w:rPr'));
  const properties = parseRunProperties(rPr ? poElementToFlat(rPr) : undefined);
  let text = '';
  let pageBreak = false;
  let columnBreak = false;
  let inlineImage: InlineImage | undefined;
  let fldChar: 'begin' | 'separate' | 'end' | undefined;
  let instrText: string | undefined;
  let footnoteRef: string | undefined;
  let endnoteRef: string | undefined;
  let commentRef: string | undefined;
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
    if (poIs(child, 'w:commentReference')) {
      const id = poAttr(child, 'id');
      if (id !== undefined) commentRef = id;
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
    } else if (poIs(child, 'w:ptab')) {
      // §17.3.3.15 — an absolute position tab is a tab: what makes it absolute
      // is where it goes, and parseParagraph collects that from the same run.
      text += '\t';
    } else if (poIs(child, 'w:br')) {
      // §17.3.3.1 — `page` forces a page break, `column` sends what follows to
      // the next column; either way the line ends here, and any other break
      // type (textWrapping/none) is a soft line break within the text flow.
      const breakType = poAttr(child, 'type');
      if (breakType === 'page') pageBreak = true;
      else {
        if (breakType === 'column') columnBreak = true;
        text += '\n';
      }
    } else if (poIs(child, 'w:noBreakHyphen')) {
      text += '‑';
    } else if (poIs(child, 'w:softHyphen')) {
      text += '­';
    } else if (poIs(child, 'w:drawing')) {
      // Only pictures render inline in M5; a wps shape in a mixed run is
      // dropped (its text is preserved). Colour is irrelevant for pictures,
      // so this deliberately does NOT take ctx.resolveColor (byte-parity with
      // the pre-context code; revisit if inline shapes ever render).
      // A shape here is drawn, not skipped, so it gets the same parsers a lone
      // drawing does: its own theme colours, its text box, its pictures.
      // chart-size.docx anchors a text box beside a run of text and we drew the
      // empty frame — its "Before.", its chart and its "After." all gone.
      const content = parseDrawing(
        child,
        ctx.resolveColor,
        (children) => parseBodyElements(children, ctx),
        ctx.resolveImage,
        ctx.resolveChartPart,
        ctx.themeLineWidths,
      );
      // §20.4.2.3 — an ANCHORED drawing is not in the line: it hangs off the
      // paragraph at a position of its own, and the text flows past it. Read as
      // an inline picture it split the line it sat in — anchor-position.docx
      // put its picture between the "A" and the "B" where every other reader
      // sets "AB" beside it.
      if (content?.float && anchored) {
        anchored.push(...(blocksForDrawing(content, {}, ctx) ?? []));
      } else if (content?.kind === 'chart' && anchored) {
        // A chart is block-sized: it takes a line of its own, so it leaves the
        // run and becomes a block ahead of the paragraph. chart-dupe.docx sets
        // one beside a trailing space, and keeping it in the run dropped it.
        anchored.push(...(blocksForDrawing(content, {}, ctx) ?? []));
      } else if (content && content.kind === 'image') {
        const resource = ctx.resolveImage?.(content.imageId);
        inlineImage = {
          ...(resource ? { resource } : {}),
          width: content.width,
          height: content.height,
          ...(content.crop ? { crop: content.crop } : {}),
          ...(content.rotation60k ? { rotation60k: content.rotation60k } : {}),
          ...(content.effectExtent ? { effectExtent: content.effectExtent } : {}),
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
      const content = parseVmlPicture(child, (children) => parseBodyElements(children, ctx));
      if (content?.float && anchored) {
        // §14.1.2 — a POSITIONED VML drawing hangs off the paragraph, picture
        // or shape alike: drawinglayer-pic-pos.docx sets its photo two inches
        // down the page and we drew it in the line.
        anchored.push(...(blocksForDrawing(content, {}, ctx) ?? []));
      } else if (content && content.kind === 'image') {
        const resource = ctx.resolveImage?.(content.imageId);
        if (resource) {
          inlineImage = { resource, width: content.width, height: content.height };
        }
      } else if (content && anchored) {
        // §14.1.2 — a drawn VML shape in a run of text: a block of its own,
        // the way an anchored drawing is. drawinglayer-pic-pos.docx frames its
        // title in one beside the paragraph's text.
        anchored.push(...(blocksForDrawing(content, {}, ctx) ?? []));
      }
    }
  }
  return {
    run: {
      text,
      properties,
      ...(inlineImage ? { inlineImage } : {}),
      ...(pageBreak ? { pageBreak: true } : {}),
      ...(columnBreak ? { columnBreak: true } : {}),
      ...(footnoteRef !== undefined ? { footnoteRef } : {}),
      ...(endnoteRef !== undefined ? { endnoteRef } : {}),
      ...(commentRef !== undefined ? { commentRef } : {}),
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

/**
 * Parse `footnotes.xml` / `endnotes.xml` (§17.11) into note content by id. The
 * separator / continuationSeparator / continuationNotice stubs (negative ids or
 * an explicit `w:type`) are skipped — the layout draws its own separator.
 *
 * @param notesXml The raw notes-part bytes.
 * @param rootTag  The part's root element (`w:footnotes` or `w:endnotes`).
 * @param noteTag  The per-note element (`w:footnote` or `w:endnote`).
 * @param ctx      The document-wide parse context.
 * @returns A map from note id to its body content.
 */
export function parseNotes(
  notesXml: Uint8Array,
  rootTag: 'w:footnotes' | 'w:endnotes',
  noteTag: 'w:footnote' | 'w:endnote',
  ctx: ParseContext = DEFAULT_PARSE_CONTEXT,
): Map<string, Array<BodyElement>> {
  const xml = resolveInternalEntities(decoder.decode(notesXml));
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

// §17.13.4 — comments.xml. Like parseNotes, but a comment carries author/date
// attribution alongside its block content, so it returns a richer Comment
// (parseNotes keeps content only). Comments have no separator/stub convention.
// Alongside the comments, capture each comment's last-paragraph w14:paraId —
// the key Microsoft's commentsExtended (w15) threads link on (CM4).
function parseCommentsRaw(
  commentsXml: Uint8Array,
  ctx: ParseContext,
): { comments: Map<string, Comment>; paraIds: Map<string, string> } {
  const xml = resolveInternalEntities(decoder.decode(commentsXml));
  const tree = parser.parse(xml) as Array<PoNode>;
  const root = poFindByPath(tree, ['w:comments']);
  const comments = new Map<string, Comment>();
  const paraIds = new Map<string, string>();
  if (!root) return { comments, paraIds };
  for (const c of poChildren(root)) {
    if (!poIs(c, 'w:comment')) continue;
    const id = poAttr(c, 'id');
    if (id === undefined) continue;
    const author = poAttr(c, 'author');
    const initials = poAttr(c, 'initials');
    const date = poAttr(c, 'date');
    comments.set(id, {
      content: parseBodyElements(poChildren(c), ctx),
      ...(author !== undefined ? { author } : {}),
      ...(initials !== undefined ? { initials } : {}),
      ...(date !== undefined ? { date } : {}),
    });
    // The thread key is the last paragraph's paraId (Word writes it on every
    // comment paragraph; commentsExtended references the final one).
    let paraId: string | undefined;
    for (const child of poChildren(c)) {
      if (!poIs(child, 'w:p')) continue;
      const pid = poAttrLocal(child, 'paraId');
      if (pid !== undefined) paraId = pid;
    }
    if (paraId !== undefined) paraIds.set(id, paraId);
  }
  return { comments, paraIds };
}

/**
 * Parse `word/comments.xml` (§17.13.4) into {@link Comment}s by id, each with its
 * block content and author/initials/date attribution. Thread metadata is added
 * separately by {@link parseCommentThreads}.
 *
 * @param commentsXml The raw `comments.xml` bytes.
 * @param ctx         The document-wide parse context.
 * @returns A map from comment id to the parsed comment.
 */
export function parseComments(
  commentsXml: Uint8Array,
  ctx: ParseContext = DEFAULT_PARSE_CONTEXT,
): Map<string, Comment> {
  return parseCommentsRaw(commentsXml, ctx).comments;
}

/** A commentsExtended entry, keyed by the comment's last-paragraph w14:paraId. */
export interface CommentExtension {
  /** The parent comment's paraId — present when this comment is a reply. */
  readonly paraIdParent?: string;
  /** w15:done — the thread is resolved. */
  readonly done: boolean;
}

/**
 * Parse `word/commentsExtended.xml` (Microsoft w15 `commentsEx`) — a flat list
 * of `commentEx`, each keyed by a comment's `paraId`: `paraIdParent` links a
 * reply to its parent, `done` flags a resolved thread. Prefix-agnostic (w15 by
 * convention).
 *
 * @param xml The raw `commentsExtended.xml` bytes.
 * @returns A map from `paraId` to its {@link CommentExtension}.
 */
export function parseCommentsExtended(xml: Uint8Array): Map<string, CommentExtension> {
  const tree = parser.parse(resolveInternalEntities(decoder.decode(xml))) as Array<PoNode>;
  const out = new Map<string, CommentExtension>();
  const root = tree.find((n) => poIsLocal(n, 'commentsEx'));
  if (!root) return out;
  for (const ex of poChildren(root)) {
    if (!poIsLocal(ex, 'commentEx')) continue;
    const paraId = poAttrLocal(ex, 'paraId');
    if (paraId === undefined) continue;
    const paraIdParent = poAttrLocal(ex, 'paraIdParent');
    const done = poAttrLocal(ex, 'done');
    out.set(paraId, {
      ...(paraIdParent !== undefined ? { paraIdParent } : {}),
      done: done === '1' || done === 'true',
    });
  }
  return out;
}

// Fold commentsExtended thread links onto the comments: a reply gains parentId
// (the comment owning its paraIdParent), a resolved thread gains done.
function linkCommentThreads(
  comments: Map<string, Comment>,
  paraIdByComment: Map<string, string>,
  ext: Map<string, CommentExtension>,
): Map<string, Comment> {
  if (ext.size === 0) return comments;
  const commentByParaId = new Map<string, string>();
  for (const [id, pid] of paraIdByComment) commentByParaId.set(pid, id);
  const out = new Map<string, Comment>();
  for (const [id, c] of comments) {
    const pid = paraIdByComment.get(id);
    const e = pid !== undefined ? ext.get(pid) : undefined;
    if (!e) {
      out.set(id, c);
      continue;
    }
    const parentId = e.paraIdParent !== undefined ? commentByParaId.get(e.paraIdParent) : undefined;
    out.set(id, {
      ...c,
      ...(parentId !== undefined ? { parentId } : {}),
      ...(e.done ? { done: true } : {}),
    });
  }
  return out;
}

/**
 * Read comments with their thread metadata: `comments.xml` for content /
 * attribution plus the optional `commentsExtended.xml` for reply links and
 * resolved flags (CM4). A reply gains `parentId`, a resolved thread gains `done`.
 *
 * @param commentsXml         The raw `comments.xml` bytes.
 * @param commentsExtendedXml The raw `commentsExtended.xml` bytes, or `undefined`.
 * @param ctx                 The document-wide parse context.
 * @returns A map from comment id to the parsed, thread-linked comment.
 */
export function parseCommentThreads(
  commentsXml: Uint8Array,
  commentsExtendedXml: Uint8Array | undefined,
  ctx: ParseContext = DEFAULT_PARSE_CONTEXT,
): Map<string, Comment> {
  const { comments, paraIds } = parseCommentsRaw(commentsXml, ctx);
  if (!commentsExtendedXml) return comments;
  return linkCommentThreads(comments, paraIds, parseCommentsExtended(commentsExtendedXml));
}

/**
 * Parse `word/people.xml` (Microsoft w15) — maps an author display name to a
 * presence identity (`w15:presenceInfo/@w15:userId`, usually an email), used to
 * enrich a comment's `authorId`. Prefix-agnostic.
 *
 * @param xml The raw `people.xml` bytes.
 * @returns A map from author name to userId.
 */
export function parsePeople(xml: Uint8Array): Map<string, string> {
  const tree = parser.parse(resolveInternalEntities(decoder.decode(xml))) as Array<PoNode>;
  const out = new Map<string, string>();
  const root = tree.find((n) => poIsLocal(n, 'people'));
  if (!root) return out;
  for (const person of poChildren(root)) {
    if (!poIsLocal(person, 'person')) continue;
    const author = poAttrLocal(person, 'author');
    if (author === undefined) continue;
    let userId: string | undefined;
    for (const child of poChildren(person)) {
      if (!poIsLocal(child, 'presenceInfo')) continue;
      const uid = poAttrLocal(child, 'userId');
      if (uid !== undefined) userId = uid;
    }
    if (userId !== undefined) out.set(author, userId);
  }
  return out;
}

/**
 * Attach each comment's `authorId` by matching its author name against the
 * {@link parsePeople} map. Comments without a matching author pass through
 * unchanged.
 *
 * @param comments The comments by id.
 * @param people   The author → userId map from `people.xml`.
 * @returns The comments with `authorId` filled in where resolvable.
 */
export function applyAuthorIds(
  comments: Map<string, Comment>,
  people: Map<string, string>,
): Map<string, Comment> {
  if (people.size === 0) return comments;
  const out = new Map<string, Comment>();
  for (const [id, c] of comments) {
    const userId = c.author !== undefined ? people.get(c.author) : undefined;
    out.set(id, userId !== undefined ? { ...c, authorId: userId } : c);
  }
  return out;
}
