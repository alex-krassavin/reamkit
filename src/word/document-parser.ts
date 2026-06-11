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
  SectionProperties,
} from '@/core/document-model';

import type { ColorResolver } from '@/core/drawingml/colors';
import type { Pt, ResourceId } from '@/core/ir';
import type { PoNode } from '@/core/po-helpers';
import { twipsToPt } from '@/core/ir';
import { parseOMath } from '@/word/omml-parser';
import { defaultColorResolver } from '@/core/drawingml/colors';
import { expandMcChildren, parseDrawing } from '@/word/drawing-parser';
import { parseParagraphProperties } from '@/word/paragraph-properties';
import { poAttr, poChildren, poFindByPath, poIntAttr, poIs, poText } from '@/core/po-helpers';
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

// Document-wide resolvers every nested parser needs — one context object
// instead of threading a parameter pair through ten signatures (oop-design §8).
export interface ParseContext {
  readonly resolveColor: ColorResolver;
  readonly resolveImage?: ImageResolver;
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
    }
  }

  return {
    ...(pageSize ? { pageSize } : {}),
    ...(margins ? { margins } : {}),
    headers,
    footers,
    ...(titlePg ? { titlePg: true } : {}),
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
  for (const child of children) {
    if (poIs(child, 'w:p')) {
      const drawing = tryExtractDrawingFromParagraph(child, ctx);
      out.push(drawing ?? { kind: 'paragraph', paragraph: parseParagraph(child, ctx) });
    } else if (poIs(child, 'w:tbl')) {
      out.push({ kind: 'table', table: parseTable(child, ctx) });
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
function tryExtractDrawingFromParagraph(p: PoNode, ctx: ParseContext): BodyElement | null {
  let drawing: PoNode | undefined;
  let hasText = false;
  for (const child of poChildren(p)) {
    if (!poIs(child, 'w:r')) continue;
    for (const runChild of expandMcChildren(poChildren(child))) {
      if (poIs(runChild, 'w:drawing')) {
        if (!drawing) drawing = runChild;
      } else if (poIs(runChild, 'w:t') && poText(runChild).length > 0) {
        hasText = true;
      } else if (
        poIs(runChild, 'w:tab') ||
        poIs(runChild, 'w:br') ||
        poIs(runChild, 'w:noBreakHyphen')
      ) {
        hasText = true;
      }
    }
  }
  if (!drawing || hasText) return null;

  // Inject parseBodyElements (bound to this context) so a shape's text box is
  // parsed without a module cycle.
  const parseBody = (children: ReadonlyArray<PoNode>): Array<BodyElement> =>
    parseBodyElements(children, ctx);
  const content = parseDrawing(drawing, ctx.resolveColor, parseBody);
  if (!content) return null;

  const pPrNode = poChildren(p).find((c) => poIs(c, 'w:pPr'));
  const paragraphProperties = pPrNode ? parseParagraphProperties(poElementToFlat(pPrNode)) : {};

  if (content.kind === 'image') {
    const resource = ctx.resolveImage?.(content.imageId);
    return {
      kind: 'image',
      image: {
        ...(resource ? { resource } : {}),
        width: content.width,
        height: content.height,
        paragraphProperties,
        ...(content.altText ? { altText: content.altText } : {}),
      },
    };
  }
  if (content.kind === 'chart') {
    return {
      kind: 'chart',
      chart: {
        chartRelId: content.chartRelId,
        width: content.width,
        height: content.height,
        paragraphProperties,
        ...(content.altText ? { altText: content.altText } : {}),
      },
    };
  }
  return {
    kind: 'shape',
    shape: {
      ...content.data,
      paragraphProperties,
      ...(content.altText ? { altText: content.altText } : {}),
    },
  };
}

function parseParagraph(p: PoNode, ctx: ParseContext): Paragraph {
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
  const runs: Array<Run> = [];
  collectRuns(p, runs, ctx);
  return { properties, runs };
}

function collectRuns(container: PoNode, out: Array<Run>, ctx: ParseContext): void {
  for (const child of poChildren(container)) {
    if (poIs(child, 'w:pPr')) continue;
    if (poIs(child, 'w:r')) {
      out.push(parseRun(child, ctx));
      continue;
    }
    // OfficeMath: an inline equation (m:oMath) or a display paragraph
    // (m:oMathPara, holding one or more m:oMath) → math runs.
    if (poIs(child, 'm:oMath')) {
      out.push({ text: '', properties: {}, math: parseOMath(child) });
      continue;
    }
    if (poIs(child, 'm:oMathPara')) {
      for (const om of poChildren(child)) {
        if (poIs(om, 'm:oMath')) out.push({ text: '', properties: {}, math: parseOMath(om) });
      }
      continue;
    }
    const tag = elementTag(child);
    if (tag && RUN_CONTAINER_TAGS.has(tag)) {
      collectRuns(child, out, ctx);
    }
  }
}

function parseRun(r: PoNode, ctx: ParseContext): Run {
  const rPr = poChildren(r).find((c) => poIs(c, 'w:rPr'));
  const properties = parseRunProperties(rPr ? poElementToFlat(rPr) : undefined);
  let text = '';
  let pageBreak = false;
  let inlineImage: InlineImage | undefined;
  for (const child of expandMcChildren(poChildren(r))) {
    if (poIs(child, 'w:rPr')) continue;
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
    }
  }
  return {
    text,
    properties,
    ...(inlineImage ? { inlineImage } : {}),
    ...(pageBreak ? { pageBreak: true } : {}),
  };
}

function elementTag(node: PoNode): string | undefined {
  for (const key of Object.keys(node)) {
    if (key !== ':@' && key !== '#text') return key;
  }
  return undefined;
}
