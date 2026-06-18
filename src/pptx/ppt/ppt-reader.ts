// Legacy `.ppt` reader (PPT-1..7) — the DocumentReader that sniffs a PowerPoint
// 97–2003 binary file (an OLE2/CFB container with a `PowerPoint Document` stream)
// and reads each slide's text into the same FlowDoc the OOXML pptx reader
// produces, so the whole render pipeline (projection → PDF/SVG/HTML, re-write to
// .docx) works on a legacy `.ppt` the way it already does for `.xls` and `.doc`.
// The shared CFB container reader (`src/core/ole`) is the keystone all three reuse.
//
// Like pptxReader, each slide becomes one page at the deck size. PPT-2 adds run
// formatting (bold/italic/underline/size/colour from the StyleTextPropAtom) and
// paragraph alignment/indent level; PPT-3 reads embedded images; PPT-4 positions a
// shape that carries a slide anchor as a floating ShapeBlock (text) / ImageBlock
// (picture), falling back to reading-order flow for un-anchored shapes; PPT-5 reads
// decorative autoshapes (preset geometry + literal fill/line); PPT-6 resolves a
// shape's scheme-relative fill/line colour through the slide's colour scheme;
// PPT-7 reads a freeform's exact custom geometry (pVertices + pSegmentInfo). What
// stays a loss: palette/system (non-scheme) colours.

import type {
  Alignment,
  BodyElement,
  CustomGeometry,
  CustomPathCmd,
  FloatAnchor,
  ImageBlock,
  ParagraphProperties,
  Run,
  RunProperties,
  ShapeBlock,
  ShapeFill,
  ShapeLine,
  UnderlineStyle,
} from '@/core/document-model';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir';
import type {
  PptAutoShape,
  PptContent,
  PptCustomGeometry,
  PptImage,
  PptParagraph,
  PptRect,
  PptRun,
} from '@/pptx/ppt/ppt-text';

import { FEATURES, ResourceStore, pt } from '@/core/ir';
import { isCfb, openCfb } from '@/core/ole/cfb';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';
import { extractPptContent, paragraphText } from '@/pptx/ppt/ppt-text';

const ZERO_WIDTH_SPACE = '​';

// PowerPoint's default deck: 10in × 7.5in (4:3) in points, the fallback when the
// DocumentAtom gives no usable slide size.
const DEFAULT_SLIDE_W = 720;
const DEFAULT_SLIDE_H = 540;
// A half-inch margin insets the reading-order fallback content off the page edge;
// positioned shapes anchor page-relative and so are unaffected by it.
const SLIDE_MARGIN = 36;
// A quarter-inch indent per bullet/outline level (mirrors the .doc list reader).
const LEVEL_INDENT_PT = 18;

// A legacy `.ppt` is an OLE2 container with a `PowerPoint Document` stream. The
// check also keeps a `.xls`/`.doc` (or an encrypted OOXML, also a CFB) from
// mis-routing here.
function looksLikePpt(bytes: Uint8Array): boolean {
  if (!isCfb(bytes)) return false;
  try {
    return openCfb(bytes).hasStream('PowerPoint Document');
  } catch {
    return false;
  }
}

// The slide text, its run/paragraph formatting, embedded images, per-shape
// placement, decorative autoshapes (preset and freeform geometry) and their
// scheme-resolved fill/line colour are read; palette/system colours are not.
// Reported (degraded, keyed on text) so a caller's loss report is honest.
const PPT_TEXT_LOSS: Loss = {
  severity: 'degraded',
  feature: FEATURES.text,
  detail:
    "legacy .ppt: each slide's text (with run formatting — bold/italic/underline/size/colour — and paragraph alignment/indent), embedded images, per-shape placement and decorative autoshapes (preset or exact freeform geometry with literal or scheme-resolved fill/line colours) are read into one page per slide; palette/system colours are not (re-save as .pptx for full fidelity)",
};

const PPT_ENCRYPTED_LOSS: Loss = {
  severity: 'dropped',
  feature: FEATURES.text,
  detail: 'legacy .ppt is encrypted/obfuscated — its text cannot be read',
};

export function readPpt(bytes: Uint8Array): ReadResult<FlowDoc> {
  const content = extractPptContent(bytes);
  const width = pt(content.slideWidthPt ?? DEFAULT_SLIDE_W);
  const height = pt(content.slideHeightPt ?? DEFAULT_SLIDE_H);
  const resources = new ResourceStore();
  const body = buildBody(content, resources);

  const doc: FlowDoc = {
    kind: 'flow',
    body: resolveBodyStyles(body, EMPTY_STYLE_SHEET),
    sections: [],
    section: {
      pageSize: { width, height },
      margins: {
        top: pt(SLIDE_MARGIN),
        right: pt(SLIDE_MARGIN),
        bottom: pt(SLIDE_MARGIN),
        left: pt(SLIDE_MARGIN),
      },
      headers: [],
      footers: [],
    },
    styles: EMPTY_STYLE_SHEET,
    resources,
  };
  return { doc, losses: [content.encrypted ? PPT_ENCRYPTED_LOSS : PPT_TEXT_LOSS] };
}

// One page per slide. A shape carrying an anchor is positioned (a floating
// ShapeBlock for text, a floating ImageBlock for a picture); an un-anchored shape
// flows in reading order. The page break sits on a paragraph (the layout honors it
// only there): an in-flow-first slide breaks on its first paragraph, a floating-
// only slide gets a breaking anchor paragraph to force its page.
function buildBody(content: PptContent, resources: ResourceStore): Array<BodyElement> {
  const body: Array<BodyElement> = [];
  content.slides.forEach((slide, slideIndex) => {
    const inFlowParas: Array<BodyElement> = [];
    const inFlowImages: Array<BodyElement> = [];
    const floats: Array<BodyElement> = [];
    for (const shape of slide.shapes) {
      if (shape.rectPt) {
        if (shape.paragraphs?.some((p) => paragraphText(p).length > 0)) {
          floats.push(positionedTextShape(shape.rectPt, shape.paragraphs));
        }
        if (shape.image) {
          const block = positionedImage(shape.rectPt, shape.image, resources);
          if (block) floats.push(block);
        }
        if (shape.autoShape) floats.push(positionedAutoShape(shape.rectPt, shape.autoShape));
      } else {
        for (const para of shape.paragraphs ?? []) {
          if (paragraphText(para).length > 0) inFlowParas.push(flowParagraph(para));
        }
        if (shape.image) {
          const block = inFlowImage(shape.image, resources);
          if (block) inFlowImages.push(block);
        }
      }
    }
    // In-flow text reads before in-flow pictures (a slide's outline then its media).
    pushSlide(body, [...inFlowParas, ...inFlowImages], floats, slideIndex > 0);
  });
  if (body.length === 0) body.push(anchorParagraph(false));
  return body;
}

// Append one slide's elements with the page break on a paragraph.
function pushSlide(
  body: Array<BodyElement>,
  inFlow: Array<BodyElement>,
  floats: Array<BodyElement>,
  breakBefore: boolean,
): void {
  if (inFlow.length === 0 && floats.length === 0) {
    body.push(anchorParagraph(breakBefore));
    return;
  }
  if (inFlow.length > 0) {
    const first = inFlow[0]!;
    if (breakBefore && first.kind === 'paragraph') {
      inFlow[0] = {
        kind: 'paragraph',
        paragraph: {
          ...first.paragraph,
          properties: { ...first.paragraph.properties, pageBreakBefore: true },
        },
      };
    } else if (breakBefore) {
      body.push(anchorParagraph(true));
    }
    body.push(...inFlow);
  } else {
    // Floating-only slide: an in-flow anchor forces the page the floats land on.
    body.push(anchorParagraph(breakBefore));
  }
  body.push(...floats);
}

function flowParagraph(para: PptParagraph): BodyElement {
  return {
    kind: 'paragraph',
    paragraph: { properties: toParaProperties(para, false), runs: para.runs.map(toRun) },
  };
}

// A page-relative float anchor at the shape's slide rectangle (slide coords = page
// coords on a margin-positioned canvas).
function floatAt(rect: PptRect): FloatAnchor {
  return {
    wrap: 'none',
    posH: { relativeFrom: 'page', offsetPt: pt(rect.x) },
    posV: { relativeFrom: 'page', offsetPt: pt(rect.y) },
  };
}

// A positioned text box: a borderless, fill-less ShapeBlock at the shape rectangle,
// its paragraphs top-anchored inside.
function positionedTextShape(rect: PptRect, paragraphs: ReadonlyArray<PptParagraph>): BodyElement {
  const shape: ShapeBlock = {
    float: floatAt(rect),
    width: pt(rect.w),
    height: pt(rect.h),
    geometry: { kind: 'preset', preset: 'rect', adjust: new Map() },
    fill: { kind: 'none' },
    text: {
      content: paragraphs.filter((p) => paragraphText(p).length > 0).map(flowParagraph),
      anchor: 't',
    },
    paragraphProperties: {},
  };
  return { kind: 'shape', shape };
}

// A positioned picture: a floating ImageBlock filling the shape rectangle.
function positionedImage(
  rect: PptRect,
  image: PptImage,
  resources: ResourceStore,
): BodyElement | undefined {
  if (!imagePixelSize(image.bytes)) return undefined; // gate: a decodable raster
  const block: ImageBlock = {
    float: floatAt(rect),
    resource: resources.put(image.bytes),
    width: pt(rect.w),
    height: pt(rect.h),
    paragraphProperties: {},
  };
  return { kind: 'image', image: block };
}

// MSOSPT (the FSP shape type) → a DrawingML preset name; unknown types fall back
// to a rectangle. Connectors (13, 20–49) are drawn as a line.
const SHAPE_PRESETS: ReadonlyMap<number, string> = new Map([
  [1, 'rect'],
  [2, 'roundRect'],
  [3, 'ellipse'],
  [4, 'diamond'],
  [5, 'triangle'],
  [6, 'rtTriangle'],
  [7, 'parallelogram'],
  [8, 'trapezoid'],
  [9, 'hexagon'],
  [10, 'octagon'],
  [13, 'line'],
  [20, 'line'],
]);
function isLineShape(shapeType: number): boolean {
  return shapeType === 13 || (shapeType >= 20 && shapeType <= 49);
}

// A PptPathCmd → the document-model CustomPathCmd (a rename of the `kind` field
// tag to `cmd`); the coordinates are already in path-bounds space.
function toCustomGeometry(g: PptCustomGeometry): CustomGeometry {
  const commands: Array<CustomPathCmd> = g.commands.map((c) =>
    c.kind === 'move'
      ? { cmd: 'move', x: c.x, y: c.y }
      : c.kind === 'line'
        ? { cmd: 'line', x: c.x, y: c.y }
        : c.kind === 'cubic'
          ? { cmd: 'cubic', x1: c.x1, y1: c.y1, x2: c.x2, y2: c.y2, x: c.x, y: c.y }
          : { cmd: 'close' },
  );
  return { pathWidth: g.pathWidth, pathHeight: g.pathHeight, commands };
}

// A decorative autoshape → a positioned vector ShapeBlock with its preset geometry
// (or its exact freeform geometry — PPT-7) and any literal fill / line colour.
function positionedAutoShape(rect: PptRect, auto: PptAutoShape): BodyElement {
  const line = isLineShape(auto.shapeType);
  const fill: ShapeFill =
    auto.fillColorHex && !line ? { kind: 'solid', colorHex: auto.fillColorHex } : { kind: 'none' };
  const stroke: ShapeLine | undefined = auto.lineColorHex
    ? { colorHex: auto.lineColorHex, fill: 'solid' }
    : line
      ? { fill: 'solid' }
      : undefined;
  const shape: ShapeBlock = {
    float: floatAt(rect),
    width: pt(rect.w),
    height: pt(rect.h),
    geometry: auto.geometry
      ? { kind: 'custom', custom: toCustomGeometry(auto.geometry) }
      : { kind: 'preset', preset: SHAPE_PRESETS.get(auto.shapeType) ?? 'rect', adjust: new Map() },
    fill,
    ...(stroke ? { line: stroke } : {}),
    paragraphProperties: {},
  };
  return { kind: 'shape', shape };
}

// An un-anchored picture → an in-flow ImageBlock sized from its intrinsic pixels
// (at 96 dpi); the layout engine clamps the width to the page. The dimensions come
// from the PNG/JPEG header directly, which also gates out formats the renderer
// cannot embed (a metafile, say).
function inFlowImage(image: PptImage, resources: ResourceStore): BodyElement | undefined {
  const size = imagePixelSize(image.bytes);
  if (!size) return undefined;
  const block: ImageBlock = {
    resource: resources.put(image.bytes),
    width: pt(size.w * 0.75),
    height: pt(size.h * 0.75),
    paragraphProperties: {},
  };
  return { kind: 'image', image: block };
}

// Intrinsic pixel dimensions from a PNG (IHDR) or JPEG (SOF) header, or undefined
// when the bytes are neither (so the picture is skipped rather than mis-sized).
function imagePixelSize(d: Uint8Array): { w: number; h: number } | undefined {
  if (d.length >= 24 && d[0] === 0x89 && d[1] === 0x50 && d[2] === 0x4e && d[3] === 0x47) {
    const w = u32be(d, 16);
    const h = u32be(d, 20);
    return w > 0 && h > 0 ? { w, h } : undefined;
  }
  if (d.length >= 4 && d[0] === 0xff && d[1] === 0xd8) {
    let off = 2;
    while (off + 9 < d.length) {
      if (d[off] !== 0xff) {
        off++;
        continue;
      }
      const marker = d[off + 1]!;
      if (
        marker >= 0xc0 &&
        marker <= 0xcf &&
        marker !== 0xc4 &&
        marker !== 0xc8 &&
        marker !== 0xcc
      ) {
        const h = (d[off + 5]! << 8) | d[off + 6]!;
        const w = (d[off + 7]! << 8) | d[off + 8]!;
        return w > 0 && h > 0 ? { w, h } : undefined;
      }
      const len = (d[off + 2]! << 8) | d[off + 3]!;
      if (len < 2) break;
      off += 2 + len;
    }
  }
  return undefined;
}

function u32be(d: Uint8Array, o: number): number {
  return ((d[o]! << 24) | (d[o + 1]! << 16) | (d[o + 2]! << 8) | d[o + 3]!) >>> 0;
}

function toRun(r: PptRun): Run {
  const underline: UnderlineStyle | undefined = r.underline ? 'single' : undefined;
  const properties: RunProperties = {
    ...(r.bold ? { bold: true } : {}),
    ...(r.italic ? { italic: true } : {}),
    ...(underline ? { underline } : {}),
    ...(r.sizePt ? { fontSizePt: pt(r.sizePt) } : {}),
    ...(r.colorHex ? { colorHex: r.colorHex } : {}),
  };
  return { text: r.text, properties };
}

function toParaProperties(p: PptParagraph, breakBefore: boolean): ParagraphProperties {
  const alignment = p.align !== undefined ? alignmentFrom(p.align) : undefined;
  const level = p.level ?? 0;
  return {
    ...(breakBefore ? { pageBreakBefore: true } : {}),
    ...(alignment ? { alignment } : {}),
    ...(level > 0 ? { indentLeft: pt(level * LEVEL_INDENT_PT) } : {}),
  };
}

// PowerPoint TextAlignmentEnum → the document-model alignment (0 = left default).
function alignmentFrom(a: number): Alignment | undefined {
  switch (a) {
    case 1:
      return 'center';
    case 2:
      return 'right';
    case 3:
    case 6: // justifyLow ≈ justify
      return 'both';
    case 4:
    case 5: // thaiDistributed ≈ distribute
      return 'distribute';
    default:
      return undefined;
  }
}

// An (otherwise empty) page anchor: a single zero-width-space run gives the page
// one line so it is actually emitted.
function anchorParagraph(breakBefore: boolean): BodyElement {
  return {
    kind: 'paragraph',
    paragraph: {
      properties: breakBefore ? { pageBreakBefore: true } : {},
      runs: [{ text: ZERO_WIDTH_SPACE, properties: {} }],
    },
  };
}

export const pptReader: DocumentReader<FlowDoc> = {
  id: 'ppt',
  produces: 'flow',
  supports: new Set([FEATURES.text]),
  sniff: looksLikePpt,
  read: (bytes) => readPpt(bytes),
};
