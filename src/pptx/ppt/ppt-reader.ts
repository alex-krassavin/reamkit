// Legacy `.ppt` reader (PPT-1..2) — the DocumentReader that sniffs a PowerPoint
// 97–2003 binary file (an OLE2/CFB container with a `PowerPoint Document` stream)
// and reads each slide's text into the same FlowDoc the OOXML pptx reader
// produces, so the whole render pipeline (projection → PDF/SVG/HTML, re-write to
// .docx) works on a legacy `.ppt` the way it already does for `.xls` and `.doc`.
// The shared CFB container reader (`src/core/ole`) is the keystone all three reuse.
//
// Like pptxReader, each slide becomes one page at the deck size: a page-breaking
// paragraph anchors the page and the slide's text follows as paragraphs. PPT-2
// adds run formatting (bold/italic/underline/size/colour from the StyleTextPropAtom)
// and paragraph alignment/indent level; PPT-3 emits the slide's embedded images
// (in-flow after its text). Per-shape placement and autoshapes come in later waves
// — recorded as a loss until then.

import type {
  Alignment,
  BodyElement,
  ImageBlock,
  ParagraphProperties,
  Run,
  RunProperties,
  UnderlineStyle,
} from '@/core/document-model';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir';
import type { PptContent, PptImage, PptParagraph, PptRun } from '@/pptx/ppt/ppt-text';

import { FEATURES, ResourceStore, pt } from '@/core/ir';
import { isCfb, openCfb } from '@/core/ole/cfb';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';
import { extractPptContent, paragraphText } from '@/pptx/ppt/ppt-text';

const ZERO_WIDTH_SPACE = '​';

// PowerPoint's default deck: 10in × 7.5in (4:3) in points, the fallback when the
// DocumentAtom gives no usable slide size.
const DEFAULT_SLIDE_W = 720;
const DEFAULT_SLIDE_H = 540;
// A half-inch margin keeps the flat text dump off the page edge (PPT-1..2 have no
// per-shape geometry yet; PPT-4 positions text and drops to a margin-less canvas).
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

// The slide text, its run/paragraph formatting and embedded images are read; the
// higher-level structure (shape placement, autoshapes) is not yet. Reported
// (degraded, keyed on text) so a caller's loss report is honest about the gap.
const PPT_TEXT_LOSS: Loss = {
  severity: 'degraded',
  feature: FEATURES.text,
  detail:
    "legacy .ppt: each slide's text (with run formatting — bold/italic/underline/size/colour — and paragraph alignment/indent) and embedded images are read into one page per slide; per-shape placement and autoshapes are not yet (re-save as .pptx for full fidelity)",
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

// One page per slide: a slide's text paragraphs come first, then its images
// (in-flow). The first body element carries a page break (after the first slide)
// so every slide starts a fresh page, even an empty one.
function buildBody(content: PptContent, resources: ResourceStore): Array<BodyElement> {
  const body: Array<BodyElement> = [];
  content.slides.forEach((slide, slideIndex) => {
    const breakBefore = slideIndex > 0;
    const els: Array<BodyElement> = [];
    for (const para of slide.paragraphs) {
      if (paragraphText(para).length === 0) continue;
      els.push({
        kind: 'paragraph',
        paragraph: { properties: toParaProperties(para, false), runs: para.runs.map(toRun) },
      });
    }
    for (const image of slide.images ?? []) {
      const block = imageBlock(image, resources);
      if (block) els.push(block);
    }
    if (els.length === 0) {
      body.push(anchorParagraph(breakBefore));
      return;
    }
    // The page break must sit on a paragraph (the layout honors it only there).
    // A text-first slide breaks on its first paragraph; an image-first slide gets
    // a breaking anchor paragraph prepended.
    if (breakBefore) {
      const first = els[0]!;
      if (first.kind === 'paragraph') {
        els[0] = {
          kind: 'paragraph',
          paragraph: {
            ...first.paragraph,
            properties: { ...first.paragraph.properties, pageBreakBefore: true },
          },
        };
      } else {
        body.push(anchorParagraph(true));
      }
    }
    body.push(...els);
  });
  if (body.length === 0) body.push(anchorParagraph(false));
  return body;
}

// An embedded picture → an in-flow ImageBlock sized from its intrinsic pixels (at
// 96 dpi); the layout engine clamps the width to the page. The dimensions come
// from the PNG/JPEG header directly, which also gates out formats the renderer
// cannot embed (a metafile, say).
function imageBlock(image: PptImage, resources: ResourceStore): BodyElement | undefined {
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
