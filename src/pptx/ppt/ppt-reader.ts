// Legacy `.ppt` reader (PPT-1) — the DocumentReader that sniffs a PowerPoint
// 97–2003 binary file (an OLE2/CFB container with a `PowerPoint Document` stream)
// and reads each slide's text into the same FlowDoc the OOXML pptx reader
// produces, so the whole render pipeline (projection → PDF/SVG/HTML, re-write to
// .docx) works on a legacy `.ppt` the way it already does for `.xls` and `.doc`.
// The shared CFB container reader (`src/core/ole`) is the keystone all three reuse.
//
// Like pptxReader, each slide becomes one page at the deck size: a page-breaking
// paragraph anchors the page and the slide's text follows as paragraphs. PPT-1
// reads the text only; per-shape placement, run/paragraph formatting, images and
// autoshapes come in later waves — recorded as a loss until then.

import type { BodyElement, ParagraphProperties } from '@/core/document-model';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir';
import type { PptContent } from '@/pptx/ppt/ppt-text';

import { FEATURES, ResourceStore, pt } from '@/core/ir';
import { isCfb, openCfb } from '@/core/ole/cfb';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';
import { extractPptContent } from '@/pptx/ppt/ppt-text';

const ZERO_WIDTH_SPACE = '​';

// PowerPoint's default deck: 10in × 7.5in (4:3) in points, the fallback when the
// DocumentAtom gives no usable slide size.
const DEFAULT_SLIDE_W = 720;
const DEFAULT_SLIDE_H = 540;
// A half-inch margin keeps the flat text dump off the page edge (PPT-1 has no
// per-shape geometry yet; PPT-2 positions text and drops to a margin-less canvas).
const SLIDE_MARGIN = 36;

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

// The slide text is read; the higher-level structure (shape placement, formatting,
// images, autoshapes) is not yet. Reported (degraded, keyed on text) so a caller's
// loss report is honest about the gap.
const PPT_TEXT_LOSS: Loss = {
  severity: 'degraded',
  feature: FEATURES.text,
  detail:
    "legacy .ppt: each slide's text is read into one page per slide; per-shape placement, run/paragraph formatting, images and autoshapes are not yet (re-save as .pptx for full fidelity)",
};

const PPT_ENCRYPTED_LOSS: Loss = {
  severity: 'dropped',
  feature: FEATURES.text,
  detail: 'legacy .ppt is encrypted/obfuscated — its text cannot be read',
};

export function readPpt(bytes: Uint8Array): ReadResult<FlowDoc> {
  const content = extractPptContent(bytes);
  const body = buildBody(content);
  const width = pt(content.slideWidthPt ?? DEFAULT_SLIDE_W);
  const height = pt(content.slideHeightPt ?? DEFAULT_SLIDE_H);

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
    resources: new ResourceStore(),
  };
  return { doc, losses: [content.encrypted ? PPT_ENCRYPTED_LOSS : PPT_TEXT_LOSS] };
}

// One page per slide: the first body element of each slide carries a page break
// (after the first slide) so every slide starts a fresh page, even an empty one.
function buildBody(content: PptContent): Array<BodyElement> {
  const body: Array<BodyElement> = [];
  content.slides.forEach((slide, slideIndex) => {
    const lines = slide.paragraphs.filter((p) => p.text.length > 0);
    const breakBefore = slideIndex > 0;
    if (lines.length === 0) {
      body.push(anchorParagraph(breakBefore));
      return;
    }
    lines.forEach((line, lineIndex) => {
      const properties: ParagraphProperties =
        breakBefore && lineIndex === 0 ? { pageBreakBefore: true } : {};
      body.push({
        kind: 'paragraph',
        paragraph: { properties, runs: [{ text: line.text, properties: {} }] },
      });
    });
  });
  if (body.length === 0) body.push(anchorParagraph(false));
  return body;
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
