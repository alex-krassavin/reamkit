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
// and paragraph alignment/indent level; per-shape placement, images and autoshapes
// come in later waves — recorded as a loss until then.

import type {
  Alignment,
  BodyElement,
  ParagraphProperties,
  Run,
  RunProperties,
  UnderlineStyle,
} from '@/core/document-model';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir';
import type { PptContent, PptParagraph, PptRun } from '@/pptx/ppt/ppt-text';

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

// The slide text and its run/paragraph formatting are read; the higher-level
// structure (shape placement, images, autoshapes) is not yet. Reported (degraded,
// keyed on text) so a caller's loss report is honest about the gap.
const PPT_TEXT_LOSS: Loss = {
  severity: 'degraded',
  feature: FEATURES.text,
  detail:
    "legacy .ppt: each slide's text (with run formatting — bold/italic/underline/size/colour — and paragraph alignment/indent) is read into one page per slide; per-shape placement, images and autoshapes are not yet (re-save as .pptx for full fidelity)",
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
    const lines = slide.paragraphs.filter((p) => paragraphText(p).length > 0);
    const breakBefore = slideIndex > 0;
    if (lines.length === 0) {
      body.push(anchorParagraph(breakBefore));
      return;
    }
    lines.forEach((para, lineIndex) => {
      body.push({
        kind: 'paragraph',
        paragraph: {
          properties: toParaProperties(para, breakBefore && lineIndex === 0),
          runs: para.runs.map(toRun),
        },
      });
    });
  });
  if (body.length === 0) body.push(anchorParagraph(false));
  return body;
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
