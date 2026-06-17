// Legacy `.doc` reader (DOC-1/2/3) — the DocumentReader that sniffs a Word
// 97–2003 binary file (an OLE2/CFB container with a `WordDocument` stream) and
// reads its text, run formatting and paragraph formatting into the same FlowDoc
// the OOXML docx reader produces, so the whole render pipeline (projection →
// PDF/SVG/HTML, re-write to .docx) works on a legacy `.doc` the way it already
// does for `.xls`. The shared CFB container reader (`src/core/ole`) is the
// keystone both legacy formats reuse.
//
// Scope: paragraphs of formatted runs — bold / italic / underline / font size
// (CHPX) on each run, and alignment / indentation / spacing (PAPX) on each
// paragraph. Tables, images, headers/footers, lists and fields are not read yet
// — recorded as a loss (mirrors how `.xls` shipped values before full styling).

import type {
  Alignment,
  BodyElement,
  ParagraphProperties,
  RunProperties,
  UnderlineStyle,
} from '@/core/document-model';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir';
import type { DocCharProps, DocContent, DocParaProps } from '@/word/doc/doc-text';

import { FEATURES, ResourceStore, pt } from '@/core/ir';
import { isCfb, openCfb } from '@/core/ole/cfb';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';
import { extractDocContent } from '@/word/doc/doc-text';

const ZERO_WIDTH_SPACE = '​';

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
    'legacy .doc: the document text, run formatting (bold/italic/underline/size) and paragraph formatting (alignment/indent/spacing) are read; tables, images, headers/footers, lists and fields are not (re-save as .docx for full fidelity)',
};

const DOC_ENCRYPTED_LOSS: Loss = {
  severity: 'dropped',
  feature: FEATURES.text,
  detail: 'legacy .doc is encrypted/obfuscated — its text cannot be read',
};

export function readDoc(bytes: Uint8Array): ReadResult<FlowDoc> {
  const content = extractDocContent(bytes);
  const body = buildBody(content);

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
    resources: new ResourceStore(),
  };
  return { doc, losses: [content.encrypted ? DOC_ENCRYPTED_LOSS : DOC_TEXT_LOSS] };
}

// Map each extracted paragraph onto a document-model paragraph, carrying the run
// and paragraph formatting. doc-text has already split paragraphs at the CR mark
// and cleaned the in-paragraph control characters.
function buildBody(content: DocContent): Array<BodyElement> {
  const body: Array<BodyElement> = content.paragraphs.map((p) => ({
    kind: 'paragraph',
    paragraph: {
      properties: toParaProperties(p.props),
      // An empty paragraph still needs a run so the line is emitted.
      runs:
        p.runs.length > 0
          ? p.runs.map((r) => ({ text: r.text, properties: toRunProperties(r.props) }))
          : [{ text: ZERO_WIDTH_SPACE, properties: {} }],
    },
  }));
  // A document ends with a trailing paragraph mark — drop the empty tail it adds.
  while (body.length > 1 && isEmptyParagraph(body[body.length - 1]!)) body.pop();
  if (body.length === 0) {
    body.push({
      kind: 'paragraph',
      paragraph: { properties: {}, runs: [{ text: ZERO_WIDTH_SPACE, properties: {} }] },
    });
  }
  return body;
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
