// Legacy `.doc` reader (DOC-1/DOC-2) — the DocumentReader that sniffs a Word
// 97–2003 binary file (an OLE2/CFB container with a `WordDocument` stream) and
// reads its text and run-level formatting into the same FlowDoc the OOXML docx
// reader produces, so the whole render pipeline (projection → PDF/SVG/HTML,
// re-write to .docx) works on a legacy `.doc` the way it already does for `.xls`.
// The shared CFB container reader (`src/core/ole`) is the keystone both legacy
// formats reuse.
//
// Scope: the document text, split into paragraphs at the Word paragraph mark,
// with bold / italic / underline / font-size read from the CHPX runs. Paragraph
// formatting, tables, images, headers/footers, lists and fields are not read yet
// — recorded as a loss (mirrors how `.xls` shipped values before full styling).

import type {
  BodyElement,
  Run,
  RunProperties,
  SectionProperties,
  UnderlineStyle,
} from '@/core/document-model';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir';
import type { DocCharProps, DocContent } from '@/word/doc/doc-text';

import { FEATURES, ResourceStore, pt } from '@/core/ir';
import { isCfb, openCfb } from '@/core/ole/cfb';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';
import { extractDocContent } from '@/word/doc/doc-text';

const ZERO_WIDTH_SPACE = '​';

// Paragraph / page-break / line-break / tab control characters.
const CR = 0x0d; // paragraph mark
const PAGE_BREAK = 0x0c;
const LINE_BREAK = 0x0b;
const TAB = 0x09;

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

// The document text + run formatting is read; paragraph formatting and the
// higher-level structure are not. Reported (degraded, keyed on text) so a
// caller's loss report is honest.
const DOC_TEXT_LOSS: Loss = {
  severity: 'degraded',
  feature: FEATURES.text,
  detail:
    'legacy .doc: the document text, paragraph breaks and run formatting (bold/italic/underline/size) are read; paragraph formatting, tables, images, headers/footers, lists and fields are not (re-save as .docx for full fidelity)',
};

const DOC_ENCRYPTED_LOSS: Loss = {
  severity: 'dropped',
  feature: FEATURES.text,
  detail: 'legacy .doc is encrypted/obfuscated — its text cannot be read',
};

export function readDoc(bytes: Uint8Array): ReadResult<FlowDoc> {
  const content = extractDocContent(bytes);
  const body = buildParagraphs(content);

  // Word's default page: US Letter with one-inch margins.
  const section: SectionProperties = {
    pageSize: { width: pt(612), height: pt(792) },
    margins: { top: pt(72), right: pt(72), bottom: pt(72), left: pt(72) },
    headers: [],
    footers: [],
  };

  const doc: FlowDoc = {
    kind: 'flow',
    body: resolveBodyStyles(body, EMPTY_STYLE_SHEET),
    sections: [],
    section,
    styles: EMPTY_STYLE_SHEET,
    resources: new ResourceStore(),
  };
  return { doc, losses: [content.encrypted ? DOC_ENCRYPTED_LOSS : DOC_TEXT_LOSS] };
}

// Turn the formatted run stream into paragraphs: split at the CR (and page-break)
// mark, carry each source run's formatting onto its document-model run, and clean
// the in-paragraph control characters (a manual line break / tab → space). One
// pass, so no control-character regex is needed.
function buildParagraphs(content: DocContent): Array<BodyElement> {
  const paragraphs: Array<BodyElement> = [];
  let runs: Array<Run> = [];
  let buf = '';
  let bufProps: RunProperties = {};

  const flushRun = (): void => {
    if (buf.length > 0) {
      runs.push({ text: buf, properties: bufProps });
      buf = '';
    }
  };
  const flushParagraph = (): void => {
    flushRun();
    paragraphs.push({
      kind: 'paragraph',
      paragraph: {
        properties: {},
        // An empty paragraph still needs a run so the line is emitted.
        runs: runs.length > 0 ? runs : [{ text: ZERO_WIDTH_SPACE, properties: {} }],
      },
    });
    runs = [];
  };

  for (const docRun of content.runs) {
    flushRun(); // the buffered text belongs to the previous run's properties
    bufProps = toRunProperties(docRun.props);
    for (const ch of docRun.text) {
      const c = ch.codePointAt(0)!;
      if (c === CR || c === PAGE_BREAK) {
        flushParagraph();
      } else if (c === LINE_BREAK || c === TAB) {
        buf += ' ';
      } else if (c >= 0x20) {
        buf += ch;
      }
    }
  }
  flushParagraph();
  // A document ends with a trailing paragraph mark — drop the empty tail it adds.
  while (paragraphs.length > 1 && isEmptyParagraph(paragraphs[paragraphs.length - 1]!)) {
    paragraphs.pop();
  }
  return paragraphs;
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

export const docReader: DocumentReader<FlowDoc> = {
  id: 'doc',
  produces: 'flow',
  supports: new Set([FEATURES.text]),
  sniff: looksLikeDoc,
  read: (bytes) => readDoc(bytes),
};
