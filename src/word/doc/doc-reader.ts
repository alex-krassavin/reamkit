// Legacy `.doc` reader (DOC-1) — the DocumentReader that sniffs a Word 97–2003
// binary file (an OLE2/CFB container with a `WordDocument` stream) and reads its
// text into the same FlowDoc the OOXML docx reader produces, so the whole render
// pipeline (projection → PDF/SVG/HTML, re-write to .docx) works on a legacy `.doc`
// the way it already does for `.xls`. The shared CFB container reader
// (`src/core/ole`) is the keystone both legacy formats reuse.
//
// Scope: the document text, split into paragraphs at the Word paragraph mark.
// Character/paragraph formatting, tables, images, headers/footers, lists and
// fields are not read yet — recorded as a loss (mirrors how `.xls` shipped cell
// values before styling).

import type { BodyElement, SectionProperties } from '@/core/document-model';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir';

import { FEATURES, ResourceStore, pt } from '@/core/ir';
import { isCfb, openCfb } from '@/core/ole/cfb';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';
import { extractDocText } from '@/word/doc/doc-text';

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

// Only the document text + paragraph structure is read; the rich formatting is
// not. Reported (degraded, keyed on text) so a caller's loss report is honest.
const DOC_TEXT_LOSS: Loss = {
  severity: 'degraded',
  feature: FEATURES.text,
  detail:
    'legacy .doc: the document text and paragraph breaks are read; character/paragraph formatting, tables, images, headers/footers, lists and fields are not (re-save as .docx for full fidelity)',
};

const DOC_ENCRYPTED_LOSS: Loss = {
  severity: 'dropped',
  feature: FEATURES.text,
  detail: 'legacy .doc is encrypted/obfuscated — its text cannot be read',
};

export function readDoc(bytes: Uint8Array): ReadResult<FlowDoc> {
  const { text, encrypted } = extractDocText(bytes);
  const paragraphs = splitParagraphs(text);

  const body: Array<BodyElement> = paragraphs.map((line) => ({
    kind: 'paragraph',
    paragraph: {
      properties: {},
      // An empty paragraph still needs a run so the line is emitted.
      runs: [{ text: line.length > 0 ? line : ZERO_WIDTH_SPACE, properties: {} }],
    },
  }));
  if (body.length === 0) {
    body.push({
      kind: 'paragraph',
      paragraph: { properties: {}, runs: [{ text: ZERO_WIDTH_SPACE, properties: {} }] },
    });
  }

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
  return { doc, losses: [encrypted ? DOC_ENCRYPTED_LOSS : DOC_TEXT_LOSS] };
}

// Split the document text into paragraphs at the CR (and page-break) mark; within
// a paragraph a manual line break / tab becomes a space and the remaining control
// characters are dropped. One pass, so no control-character regex is needed.
function splitParagraphs(text: string): Array<string> {
  const paragraphs: Array<string> = [];
  let current = '';
  for (const ch of text) {
    const c = ch.codePointAt(0)!;
    if (c === CR || c === PAGE_BREAK) {
      paragraphs.push(current);
      current = '';
    } else if (c === LINE_BREAK || c === TAB) {
      current += ' ';
    } else if (c >= 0x20) {
      current += ch;
    }
  }
  paragraphs.push(current);
  // A document ends with a trailing paragraph mark — drop the empty tail it adds.
  while (paragraphs.length > 1 && paragraphs[paragraphs.length - 1] === '') paragraphs.pop();
  return paragraphs;
}

export const docReader: DocumentReader<FlowDoc> = {
  id: 'doc',
  produces: 'flow',
  supports: new Set([FEATURES.text]),
  sniff: looksLikeDoc,
  read: (bytes) => readDoc(bytes),
};
