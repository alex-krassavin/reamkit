// Backwards-compatible plain-text extractor used by tests/tools that still
// expect a flat string[] of paragraph text. Defers to parseDocument and
// flattens only paragraph runs (table contents are skipped).

import { parseDocument } from '@/ooxml/wordproc/document-parser';

export function extractParagraphs(documentXml: Uint8Array): Array<string> {
  return parseDocument(documentXml)
    .filter((b) => b.kind === 'paragraph')
    .map((b) => b.paragraph)
    .map((p) => p.runs.map((r) => r.text).join(''));
}
