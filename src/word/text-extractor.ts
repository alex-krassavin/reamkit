// Backwards-compatible plain-text extractor used by tests/tools that still
// expect a flat string[] of paragraph text. Defers to parseDocument and
// flattens only paragraph runs (table contents are skipped).

import { parseDocument } from '@/word/document-parser';

/**
 * Extract the body paragraphs of a WordprocessingML document as a flat array of
 * plain text — one string per paragraph, runs concatenated. A backwards-compatible
 * convenience over {@link parseDocument} for tests/tools that want only text;
 * table contents and all formatting are dropped.
 *
 * @param documentXml The raw `word/document.xml` bytes.
 * @returns One joined run-text string per body paragraph, in document order.
 */
export function extractParagraphs(documentXml: Uint8Array): Array<string> {
  return parseDocument(documentXml)
    .filter((b) => b.kind === 'paragraph')
    .map((b) => b.paragraph)
    .map((p) => p.runs.map((r) => r.text).join(''));
}
