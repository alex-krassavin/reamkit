export { extractParagraphs } from '@/ooxml/wordproc/text-extractor';
export {
  parseDocument,
  parseBodyElements,
  parseSection,
  parseSections,
  parseHeaderFooter,
  EMPTY_SECTION,
} from '@/ooxml/wordproc/document-parser';
export { parseRunProperties } from '@/ooxml/wordproc/run-properties';
export { parseParagraphProperties } from '@/ooxml/wordproc/paragraph-properties';
export { parseStyles, EMPTY_STYLE_SHEET } from '@/ooxml/wordproc/styles-parser';
export { parseTable } from '@/ooxml/wordproc/table-parser';
export {
  loadEmbeddedFonts,
  deobfuscateEmbeddedFont,
  parseFontTable,
} from '@/ooxml/wordproc/font-table';
export { parseNumbering, EMPTY_NUMBERING } from '@/ooxml/wordproc/numbering-parser';
export { parseSettings, EMPTY_SETTINGS } from '@/ooxml/wordproc/settings-parser';
export type { DocumentSettings } from '@/ooxml/wordproc/settings-parser';
