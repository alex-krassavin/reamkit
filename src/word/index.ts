export type { HyperlinkResolver, ImageResolver, ParseContext } from '@/word/document-parser';
export { extractParagraphs } from '@/word/text-extractor';
export {
  parseDocument,
  parseBodyElements,
  parseSection,
  parseSections,
  parseHeaderFooter,
  EMPTY_SECTION,
} from '@/word/document-parser';
export { parseRunProperties } from '@/word/run-properties';
export { parseParagraphProperties } from '@/word/paragraph-properties';
export { parseStyles } from '@/word/styles-parser';
export { parseTable } from '@/word/table-parser';
export { loadEmbeddedFonts, deobfuscateEmbeddedFont, parseFontTable } from '@/word/font-table';
export { parseNumbering, EMPTY_NUMBERING } from '@/word/numbering-parser';
export { parseSettings, EMPTY_SETTINGS } from '@/word/settings-parser';
export type { DocumentSettings } from '@/word/settings-parser';
