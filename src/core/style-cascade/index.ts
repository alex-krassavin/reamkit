export {
  EMPTY_STYLE_SHEET,
  resolveBodyStyles,
  resolveHeadersFootersStyles,
  resolveParagraphProperties,
  resolveRunProperties,
} from '@/core/style-cascade/resolver';
export type {
  ResolvedRunProperties,
  ResolvedParagraphProperties,
} from '@/core/style-cascade/types';
export { DEFAULT_RESOLVED_RUN, DEFAULT_RESOLVED_PARAGRAPH } from '@/core/style-cascade/types';
export { resolveTableStyles } from '@/core/style-cascade/table';
