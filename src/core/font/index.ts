export { parseTtf } from '@/core/font/ttf-parser';
export type { ParsedTtf, TtfTableInfo } from '@/core/font/ttf-parser';
export { subsetTtf, glyphClosure } from '@/core/font/ttf-subset';
export { FontRegistry } from '@/core/font/font-registry';
export type { FontVariant, FontBytesByVariant } from '@/core/font/font-registry';
export { pickVariant } from '@/core/font/font-registry';
export { parseGposKerning, parseGsubLigatures, shapeText } from '@/core/font/opentype-layout';
export type { KerningMap, LigatureMap, ShapedRun } from '@/core/font/opentype-layout';
