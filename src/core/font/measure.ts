// Font measurement/encoding derived purely from the parsed font — no PDF
// involved (ir-design stage 3b / 6.4). The layout phase measures text with
// these; the PDF emitter reuses the same functions (via embedTtfFont) so emit
// encodes exactly what layout measured. CIDs are GIDs (Identity ordering), so
// the hex encoding is writer-agnostic glyph addressing.

import type { ParsedTtf } from '@/core/font/ttf-parser';
import { shapeText } from '@/core/font/opentype-layout';

/**
 * Text measurement and glyph encoding derived purely from a parsed font (no PDF
 * involved). Layout measures with these, and the PDF emitter reuses the same
 * functions, so emit encodes exactly what layout measured. CIDs are GIDs
 * (Identity ordering), so the hex encoding is writer-agnostic glyph addressing.
 */
export interface FontMeasure {
  /** Glyph advance width in 1000-unit PDF text space, for a glyph id. */
  readonly pdfWidthForGid: (gid: number) => number;
  /** Shaped width of `text` at `fontSize`, in points. */
  readonly textWidthPt: (text: string, fontSize: number) => number;
  /** Encode `text` as a hex string of 4-digit glyph ids (Identity-H addressing). */
  readonly encodeTextAsCidHex: (text: string) => string;
}

const EMPTY_KERNING = new Map<string, number>();

/**
 * Build a {@link FontMeasure} over a parsed font.
 *
 * @param parsed The parsed TTF/OTF.
 * @param kern   Whether to apply pair kerning to measured advances (E-PARITY
 *               FP4). The default keeps it; the `'word'` layout profile turns it
 *               off (Word leaves font kerning off by default, and Ream's `Tj`
 *               output is un-kerned anyway). Glyph identity is kern-independent,
 *               so only the widths change.
 * @returns The measurement / encoding closures.
 */
export function createFontMeasure(parsed: ParsedTtf, kern = true): FontMeasure {
  const scale = 1000 / parsed.unitsPerEm;
  const kerning = kern ? parsed.kerning : EMPTY_KERNING;
  const widths: Array<number> = new Array(parsed.numGlyphs);
  for (let i = 0; i < parsed.numGlyphs; i++) {
    widths[i] = Math.round((parsed.advanceWidths[i] ?? 0) * scale);
  }
  const pdfWidthForGid = (gid: number): number => {
    if (gid < 0 || gid >= parsed.numGlyphs) return 1000;
    return widths[gid]!;
  };
  const textWidthPt = (text: string, fontSize: number): number => {
    const shaped = shapeText(
      text,
      parsed.glyphForCodepoint,
      parsed.advanceWidths,
      parsed.ligatures,
      kerning,
      parsed.joiningForms,
    );
    let totalEm = 0;
    for (const a of shaped.advances) totalEm += a;
    return (totalEm * fontSize) / parsed.unitsPerEm;
  };
  const encodeTextAsCidHex = (text: string): string => {
    const shaped = shapeText(
      text,
      parsed.glyphForCodepoint,
      parsed.advanceWidths,
      parsed.ligatures,
      kerning,
      parsed.joiningForms,
    );
    let out = '';
    for (const gid of shaped.gids) {
      out += gid.toString(16).padStart(4, '0').toUpperCase();
    }
    return out;
  };
  return { pdfWidthForGid, textWidthPt, encodeTextAsCidHex };
}
