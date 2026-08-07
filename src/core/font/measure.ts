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
  /**
   * How far this text's INK actually reaches from its baseline, in points:
   * `above` up to its highest glyph edge, `below` down to its lowest. Not the
   * font's line box — "HHH" reaches a cap height and no further, and only a
   * string carrying a descender reaches under the baseline at all. WordArt
   * stretches the ink it draws to fill its shape (§20.1.9.10), so it is the
   * ink and not the line box that has to be measured.
   */
  readonly textInkPt: (text: string, fontSize: number) => { above: number; below: number };
  /** Encode `text` as a hex string of 4-digit glyph ids (Identity-H addressing). */
  readonly encodeTextAsCidHex: (text: string) => string;
  /**
   * The complete PDF text-showing operator for `text` — `<hex> Tj`, or a
   * `[…] TJ` array when shaping moved a glyph off its own advance.
   */
  readonly showText: (text: string) => string;
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
  // ISO 32000-1 §9.4.3 — a `Tj` advances the cursor by the glyph's OWN width,
  // the one the /W array declares. Kerning is not part of that width: it is an
  // adjustment the show operator has to carry, and a `Tj` has nowhere to put
  // one. So a run measured WITH kerning and shown with `Tj` is drawn wider
  // than layout believed — invisible while the font's advances carry the
  // cursor, and plain to see the moment a line is positioned token by token,
  // where each token lands at the measured width of the ones before it. It
  // did: HeaderFooterUnicode.docx printed "L'Avareou l'École du mensonge",
  // its space eaten by the L-apostrophe pair kern in the token before it.
  const showText = (text: string): string => {
    const shaped = shapeText(
      text,
      parsed.glyphForCodepoint,
      parsed.advanceWidths,
      parsed.ligatures,
      kerning,
      parsed.joiningForms,
    );
    const parts: Array<string> = [];
    let hex = '';
    let kerned = false;
    for (let i = 0; i < shaped.gids.length; i++) {
      const gid = shaped.gids[i]!;
      hex += gid.toString(16).padStart(4, '0').toUpperCase();
      // The shaped advance less the glyph's own: whatever pair kerning added.
      const delta = (shaped.advances[i] ?? 0) - (parsed.advanceWidths[gid] ?? 0);
      if (delta === 0) continue;
      // §9.4.3 subtracts the number, in thousandths of text space, from the
      // displacement — so a tightening (negative) kern becomes a positive one.
      const adj = (-delta * 1000) / parsed.unitsPerEm;
      parts.push(`<${hex}>`, String(Number(adj.toFixed(2))));
      hex = '';
      kerned = true;
    }
    if (!kerned) return `<${hex}> Tj`;
    if (hex !== '') parts.push(`<${hex}>`);
    return `[${parts.join(' ')}] TJ`;
  };

  // A glyph record opens with its own bounding box: numberOfContours, then
  // xMin/yMin/xMax/yMax, all int16 (ISO/IEC 14496-22 §5.3.3). An empty record
  // is a blank glyph — a space — and contributes no ink at all.
  const glyf = parsed.tables.get('glyf');
  const glyphInk = (gid: number): { yMin: number; yMax: number } | undefined => {
    if (!glyf || gid < 0 || gid >= parsed.numGlyphs) return undefined;
    const start = parsed.glyphOffsets[gid] ?? 0;
    const end = parsed.glyphOffsets[gid + 1] ?? 0;
    if (end - start < 10) return undefined;
    const at = glyf.offset + start;
    const i16 = (o: number): number => {
      const v = ((parsed.raw[at + o] ?? 0) << 8) | (parsed.raw[at + o + 1] ?? 0);
      return v >= 0x8000 ? v - 0x10000 : v;
    };
    return { yMin: i16(4), yMax: i16(8) };
  };
  const textInkPt = (text: string, fontSize: number): { above: number; below: number } => {
    const shaped = shapeText(
      text,
      parsed.glyphForCodepoint,
      parsed.advanceWidths,
      parsed.ligatures,
      kerning,
      parsed.joiningForms,
    );
    let yMax = -Infinity;
    let yMin = Infinity;
    for (const gid of shaped.gids) {
      const ink = glyphInk(gid);
      if (!ink) continue;
      if (ink.yMax > yMax) yMax = ink.yMax;
      if (ink.yMin < yMin) yMin = ink.yMin;
    }
    // A CFF font has no `glyf` to read, and a string of nothing but spaces has
    // no ink in any font: fall back to the metrics that always exist.
    if (!Number.isFinite(yMax) || !Number.isFinite(yMin)) {
      yMax = parsed.capHeight || parsed.ascender;
      yMin = 0;
    }
    const k = fontSize / parsed.unitsPerEm;
    return { above: yMax * k, below: -yMin * k };
  };

  return { pdfWidthForGid, textWidthPt, textInkPt, encodeTextAsCidHex, showText };
}
