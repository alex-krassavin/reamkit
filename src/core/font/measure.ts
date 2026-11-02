// Font measurement/encoding derived purely from the parsed font — no PDF
// involved (ir-design stage 3b / 6.4). The layout phase measures text with
// these; the PDF emitter reuses the same functions (via embedTtfFont) so emit
// encodes exactly what layout measured. CIDs are GIDs (Identity ordering), so
// the hex encoding is writer-agnostic glyph addressing.

import type { IndicRun } from '@/core/font/indic-shaper';
import type { ParsedTtf } from '@/core/font/ttf-parser';
import { hasIndic, shapeIndic } from '@/core/font/indic-shaper';
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
   * The complete PDF text-showing operators for `text` — `<hex> Tj`, or a
   * `[…] TJ` array when shaping moved a glyph off its own advance. Text the
   * Indic shaper set (see {@link clustered}) may place a glyph off its pen
   * position and off the baseline, and takes its letter spacing after each
   * cluster rather than each glyph: give it the size it is set at and that
   * spacing, and leave the text state's `Tc` at 0 for it.
   */
  readonly showText: (text: string, fontSizePt?: number, clusterSpacingPt?: number) => string;
  /**
   * The glyphs `text` is drawn with and the advance each one takes, in font
   * units — exactly what {@link showText} shows, kerning and ligatures
   * included — for a writer that places each glyph itself. For text the Indic
   * shaper set, each glyph's offset off its pen position, and which glyphs end
   * a cluster.
   */
  readonly glyphRun: (text: string) => {
    readonly gids: ReadonlyArray<number>;
    readonly advances: ReadonlyArray<number>;
    readonly xOffsets?: ReadonlyArray<number>;
    readonly yOffsets?: ReadonlyArray<number>;
    readonly clusterEnds?: ReadonlyArray<boolean>;
  };
  /**
   * Whether `text` is set by the Indic shaper in this face: its glyphs are not
   * in the order of its characters, and a cluster of them is one letter that
   * spacing must not split.
   */
  readonly clustered: (text: string) => boolean;
  /** How many places letter spacing goes in `text`: one per character, or per cluster. */
  readonly spacingSlots: (text: string) => number;
}

const EMPTY_KERNING = new Map<string, number>();

/** The most shaped words a face keeps; past it the cache starts over. */
const SHAPED_CACHE = 4096;

/** What a face's shaping gives a string, whichever shaper gave it. */
interface Shaped {
  readonly gids: ReadonlyArray<number>;
  readonly advances: ReadonlyArray<number>;
  /** Set by the Indic shaper only. */
  readonly indic?: IndicRun;
}

/**
 * The glyphs a face draws a string with, and for each glyph the Indic shaper
 * made out of several characters (a conjunct, a reph, a vowel sign's
 * contextual form), the text it stands for — which a PDF's /ToUnicode has to
 * name, since no character maps to it.
 *
 * @param parsed The face.
 * @param text   The text.
 * @returns The glyph ids, and the text of each glyph no character reaches.
 */
export function shapeGlyphs(
  parsed: ParsedTtf,
  text: string,
): { readonly gids: ReadonlyArray<number>; readonly glyphText?: ReadonlyMap<number, string> } {
  const run = hasIndic(text) ? shapeIndic(parsed, text) : undefined;
  if (!run) {
    const shaped = shapeText(
      text,
      parsed.glyphForCodepoint,
      parsed.advanceWidths,
      parsed.ligatures,
      parsed.kerning,
      parsed.joiningForms,
    );
    return { gids: shaped.gids };
  }
  return { gids: run.gids, glyphText: clusterGlyphText(parsed, text, run) };
}

/**
 * Each glyph of a cluster that no character of it maps to, with the
 * cluster's text less the characters that do map to its other glyphs: the
 * short i of हि is a contextual form no character reaches, and it stands for
 * the i once the ha is accounted for.
 */
function clusterGlyphText(parsed: ParsedTtf, text: string, run: IndicRun): Map<number, string> {
  const cps = [...text].map((ch) => ch.codePointAt(0)!);
  const starts = [...new Set(run.clusters)].sort((a, b) => a - b);
  const out = new Map<number, string>();
  for (let s = 0; s < starts.length; s++) {
    const from = starts[s]!;
    const to = starts[s + 1] ?? cps.length;
    const left = cps.slice(from, to);
    const glyphs = run.gids.filter((_, i) => run.clusters[i] === from);
    const unmapped: Array<number> = [];
    for (const gid of glyphs) {
      const at = left.findIndex((cp) => parsed.glyphForCodepoint(cp) === gid);
      if (at >= 0) left.splice(at, 1);
      else unmapped.push(gid);
    }
    const rest = String.fromCodePoint(...left);
    if (rest === '') continue;
    for (const gid of unmapped) if (!out.has(gid)) out.set(gid, rest);
  }
  return out;
}

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
  // The scripts of India are shaped whole — reordered, joined, their marks set
  // by anchors — and a word is measured, broken and drawn several times over,
  // so what the shaper gave it is kept.
  const indicRuns = new Map<string, IndicRun | null>();
  const shape = (text: string): Shaped => {
    if (hasIndic(text)) {
      let run = indicRuns.get(text);
      if (run === undefined) {
        if (indicRuns.size >= SHAPED_CACHE) indicRuns.clear();
        run = shapeIndic(parsed, text, kern) ?? null;
        indicRuns.set(text, run);
      }
      if (run) return { gids: run.gids, advances: run.advances, indic: run };
    }
    return shapeText(
      text,
      parsed.glyphForCodepoint,
      parsed.advanceWidths,
      parsed.ligatures,
      kerning,
      parsed.joiningForms,
    );
  };
  const textWidthPt = (text: string, fontSize: number): number => {
    const shaped = shape(text);
    let totalEm = 0;
    for (const a of shaped.advances) totalEm += a;
    return (totalEm * fontSize) / parsed.unitsPerEm;
  };
  const encodeTextAsCidHex = (text: string): string => {
    const shaped = shape(text);
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
  const showText = (text: string, fontSizePt?: number, clusterSpacingPt?: number): string => {
    const shaped = shape(text);
    if (shaped.indic) return showPlaced(shaped.indic, fontSizePt, clusterSpacingPt ?? 0);
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

  // ISO 32000-1 §9.4.3 again, for glyphs placed off their pen position: the
  // pen moves to each glyph's x offset before it and back after, inside the
  // `TJ` array, and a glyph raised or lowered off the baseline is shown with
  // the text rise (`Ts`, §9.3.7) set to its y offset and put back to 0. Letter
  // spacing goes after the last glyph of each cluster, where `Tc` would put it
  // after every glyph and pull a vowel sign off its consonant.
  const showPlaced = (run: IndicRun, fontSizePt: number | undefined, spacingPt: number): string => {
    const upem = parsed.unitsPerEm;
    const ops: Array<string> = [];
    let parts: Array<string> = [];
    let hex = '';
    let shift = 0;
    let rise = 0;
    const flushShift = (): void => {
      if (shift === 0) return;
      if (hex !== '') parts.push(`<${hex}>`);
      hex = '';
      parts.push(String(Number(shift.toFixed(2))));
      shift = 0;
    };
    const flushShow = (): void => {
      if (hex !== '') parts.push(`<${hex}>`);
      hex = '';
      if (parts.length === 1 && parts[0]!.startsWith('<')) ops.push(`${parts[0]} Tj`);
      else if (parts.length > 0) ops.push(`[${parts.join(' ')}] TJ`);
      parts = [];
    };
    // A pen move of `units` font units to the right, as a TJ number.
    const move = (units: number): void => {
      shift -= (units * 1000) / upem;
    };
    for (let i = 0; i < run.gids.length; i++) {
      const gid = run.gids[i]!;
      const dx = run.xOffsets[i] ?? 0;
      const dy = fontSizePt === undefined ? 0 : ((run.yOffsets[i] ?? 0) * fontSizePt) / upem;
      if (dy !== rise) {
        flushShift();
        flushShow();
        ops.push(`${String(Number(dy.toFixed(3)))} Ts`);
        rise = dy;
      }
      move(dx);
      flushShift();
      hex += gid.toString(16).padStart(4, '0').toUpperCase();
      move((run.advances[i] ?? 0) - (parsed.advanceWidths[gid] ?? 0) - dx);
      const endsCluster = run.clusters[i + 1] !== run.clusters[i];
      if (endsCluster && spacingPt !== 0 && fontSizePt) {
        shift -= (spacingPt * 1000) / fontSizePt;
      }
    }
    // What is left of the shift past the last glyph moves the pen on to where
    // the next token begins.
    flushShift();
    flushShow();
    if (rise !== 0) ops.push('0 Ts');
    return ops.join(' ');
  };

  const glyphRun = (
    text: string,
  ): {
    gids: ReadonlyArray<number>;
    advances: ReadonlyArray<number>;
    xOffsets?: ReadonlyArray<number>;
    yOffsets?: ReadonlyArray<number>;
    clusterEnds?: ReadonlyArray<boolean>;
  } => {
    const shaped = shape(text);
    const run = shaped.indic;
    if (!run) return { gids: shaped.gids, advances: shaped.advances };
    return {
      gids: run.gids,
      advances: run.advances,
      xOffsets: run.xOffsets,
      yOffsets: run.yOffsets,
      clusterEnds: run.clusters.map((c, i) => run.clusters[i + 1] !== c),
    };
  };
  const clustered = (text: string): boolean => shape(text).indic !== undefined;
  const spacingSlots = (text: string): number => {
    const run = shape(text).indic;
    return run ? new Set(run.clusters).size : [...text].length;
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
    const shaped = shape(text);
    let yMax = -Infinity;
    let yMin = Infinity;
    shaped.gids.forEach((gid, i) => {
      const ink = glyphInk(gid);
      if (!ink) return;
      // A mark set above or below by its anchor reaches as far as it was moved.
      const dy = shaped.indic?.yOffsets[i] ?? 0;
      if (ink.yMax + dy > yMax) yMax = ink.yMax + dy;
      if (ink.yMin + dy < yMin) yMin = ink.yMin + dy;
    });
    // A CFF font has no `glyf` to read, and a string of nothing but spaces has
    // no ink in any font: fall back to the metrics that always exist.
    if (!Number.isFinite(yMax) || !Number.isFinite(yMin)) {
      yMax = parsed.capHeight || parsed.ascender;
      yMin = 0;
    }
    const k = fontSize / parsed.unitsPerEm;
    return { above: yMax * k, below: -yMin * k };
  };

  return {
    pdfWidthForGid,
    textWidthPt,
    textInkPt,
    encodeTextAsCidHex,
    showText,
    glyphRun,
    clustered,
    spacingSlots,
  };
}
