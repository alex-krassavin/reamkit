// §9.6.2.2 — the widths of a font the file does not measure.
//
// The 14 standard Type 1 fonts need be neither embedded nor given a `/Widths`
// array, "because a conforming reader shall have built-in font metrics for
// them". Without them every advance in such a file is a guess: basicapi.pdf
// sets "Chapter 1 " in 12pt Helvetica, which is 56pt of type, and a flat
// 500/1000 made it 70.5 — a quarter again too wide, and with it the line's
// width, the gaps between its runs, its alignment and the column it is in.
//
// The tables are Adobe's own ({@link STANDARD_METRICS}); this is the reading of
// them. A width is looked up by the CHARACTER the code decodes to, so whatever
// `/Encoding` the file states — WinAnsi, MacRoman, a `/Differences` list — has
// already been applied. Where the glyph has no character to be looked up by
// (`Lslash`, and every glyph of Symbol and ZapfDingbats), the code Adobe's own
// encoding gives it stands in.

import { COURIER_WIDTH, STANDARD_METRICS } from './standard-metrics';
import { textForGlyphName } from './glyph-names';

/** A face's metrics, in the two ways a code can reach them. */
interface Metrics {
  readonly byChar: ReadonlyMap<string, number>;
  readonly byCode: ReadonlyMap<number, number>;
}

/** Parsed on first use and kept: a face is asked for once per font, not once per glyph. */
const parsed = new Map<string, Metrics>();

function metricsOf(face: string): Metrics {
  const had = parsed.get(face);
  if (had) return had;
  const byChar = new Map<string, number>();
  const byCode = new Map<number, number>();
  const triples = (STANDARD_METRICS[face] ?? '').split(' ');
  for (let i = 0; i + 2 < triples.length; i += 3) {
    const name = triples[i]!;
    const code = Number(triples[i + 1]);
    const width = Number(triples[i + 2]);
    if (!Number.isFinite(width)) continue;
    if (Number.isFinite(code) && code >= 0) byCode.set(code, width);
    const text = textForGlyphName(name);
    // First name wins: Adobe lists `space` before `spacehackarabic`, and the
    // width is the same either way, but the order keeps the table stable.
    if (text !== undefined && text.length > 0 && !byChar.has(text)) byChar.set(text, width);
  }
  const made = { byChar, byCode };
  parsed.set(face, made);
  return made;
}

/**
 * Which of the standard 14 a `/BaseFont` names, or `undefined` for a face this
 * has no metrics for.
 *
 * The four obliques are the upright slanted rather than respaced, so they share
 * its table, and every Courier is 600 flat. The names Arial, Times New Roman
 * and Courier New are matched too: a file that names one and embeds nothing is
 * asking for the metric-compatible face every reader substitutes, and giving it
 * anything else measures the page wrong.
 */
export function standardFace(baseFont: string): string | undefined {
  // The subset prefix (`AAAAAE+`) and any style suffix a producer appends.
  const name = baseFont.replace(/^[A-Z]{6}\+/u, '').toLowerCase();
  const bold = /bold|black|heavy|[-,]bd\b/u.test(name);
  const italic = /italic|oblique|[-,]it\b/u.test(name);
  if (/courier|mono/u.test(name)) return 'Courier';
  if (/times|serif|roman|georgia|book/u.test(name) && !/sans/u.test(name)) {
    return bold
      ? italic
        ? 'Times-BoldItalic'
        : 'Times-Bold'
      : italic
        ? 'Times-Italic'
        : 'Times-Roman';
  }
  if (/zapf|dingbat/u.test(name)) return 'ZapfDingbats';
  if (/symbol/u.test(name)) return 'Symbol';
  if (/helvetica|arial|sans|verdana|tahoma|segoe|calibri/u.test(name)) {
    return bold ? 'Helvetica-Bold' : 'Helvetica';
  }
  return undefined;
}

/**
 * The advance of one code in a standard face, in 1000ths of an em.
 *
 * @param face   One of the names {@link standardFace} returns.
 * @param code   The character code, as the content stream showed it.
 * @param text   What that code decodes to, when the font could say.
 * @returns The width, or `undefined` where the face states none for it.
 */
export function standardWidth(face: string, code: number, text: string): number | undefined {
  if (face === 'Courier') return COURIER_WIDTH;
  const metrics = metricsOf(face);
  // One character, because a width is one glyph's: a code that decoded to a
  // ligature's two letters is not in the table under either of them.
  const byChar = text.length > 0 ? metrics.byChar.get(text) : undefined;
  return byChar ?? metrics.byCode.get(code);
}
