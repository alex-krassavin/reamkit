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
   *
   * For such text, `before` puts an operator of the writer's own — a colour,
   * say — before a glyph, keyed by the glyph's index in {@link glyphRun}'s
   * order: the show is cut there, its pen moves and rise kept.
   */
  readonly showText: (
    text: string,
    fontSizePt?: number,
    clusterSpacingPt?: number,
    before?: ReadonlyMap<number, string>,
  ) => string;
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
  /**
   * For text the Indic shaper set, per glyph in {@link glyphRun}'s order, the
   * index in code points of the first character it stands for: the one a
   * writer takes a glyph's formatting from. Undefined for any other text.
   */
  readonly glyphSources: (text: string) => ReadonlyArray<number> | undefined;
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
  return { gids: run.gids, glyphText: shapedGlyphText(parsed, text, run) };
}

/** The characters a face's cmap gives each glyph, scanned as /ToUnicode scans them: U+0009 to U+FFFF. */
const cmapNames = new WeakMap<ParsedTtf, ReadonlyMap<number, ReadonlyArray<number>>>();

function namesOf(parsed: ParsedTtf): ReadonlyMap<number, ReadonlyArray<number>> {
  let names = cmapNames.get(parsed);
  if (!names) {
    const map = new Map<number, Array<number>>();
    for (let cp = 0x0009; cp <= 0xffff; cp++) {
      if (cp >= 0xd800 && cp <= 0xdfff) continue;
      const gid = parsed.glyphForCodepoint(cp);
      if (gid === 0) continue;
      const had = map.get(gid);
      if (had) had.push(cp);
      else map.set(gid, [cp]);
    }
    names = map;
    cmapNames.set(parsed, names);
  }
  return names;
}

/**
 * The characters a face's `cmap` sends to one glyph, when it sends more than
 * one: a /ToUnicode can name only one of them, and which one is a question of
 * what the document drew.
 *
 * Arimo, Tinos and Cousine draw the Latin schwa U+0259 and the Cyrillic one
 * U+04D9 with the same glyph (U+018F and U+04D8 likewise), and a scan that
 * keeps the first code reaching a glyph named it Latin: every "және" of a
 * Kazakh document set in Arial read back as "жəне".
 *
 * @param parsed The face.
 * @param gid    The glyph.
 * @returns The code points mapped to it, lowest first, or `undefined` when at
 *          most one is.
 */
export function sharedGlyphCodes(
  parsed: ParsedTtf,
  gid: number,
): ReadonlyArray<number> | undefined {
  const codes = namesOf(parsed).get(gid);
  return codes !== undefined && codes.length > 1 ? codes : undefined;
}

/** A character's canonical parts: क़ is क and the nukta, ো is ে and া. */
const partsOf = (cp: number): Array<number> =>
  [...String.fromCodePoint(cp).normalize('NFD')].map((ch) => ch.codePointAt(0)!);

const IGNORABLE_RE = /\p{Default_Ignorable_Code_Point}/u;

/**
 * Each glyph the shaper made that no character maps to, with the characters
 * it was made from, in the order they were written: the reph of पार्श्व
 * stands for र्, the conjunct it rides on for श्व, the short i of हि for ि.
 * Each pair is one cluster, and a glyph given the cluster's text less what
 * its other glyphs are named after read the tail of पार्श्व twice: once for
 * the conjunct and once for the reph.
 *
 * A cluster is read glyph by glyph:
 * - A glyph a character maps to is named after it by /ToUnicode, and reads
 *   that much of what it was made from — or, where a face moved a glyph by
 *   swapping two (Noto Sans Bengali sets a reph after a vowel sign so), of
 *   the rest of its cluster.
 * - A made glyph reads what is left of the characters it was made from. Of a
 *   character drawn in several, each takes a part: a two-part vowel sign
 *   whose first part joined the consonant stands there for its first part,
 *   and the glyph of the second part for the second.
 * - What is left of the cluster goes to a made glyph that read nothing, or
 *   else to the one after the glyph it was drawn as: the o of कों is set as
 *   the glyph of the aa, its top in the anusvara's, and the anusvara stands
 *   for both. Left out, a reader lost the o of a plural.
 * - A made glyph that still reads nothing — one of the pieces a face draws a
 *   letter in — stands for its letter, so that every glyph drawn reads.
 */
function shapedGlyphText(parsed: ParsedTtf, text: string, run: IndicRun): Map<number, string> {
  const cps = [...text].map((ch) => ch.codePointAt(0)!);
  const names = namesOf(parsed);
  // The missing glyph stands for whatever the face lacks, and is named after nothing.
  const made = run.gids.map((gid) => gid !== 0 && !names.has(gid));
  // The characters of each cluster, and the cluster each character is in.
  const starts = [...new Set(run.clusters)].sort((a, b) => a - b);
  const members = new Map<number, Array<number>>();
  const clusterOf: Array<number> = [];
  starts.forEach((c, k) => {
    const all: Array<number> = [];
    for (let s = k === 0 ? 0 : c; s < (starts[k + 1] ?? cps.length); s++) {
      all.push(s);
      clusterOf[s] = c;
    }
    members.set(c, all);
  });
  // The glyphs each character went into, and the made glyphs of each
  // cluster, in the order they are drawn.
  const carriers = cps.map((): Array<number> => []);
  const madeIn = new Map<number, Array<number>>();
  run.sources.forEach((own, i) => {
    for (const s of own) carriers[s]?.push(i);
    if (!made[i]) return;
    const had = madeIn.get(run.clusters[i]!);
    if (had) had.push(i);
    else madeIn.set(run.clusters[i]!, [i]);
  });
  // What of each character is still to be read, as its canonical parts.
  const unread = cps.map(partsOf);
  // Take the parts of a name out of what some characters have unread, each
  // from the first that has it — or the last, `late`. False if one is missing.
  const take = (
    parts: ReadonlyArray<number>,
    from: ReadonlyArray<number>,
    late: boolean,
  ): boolean => {
    const pool = from.map((s) => [...unread[s]!]);
    for (const part of parts) {
      let q = -1;
      pool.forEach((have, at) => {
        if (have.includes(part) && (q < 0 || late)) q = at;
      });
      if (q < 0) return false;
      pool[q]!.splice(pool[q]!.indexOf(part), 1);
    }
    from.forEach((s, at) => (unread[s] = pool[at]!));
    return true;
  };

  // The characters the made glyphs drawn so far were made from.
  const madeFrom = new Set<number>();
  run.gids.forEach((gid, j) => {
    const own = run.sources[j]!;
    if (made[j]) for (const s of own) madeFrom.add(s);
    const said = names.get(gid);
    if (!said) return;
    // Of two characters with the part, a glyph drawn after a made glyph of
    // the same ones reads the later: of ડ્ડ cut in two, the half form drawn
    // first is the first ડ, and the full ડ after it the second.
    const late = own.some((s) => madeFrom.has(s));
    if (said.some((name) => take(partsOf(name), own, late))) return;
    const others = (members.get(run.clusters[j]!) ?? []).filter((s) => !own.includes(s));
    if (said.some((name) => take(partsOf(name), others, late))) return;
    // A ligature named after another spelling of its letters — ত্‍ as ৎ —
    // reads them all.
    if (own.length > 1 && !said.some((name) => IGNORABLE_RE.test(String.fromCodePoint(name)))) {
      for (const s of own) unread[s] = [];
    }
  });

  // What each made glyph reads, piece by piece, with the character each is from.
  const pieces = run.gids.map((): Array<readonly [number, string]> => []);
  // What is left of a character, its parts put back together: ಯೋ less the ೕ
  // its own glyph reads is ಯೊ, as where it was written ಯೊ.
  const rest = (s: number): string => {
    const left = unread[s]!;
    if (left.length === partsOf(cps[s]!).length) return String.fromCodePoint(cps[s]!);
    return String.fromCodePoint(...left).normalize('NFC');
  };
  cps.forEach((_, s) => {
    const left = unread[s]!;
    const mine = carriers[s]!.filter((i) => made[i]);
    if (left.length === 0 || mine.length === 0) return;
    if (mine.length === 1) pieces[mine[0]!]!.push([s, rest(s)]);
    else {
      // One part each, in the order the parts are written and drawn; the last
      // takes what is left over.
      mine.forEach((i, k) => {
        const part = k === mine.length - 1 ? left.slice(k) : left.slice(k, k + 1);
        if (part.length > 0) pieces[i]!.push([s, String.fromCodePoint(...part).normalize('NFC')]);
      });
    }
    unread[s] = [];
  });
  cps.forEach((cp, s) => {
    if (unread[s]!.length === 0 || IGNORABLE_RE.test(String.fromCodePoint(cp))) return;
    const near = madeIn.get(clusterOf[s]!) ?? [];
    const after = carriers[s]![0] ?? -1;
    const to =
      near.find((i) => pieces[i]!.length === 0) ?? near.find((i) => i > after) ?? near.at(-1);
    if (to !== undefined) pieces[to]!.push([s, rest(s)]);
  });

  const out = new Map<number, string>();
  run.gids.forEach((gid, i) => {
    if (!made[i] || out.has(gid)) return;
    const read =
      pieces[i]!.length > 0
        ? pieces[i]!
        : run.sources[i]!.map((s) => [s, String.fromCodePoint(cps[s]!)] as const);
    const said = [...read]
      .sort((a, b) => a[0] - b[0])
      .map(([, t]) => t)
      .join('');
    if (said !== '') out.set(gid, said);
  });
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
  const showText = (
    text: string,
    fontSizePt?: number,
    clusterSpacingPt?: number,
    before?: ReadonlyMap<number, string>,
  ): string => {
    const shaped = shape(text);
    if (shaped.indic) return showPlaced(shaped.indic, fontSizePt, clusterSpacingPt ?? 0, before);
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
  // after every glyph and pull a vowel sign off its consonant. An operator the
  // writer puts before a glyph ends the show there, the pen moves owed so far
  // with it, and the next show carries on from where that one left the pen.
  const showPlaced = (
    run: IndicRun,
    fontSizePt: number | undefined,
    spacingPt: number,
    before: ReadonlyMap<number, string> | undefined,
  ): string => {
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
      const op = before?.get(i);
      if (op !== undefined) {
        flushShift();
        flushShow();
        ops.push(op);
      }
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
  const glyphSources = (text: string): ReadonlyArray<number> | undefined => {
    const run = shape(text).indic;
    // A dotted circle the shaper put in stands for nothing: it goes with the
    // sign it carries, which begins its cluster.
    return run?.sources.map((s, i) => s[0] ?? run.clusters[i]!);
  };
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
    glyphSources,
    spacingSlots,
  };
}
