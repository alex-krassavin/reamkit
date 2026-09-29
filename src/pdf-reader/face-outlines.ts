// §9.9 — the outlines each face drew the document's characters with, gathered
// so a writer can EMBED the face.
//
// A reconstructed document names its faces, and a reader without them sets the
// text in a substitute: a fraction wider or narrower, so every line breaks in
// another place and a page that fit on one sheet runs onto two. The program the
// page was drawn with is in the file, and what a word processor needs of it is
// small: for each character the document shows, the outline the page drew and
// the advance it set it with.
//
// Only the codes the page actually PAINTS are read (the interpreter keeps
// them, see `InterpretResult.shown`), so what is embedded is the subset the
// document uses — and a scanned page's invisible words, shown in a face with no
// ink in it, embed nothing.

import { ASCENDER, NATURAL_LINE_EM } from './flow-build';
import { isRightToLeft } from './content';
import type { ContentFont, FaceProgram, TextRun } from './content';
import type { FaceGlyph, FaceOutlines } from '@/core/ir/flow';
import { editableEmbedding } from '@/core/font';

/** The codes each font showed where the pages paint them, gathered across pages. */
export type ShownCodes = Map<ContentFont, Set<number>>;

/**
 * Fold one content stream's shown codes into the document's.
 *
 * @param into The document's codes, added to.
 * @param from One stream's (`InterpretResult.shown`).
 */
export function addShown(
  into: ShownCodes,
  from: ReadonlyMap<ContentFont, ReadonlySet<number>>,
): void {
  for (const [font, codes] of from) {
    let set = into.get(font);
    if (!set) into.set(font, (set = new Set()));
    for (const code of codes) set.add(code);
  }
}

/**
 * The outlines of every face the pages painted characters in, keyed by the name
 * a run in the face carries.
 *
 * A face is left out where its program cannot be trusted to say what the page
 * showed: where more than a fifth of its characters have no glyph the reader
 * can find, the way to its glyphs is not the one the page took, and a face
 * embedded that way would show the wrong letters in place of a substitute's
 * right ones. A character whose glyph cannot be found is left out alone — a
 * reader draws it in another face.
 *
 * So is a face that shows a script which is SHAPED: an Arabic letter takes one
 * of four forms by its neighbours, and a face rebuilt here carries one glyph
 * per character and none of the rules that choose between them. Nor can it
 * carry the part of such a face it could: ArabicCIDTrueType.pdf maps half its
 * glyphs to presentation forms, which are shaped already, and half to plain
 * letters — embedded, each line came out in two faces, a bold word with light
 * letters in it.
 *
 * @param shown   The codes the pages painted, per font.
 * @param spacing How the pages space each face's text apart where they move the
 *                pen rather than draw a glyph (see {@link pageSpacing}).
 * @returns Run font name → the face's outlines.
 */
export function faceOutlinesOf(
  shown: ShownCodes,
  spacing: PageSpacing = NO_SPACING,
): Map<string, FaceOutlines> {
  const faces = new Map<string, Gathered>();
  for (const [font, codes] of shown) {
    const program = font.program;
    if (!program || font.name === undefined) continue;
    let face = faces.get(font.name);
    if (!face) {
      face = {
        font,
        program,
        glyphs: new Map(),
        ligatures: new Map(),
        sought: 0,
        missed: 0,
        shaped: false,
      };
      faces.set(font.name, face);
    }
    // The strictest licence any of the face's programs states.
    if (
      program.fsType !== undefined &&
      (face.fsType === undefined || editableEmbedding(face.fsType))
    ) {
      face.fsType = program.fsType;
    }
    for (const code of [...codes].sort((a, b) => a - b)) gather(face, font, program, code);
  }
  const out = new Map<string, FaceOutlines>();
  for (const [name, face] of faces) {
    const outlines = outlinesOf(face, spacing.wordGaps.get(name), spacing.kerning.get(name));
    if (outlines) out.set(name, outlines);
  }
  return out;
}

/**
 * How a page spaces a face's text apart where it moves the pen rather than
 * draws a glyph — between words, and between two letters of one word.
 */
export interface PageSpacing {
  /** Face → the white it leaves between words, in thousandths of an em. */
  readonly wordGaps: ReadonlyMap<string, number>;
  /**
   * Face → the pairs of letters it is KERNED by: the two characters → how far
   * the second stands from where the first's advance leaves it, in thousandths
   * of an em (negative tightens).
   */
  readonly kerning: ReadonlyMap<string, ReadonlyMap<string, number>>;
}

const NO_SPACING: PageSpacing = { wordGaps: new Map(), kerning: new Map() };

/**
 * §9.4.3 — what the gaps between a face's runs say about how it is spaced.
 *
 * The interpreter emits a run for every string of a `TJ` array, so where the
 * page nudges the pen inside a word — which is how a producer KERNS it — two
 * runs meet with the nudge between them, and the letters on either side are
 * the pair. bigboundingbox.pdf kerns its Calibri by eight pairs in one short
 * line, and set without them the line came out 0.67pt longer than the page's
 * and no longer fit the cell it was measured into. Each pair is the median of
 * every nudge the page gives it.
 *
 * A wider gap is the white between two words, where the page moves the pen
 * instead of showing a space; the face's space is the lower quartile of them,
 * which a justified line's stretched gaps do not reach.
 *
 * @param pages The runs each page set, in the order they were painted.
 * @returns Per face, its word gap and its kerning pairs.
 */
export function pageSpacing(pages: ReadonlyArray<ReadonlyArray<TextRun>>): PageSpacing {
  const gaps = new Map<string, Array<number>>();
  // Each nudge in thousandths of an em, and in points, which is what says
  // whether it is a kern or a producer's rounding.
  const nudges = new Map<string, Map<string, Array<[number, number]>>>();
  // …and the nudges beside a space, in points: no face is kerned against its
  // space, so a page that nudges those too is spacing its letters out.
  const beside = new Map<string, Array<number>>();
  for (const runs of pages) {
    for (let i = 1; i < runs.length; i++) {
      const a = runs[i - 1]!;
      const b = runs[i]!;
      const name = a.fontName;
      if (name === undefined || b.fontName !== name || a.angleDeg !== undefined) continue;
      if (b.angleDeg !== undefined || a.fontSizePt !== b.fontSizePt || !(a.fontSizePt > 0))
        continue;
      if (Math.abs(a.y - b.y) > a.fontSizePt * SAME_LINE_EM) continue;
      // A pair painted right to left meets the other way round; a space the
      // run brings is a gap it measures itself.
      if (isRightToLeft(a.text) || isRightToLeft(b.text)) continue;
      const left = [...a.text].at(-1);
      const right = [...b.text][0];
      if (left === undefined || right === undefined) continue;
      const gap = ((b.x - a.endX) / a.fontSizePt) * 1000;
      if (/\s/u.test(left) || /\s/u.test(right)) {
        if (gap > MIN_KERN && gap < MAX_KERN) {
          let list = beside.get(name);
          if (!list) beside.set(name, (list = []));
          list.push(b.x - a.endX);
        }
        continue;
      }
      if (gap >= MIN_WORD_GAP && gap <= MAX_WORD_GAP) {
        let list = gaps.get(name);
        if (!list) gaps.set(name, (list = []));
        list.push(gap);
      } else if (gap > MIN_KERN && gap < MAX_KERN) {
        let pairs = nudges.get(name);
        if (!pairs) nudges.set(name, (pairs = new Map()));
        const pair = `${left}${right}`;
        let list = pairs.get(pair);
        if (!list) pairs.set(pair, (list = []));
        list.push([gap, b.x - a.endX]);
      }
    }
  }
  const wordGaps = new Map<string, number>();
  for (const [name, list] of gaps) {
    if (list.length < MIN_GAP_SAMPLES) continue;
    list.sort((p, q) => p - q);
    wordGaps.set(name, list[Math.floor(list.length / 4)]!);
  }
  const kerning = new Map<string, Map<string, number>>();
  for (const [name, pairs] of nudges) {
    if (
      tracked(
        [...pairs.values()].flat().map(([, points]) => points),
        beside.get(name) ?? [],
      )
    ) {
      continue;
    }
    const kept = new Map<string, number>();
    for (const [pair, list] of pairs) {
      const inEm = median(list.map(([em]) => em));
      const inPt = median(list.map(([, points]) => points));
      if (Math.abs(inEm) < MIN_KERN_VALUE || Math.abs(inPt) < MIN_KERN_PT) continue;
      // A kern is the same every time the pair meets; a rounding is not. One
      // sighting says nothing unless it is too wide to be a rounding at all.
      const steady =
        list.length >= 2 && list.every(([, points]) => Math.abs(points - inPt) <= TWIP_PT);
      if (steady || Math.abs(inPt) >= SURE_KERN_PT) kept.set(pair, inEm);
    }
    // A face nudged in fewer pairs than this was not kerned: its producer
    // rounded, and the pairs are its rounding.
    if (kept.size >= MIN_KERNED_PAIRS) kerning.set(name, kept);
  }
  return { wordGaps, kerning };
}

/**
 * Whether a face's nudges are TRACKING rather than kerning: one amount between
 * nearly every two letters, and beside its spaces too, where no face is kerned.
 * bug1157493.pdf sets a line of Courier half a tenth of an em tight, every
 * letter and every space alike; read as kerning pairs, its letters closed up
 * and its spaces did not.
 *
 * @param nudges The face's nudges between letters, in points.
 * @param beside Its nudges beside a space, in points.
 */
function tracked(nudges: ReadonlyArray<number>, beside: ReadonlyArray<number>): boolean {
  if (beside.length < MIN_TRACKED_SPACES || nudges.length === 0) return false;
  const typical = median([...nudges, ...beside]);
  if (Math.abs(typical) < MIN_KERN_PT) return false;
  const near = (points: number): boolean => Math.abs(points - typical) <= TWIP_PT;
  const all = [...nudges, ...beside];
  return (
    beside.filter(near).length >= MIN_TRACKED_SPACES &&
    all.filter(near).length >= all.length * TRACKED_SHARE
  );
}

/** How many nudges beside spaces, and what share of all, make a face tracked. */
const MIN_TRACKED_SPACES = 2;
const TRACKED_SHARE = 0.8;

/**
 * The faces a page KERNED: those with kerning pairs (see {@link pageSpacing},
 * which keeps a face's pairs only where there are enough of them to say it was
 * kerning and not an accident of how its strings were cut).
 *
 * @param spacing How the pages space their faces (see {@link pageSpacing}).
 * @returns The run font names of the kerned faces.
 */
export function kernedFaces(spacing: PageSpacing): Set<string> {
  return new Set(spacing.kerning.keys());
}

/** The middle of some numbers: the mean of the middle two where they are even. */
function median(values: ReadonlyArray<number>): number {
  const sorted = [...values].sort((p, q) => p - q);
  const mid = sorted.length / 2;
  return sorted.length % 2 === 1 ? sorted[Math.floor(mid)]! : (sorted[mid - 1]! + sorted[mid]!) / 2;
}

/** Two runs within this much of an em of each other's baseline share a line. */
const SAME_LINE_EM = 0.1;

/** A gap between words: wider than a kern, narrower than a column's white. */
const MIN_WORD_GAP = 100;
const MAX_WORD_GAP = 600;

/**
 * A nudge inside a word, in thousandths of an em: kerning tightens a pair by as
 * much as a quarter of an em and opens one by far less — a gap wider than this
 * is a space the page did not draw.
 */
const MIN_KERN = -250;
const MAX_KERN = 60;

/** A nudge smaller than this is the rounding of a producer's positions, not a kern. */
const MIN_KERN_VALUE = 2;

/**
 * …and so is one smaller than a twip and a half, in points: Word lays a line
 * out in twips and writes the difference from the face's own widths as a
 * nudge. annotation-highlight.pdf's Calibri is nudged three to six thousandths
 * of an em between letters no kerning table pairs, and taken for kerns they
 * moved its words a fraction of a pixel off the page's.
 */
const MIN_KERN_PT = 0.075;

/** A twip, in points: the grid a Word page's positions are rounded to. */
const TWIP_PT = 0.05;

/**
 * A nudge this wide is a kern seen once: rounding moves a letter a twip or
 * two, and bug793632.pdf's Calibri is nudged two twips either way — "st" in
 * and "on" out — once each, by no kerning at all.
 */
const SURE_KERN_PT = 0.15;

/** How many pairs make a face one the page kerned. */
const MIN_KERNED_PAIRS = 2;

/** How many gaps it takes to say how wide a face's space is. */
const MIN_GAP_SAMPLES = 3;

/** One face's glyphs as they are gathered, and how many of them were found. */
interface Gathered {
  /** The first font seen in the face, whose style and metrics stand for it. */
  readonly font: ContentFont;
  readonly program: FaceProgram;
  readonly glyphs: Map<string, FaceGlyph>;
  /** The glyphs the face draws for a run of letters at once, by the letters. */
  readonly ligatures: Map<string, FaceGlyph>;
  fsType?: number;
  /** Characters with ink that were looked for, and those not found. */
  sought: number;
  missed: number;
  /** Whether the face shows a letter of a shaped script. */
  shaped: boolean;
}

/** Read one code's glyph into its face, where it stands for one character or a ligature. */
function gather(face: Gathered, font: ContentFont, program: FaceProgram, code: number): void {
  const chars = [...font.decode([code])];
  if (chars.length > 1) {
    gatherLigature(face, font, program, code, chars);
    return;
  }
  if (chars.length !== 1) return; // no character at all
  const char = chars[0]!;
  if (face.glyphs.has(char) || !embeddable(char)) return;
  if (shaped(char)) {
    face.shaped = true;
    return;
  }
  const blank = WHITE.test(char);
  if (!blank) face.sought++;
  const outline = program.glyph(code);
  // A character with ink whose glyph draws nothing is a glyph not found: the
  // page showed nothing there, or the way to its glyph is not this one.
  if (outline === undefined || (outline.length === 0 && !blank)) {
    if (!blank) face.missed++;
    return;
  }
  face.glyphs.set(char, { outline, advance: font.width(code) });
}

/**
 * §9.10.2 — a code that stands for a run of letters is a LIGATURE: one glyph
 * the page drew for "fi", "ffl" or Calibri's "tt". The text carries the letters
 * and the rebuilt face the glyph, so a reader that joins them draws what the
 * page drew — set apart, attachment.pdf's "attachment" came out with two
 * crossbars where the page has one.
 *
 * Only letters, and no more than {@link MAX_LIGATURE} of them: a producer that
 * maps a glyph to a longer string — a logo to a company's name — has not drawn
 * a ligature.
 */
function gatherLigature(
  face: Gathered,
  font: ContentFont,
  program: FaceProgram,
  code: number,
  chars: ReadonlyArray<string>,
): void {
  const key = chars.join('');
  if (chars.length > MAX_LIGATURE || face.ligatures.has(key)) return;
  if (!chars.every((c) => /^\p{L}$/u.test(c) && !shaped(c))) return;
  const outline = program.glyph(code);
  if (outline === undefined || outline.length === 0) return;
  face.ligatures.set(key, { outline, advance: font.width(code) });
}

/** The most letters one ligature joins: ffi, ffl. */
const MAX_LIGATURE = 3;

/** A face's gathered glyphs as {@link FaceOutlines}, where it is one to embed. */
function outlinesOf(
  face: Gathered,
  gap: number | undefined,
  kerning: ReadonlyMap<string, number> | undefined,
): FaceOutlines | undefined {
  if (face.shaped || face.sought === 0 || face.missed > face.sought * MAX_MISSED_SHARE) {
    return undefined;
  }
  const glyphs = new Map(face.glyphs);
  const program = face.program;
  // The space is rarely a glyph a page shows — it moves the pen instead — and
  // a writer puts one between every two words.
  if (!glyphs.has(' ')) glyphs.set(' ', { outline: [], advance: spaceAdvance(face, gap) });
  if (!glyphs.has('\u00a0')) glyphs.set('\u00a0', glyphs.get(' ')!);
  const { yMin, yMax } = extent(glyphs);
  // The pairs whose two glyphs the face carries: a kern for a letter drawn in
  // another face would move nothing of this one's.
  const pairs = new Map(
    [...(kerning ?? [])].filter(([pair]) => [...pair].every((c) => glyphs.has(c))),
  );
  // …and so for a ligature: one of letters the face does not carry alone is
  // one a reader never lays out in it.
  const ligatures = new Map(
    [...face.ligatures].filter(([letters]) => [...letters].every((c) => glyphs.has(c))),
  );
  return {
    glyphs,
    ...(pairs.size > 0 ? { kerning: pairs } : {}),
    ...(ligatures.size > 0 ? { ligatures } : {}),
    ...(face.fsType !== undefined ? { fsType: face.fsType } : {}),
    // The slot a run in the face looks the face up in, whatever the program
    // calls itself: bug900822.pdf sets `LucidaSansUnicode,Bold` in the
    // regular program, and the page shows it regular — embedded as the
    // family's regular, a reader thickened it into a bold the page never drew.
    bold: face.font.bold === true,
    italic: face.font.italic === true,
    postScriptName: program.postScriptName,
    // The line the reconstruction measured the page against, not the one the
    // descriptor states: a paragraph whose pitch nobody measured is set in the
    // face's own line, and its neighbours were spaced for one of
    // `NATURAL_LINE_EM` with the baseline `ASCENDER` down it. Set in Calibri's
    // descriptor line — 750 up and 250 down, a line of one em —
    // annotation-highlight.pdf's second paragraph rose a fifth of a line. The
    // glyphs' own reach stretches it, as a tall capital would.
    ascent: Math.max(ASCENDER * 1000, yMax),
    descent: Math.min((ASCENDER - NATURAL_LINE_EM) * 1000, yMin),
    ...(program.capHeight !== undefined ? { capHeight: program.capHeight } : {}),
    ...(program.xHeight !== undefined ? { xHeight: program.xHeight } : {}),
    italicAngle: program.italicAngle,
    fixedPitch: program.fixedPitch,
  };
}

/** A face may miss no more than this share of the glyphs looked for. */
const MAX_MISSED_SHARE = 0.2;

/**
 * The advance of a space the page never showed: the face's own width for the
 * code a space has in its encoding, where there is one; the width the program
 * gives its space glyph — which a subset keeps though it drops the outline;
 * the white the page itself leaves between the face's words; the one advance
 * every glyph has, in a fixed-pitch face; a quarter of an em otherwise.
 * bigboundingbox.pdf sets its Calibri in a subset that keeps no space and no
 * way to one, and spaces the words by hand: with a quarter of an em for
 * Calibri's 226 thousandths, "Enter the amount you are paying above" no longer
 * fit its cell.
 */
function spaceAdvance(face: Gathered, gap: number | undefined): number {
  const font = face.font;
  if (font.bytesPerCode === 1 && font.decode([0x20]) === ' ') {
    const width = font.width(0x20);
    if (width > 0 && width < 1000) return width;
  }
  if (face.program.spaceAdvance !== undefined) return face.program.spaceAdvance;
  if (gap !== undefined) return gap;
  if (face.program.fixedPitch) {
    const first = [...face.glyphs.values()].find((g) => g.advance > 0);
    if (first) return first.advance;
  }
  return 250;
}

/** How high and how low the glyphs' contours reach, in thousandths of an em. */
function extent(glyphs: ReadonlyMap<string, FaceGlyph>): { yMin: number; yMax: number } {
  let yMin = 0;
  let yMax = 0;
  for (const glyph of glyphs.values()) {
    for (const seg of glyph.outline) {
      if (seg.op === 'close') continue;
      yMin = Math.min(yMin, seg.y * 1000);
      yMax = Math.max(yMax, seg.y * 1000);
    }
  }
  return { yMin, yMax };
}

/** Whitespace: a glyph that rightly draws nothing. */
const WHITE = /^\s$/u;

/** Whether a character is one a face carries at all: not a control or a format character. */
function embeddable(char: string): boolean {
  return !/^[\p{Cc}\p{Cf}\p{Cs}\uFFFD]$/u.test(char);
}

/**
 * Whether a character is a letter of a script whose glyphs a face's shaping
 * rules choose — which a rebuilt face does not have.
 */
function shaped(char: string): boolean {
  const cp = char.codePointAt(0)!;
  // Arabic presentation forms are shaped already: each IS one form.
  if ((cp >= 0xfb50 && cp <= 0xfdff) || (cp >= 0xfe70 && cp <= 0xfeff)) return false;
  // Hangul's conjoining jamo combine into syllables; its syllables are whole.
  if ((cp >= 0x1100 && cp <= 0x11ff) || (cp >= 0xa960 && cp <= 0xa97f)) return true;
  if (cp >= 0xd7b0 && cp <= 0xd7ff) return true;
  return SHAPED.test(char);
}

/** The scripts whose glyphs a face's shaping rules choose. */
const SHAPED =
  /\p{Script=Arabic}|\p{Script=Syriac}|\p{Script=Thaana}|\p{Script=Nko}|\p{Script=Mandaic}|\p{Script=Devanagari}|\p{Script=Bengali}|\p{Script=Gurmukhi}|\p{Script=Gujarati}|\p{Script=Oriya}|\p{Script=Tamil}|\p{Script=Telugu}|\p{Script=Kannada}|\p{Script=Malayalam}|\p{Script=Sinhala}|\p{Script=Thai}|\p{Script=Lao}|\p{Script=Tibetan}|\p{Script=Myanmar}|\p{Script=Khmer}|\p{Script=Mongolian}|\p{Script=Javanese}|\p{Script=Balinese}/u;
