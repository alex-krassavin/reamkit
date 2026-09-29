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
 * @param shown The codes the pages painted, per font.
 * @param pages The runs each page set, where a face's words are spaced apart
 *              by moving the pen rather than by a glyph (see `spaceAdvance`).
 * @returns Run font name → the face's outlines.
 */
export function faceOutlinesOf(
  shown: ShownCodes,
  pages: ReadonlyArray<ReadonlyArray<TextRun>> = [],
): Map<string, FaceOutlines> {
  const gaps = wordGaps(pages);
  const faces = new Map<string, Gathered>();
  for (const [font, codes] of shown) {
    const program = font.program;
    if (!program || font.name === undefined) continue;
    let face = faces.get(font.name);
    if (!face) {
      face = { font, program, glyphs: new Map(), sought: 0, missed: 0, shaped: false };
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
    const outlines = outlinesOf(face, gaps.get(name));
    if (outlines) out.set(name, outlines);
  }
  return out;
}

/**
 * The white a page leaves between two words of a face, in thousandths of an
 * em, where it moves the pen instead of showing a space: the gap from one run's
 * end to the next's start on the same line, where neither brings a space of
 * its own. A quarter of them are the narrowest — the lower quartile, which a
 * justified line's stretched gaps do not reach.
 */
function wordGaps(pages: ReadonlyArray<ReadonlyArray<TextRun>>): Map<string, number> {
  const samples = new Map<string, Array<number>>();
  for (const runs of pages) {
    for (let i = 1; i < runs.length; i++) {
      const a = runs[i - 1]!;
      const b = runs[i]!;
      const name = a.fontName;
      if (name === undefined || b.fontName !== name || a.angleDeg !== undefined) continue;
      if (b.angleDeg !== undefined || a.fontSizePt !== b.fontSizePt || !(a.fontSizePt > 0))
        continue;
      if (Math.abs(a.y - b.y) > a.fontSizePt * SAME_LINE_EM) continue;
      if (/\s$/u.test(a.text) || /^\s/u.test(b.text)) continue;
      const gap = ((b.x - a.endX) / a.fontSizePt) * 1000;
      if (gap < MIN_WORD_GAP || gap > MAX_WORD_GAP) continue;
      let list = samples.get(name);
      if (!list) samples.set(name, (list = []));
      list.push(gap);
    }
  }
  const out = new Map<string, number>();
  for (const [name, list] of samples) {
    if (list.length < MIN_GAP_SAMPLES) continue;
    list.sort((p, q) => p - q);
    out.set(name, list[Math.floor(list.length / 4)]!);
  }
  return out;
}

/** Two runs within this much of an em of each other's baseline share a line. */
const SAME_LINE_EM = 0.1;

/** A gap between words: wider than a kern, narrower than a column's white. */
const MIN_WORD_GAP = 100;
const MAX_WORD_GAP = 600;

/** How many gaps it takes to say how wide a face's space is. */
const MIN_GAP_SAMPLES = 3;

/** One face's glyphs as they are gathered, and how many of them were found. */
interface Gathered {
  /** The first font seen in the face, whose style and metrics stand for it. */
  readonly font: ContentFont;
  readonly program: FaceProgram;
  readonly glyphs: Map<string, FaceGlyph>;
  fsType?: number;
  /** Characters with ink that were looked for, and those not found. */
  sought: number;
  missed: number;
  /** Whether the face shows a letter of a shaped script. */
  shaped: boolean;
}

/** Read one code's glyph into its face, where it stands for one character. */
function gather(face: Gathered, font: ContentFont, program: FaceProgram, code: number): void {
  const chars = [...font.decode([code])];
  if (chars.length !== 1) return; // a ligature, or no character at all
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

/** A face's gathered glyphs as {@link FaceOutlines}, where it is one to embed. */
function outlinesOf(face: Gathered, gap: number | undefined): FaceOutlines | undefined {
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
  return {
    glyphs,
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
