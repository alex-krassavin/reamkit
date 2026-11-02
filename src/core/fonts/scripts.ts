// Which writing system a character belongs to, and which of them a document
// needs a face for.
//
// The curated substitutes are Latin families. Greek, Cyrillic and Hebrew ride
// along in them; Han, Kana, Hangul, Arabic, Thai, the scripts of India and the
// geometric symbols do not, and 2145 characters of the corpus are exactly
// those — a notdef box every one. A face for them is fetched only when a
// document holds such a character (see `fetchScriptFont`), so the question
// this answers is: which ones.

import type { BodyElement, ShapeBlock } from '@/core/document-model';
import type { FlowDoc } from '@/core/ir/flow';
import type { ScriptKey } from '@/core/fonts/remote-fonts';

// The blocks each face covers. Ordered: the first range that contains the code
// point wins, and HAN is deliberately last because which of the three CJK faces
// draws it is a property of the DOCUMENT, not of the character (see below).
const RANGES: ReadonlyArray<readonly [number, number, ScriptKey | 'han']> = [
  [0x0590, 0x05ff, 'hebrew'],
  [0x0600, 0x06ff, 'arabic'],
  [0x0750, 0x077f, 'arabic'],
  [0x08a0, 0x08ff, 'arabic'],
  // The scripts of India, a block each. The two dandas, U+0964 and U+0965, are
  // in the Devanagari block and are the full stop of Bengali, Gurmukhi and
  // Odia too, so they name no face of their own: every one of these carries
  // them, and a Bengali document that ends a sentence with one is not sent
  // for Devanagari as well.
  [0x0900, 0x0963, 'devanagari'],
  [0x0966, 0x097f, 'devanagari'],
  [0x0980, 0x09ff, 'bengali'],
  [0x0a00, 0x0a7f, 'gurmukhi'],
  [0x0a80, 0x0aff, 'gujarati'],
  [0x0b00, 0x0b7f, 'oriya'],
  [0x0b80, 0x0bff, 'tamil'],
  [0x0c00, 0x0c7f, 'telugu'],
  [0x0c80, 0x0cff, 'kannada'],
  [0x0d00, 0x0d7f, 'malayalam'],
  [0x0e00, 0x0e7f, 'thai'],
  [0x1100, 0x11ff, 'kr'], // Hangul Jamo
  [0x1cd0, 0x1cff, 'devanagari'], // Vedic extensions
  // Both symbol faces are fetched whenever either is called for, so these say
  // only which one is TRIED first — the other answers for what it lacks.
  [0x2190, 0x21ff, 'symbols1'], // arrows
  [0x2300, 0x23ff, 'symbols1'], // technical
  [0x2460, 0x24ff, 'symbols1'], // enclosed alphanumerics
  [0x2500, 0x27bf, 'symbols2'], // box drawing, geometric shapes, dingbats
  [0x2b00, 0x2bff, 'symbols2'],
  [0x3000, 0x303f, 'han'], // CJK punctuation
  [0x3040, 0x30ff, 'jp'], // Hiragana + Katakana
  [0x3130, 0x318f, 'kr'], // Hangul compatibility Jamo
  [0x3400, 0x4dbf, 'han'],
  [0x4e00, 0x9fff, 'han'],
  [0xa8e0, 0xa8ff, 'devanagari'], // Devanagari extended
  [0xac00, 0xd7af, 'kr'], // Hangul syllables
  [0xf900, 0xfaff, 'han'],
  [0xfb1d, 0xfb4f, 'hebrew'],
  [0xfb50, 0xfdff, 'arabic'],
  [0xfe70, 0xfeff, 'arabic'],
  [0xff00, 0xffef, 'han'], // halfwidth + fullwidth forms
];

/**
 * The writing system a character needs a face for, or `undefined` when a Latin
 * family covers it.
 *
 * @param cp       The code point.
 * @param hanFace  Which face draws unified Han in this document (see
 *                 {@link scriptsInFlow}); defaults to Simplified.
 */
export function scriptForCodepoint(cp: number, hanFace: ScriptKey = 'sc'): ScriptKey | undefined {
  for (const [lo, hi, key] of RANGES) {
    if (cp < lo) break;
    if (cp <= hi) return key === 'han' ? hanFace : key;
  }
  return undefined;
}

/**
 * §17.3.2.26 — the slot of `w:rFonts` a character is set in: `ascii`,
 * `hAnsi` (high ANSI), `eastAsia` or `cs` (complex script).
 */
export type FontSlot = 'ascii' | 'hAnsi' | 'eastAsia' | 'cs';

/** The scripts Word sets in a run's complex-script font, `w:cs`. */
const COMPLEX_SCRIPTS: ReadonlySet<ScriptKey> = new Set<ScriptKey>([
  'hebrew',
  'arabic',
  'thai',
  'devanagari',
  'bengali',
  'gurmukhi',
  'gujarati',
  'oriya',
  'tamil',
  'telugu',
  'kannada',
  'malayalam',
]);

/**
 * §17.3.2.26 — the slot of a run's fonts Word sets a character in: ASCII in
 * `ascii`, East Asian text in `eastAsia`, complex script in `cs`, and the rest
 * above ASCII — Cyrillic, Greek, accented Latin, curly quotes — in `hAnsi`. A
 * `w:rtl` run is complex script whole (§17.3.2.30), its digits and Latin too.
 *
 * Measured in Word for Mac on a run whose ascii font is Verdana and whose
 * hAnsi one is Courier New: "split" reports Verdana, "Кириллица" Courier New.
 *
 * @param cp  The code point.
 * @param rtl Whether the run is right-to-left.
 * @returns The slot.
 */
export function fontSlotOf(cp: number, rtl = false): FontSlot {
  if (rtl) return 'cs';
  if (cp <= 0x7f) return 'ascii';
  const script = scriptForCodepoint(cp);
  if (script !== undefined && COMPLEX_SCRIPTS.has(script)) return 'cs';
  return script === 'kr' || script === 'jp' || script === 'sc' ? 'eastAsia' : 'hAnsi';
}

/**
 * The writing systems a document holds text in, beside the Latin one.
 *
 * Unified Han is written the same in Japanese, Korean and Chinese and drawn
 * differently in each, and the character does not say which — the document
 * does: Kana beside it means Japanese, Hangul means Korean. So the whole text
 * is read once, and the answer applies to every Han character in it.
 *
 * @param flow The parsed document.
 * @returns The scripts to fetch a face for, and the one Han is drawn with.
 */
export function scriptsInFlow(flow: FlowDoc): {
  scripts: Set<ScriptKey>;
  hanFace: ScriptKey;
} {
  const scripts = new Set<ScriptKey>();
  // Whether the document holds unified Han at all — kept in a set rather than a
  // flag so the answer survives the closure the walk runs in.
  const marks = new Set<'han'>();
  const read = (text: string): void => {
    for (const ch of text) {
      const cp = ch.codePointAt(0) ?? 0;
      if (cp < 0x0590) continue;
      // 'sc' here means "unified Han", the placeholder the default resolves
      // to — which face draws it is decided once, below.
      const key = scriptForCodepoint(cp, 'sc');
      if (key === undefined) continue;
      if (key === 'sc') {
        marks.add('han');
        continue;
      }
      // One symbol calls for both symbol faces. Neither covers the other's
      // half — ZapfDingbats.pdf prints the whole repertory, and read with the
      // second face alone its crosses and its circled numbers came out boxes —
      // and which of the two holds a given character is not something the
      // ranges above can say block by block.
      if (key === 'symbols1' || key === 'symbols2') {
        scripts.add('symbols1');
        scripts.add('symbols2');
        continue;
      }
      scripts.add(key);
    }
  };

  const shapeText = (shape: ShapeBlock): void => {
    if (shape.text) visit(shape.text.content);
    for (const child of shape.children ?? []) shapeText(child.shape);
  };
  const visit = (elements: ReadonlyArray<BodyElement>): void => {
    for (const el of elements) {
      if (el.kind === 'paragraph') {
        for (const run of el.paragraph.runs) read(run.text);
      } else if (el.kind === 'table') {
        for (const row of el.table.rows) for (const cell of row.cells) visit(cell.content);
        // …and the drawings over a spreadsheet grid (TableOverlay).
        visit(el.table.overlay?.drawings ?? []);
      } else if (el.kind === 'shape') {
        shapeText(el.shape);
      }
    }
  };
  visit(flow.body);
  for (const band of flow.headersFooters?.values() ?? []) visit(band);
  for (const note of flow.footnotes?.values() ?? []) visit(note);
  for (const note of flow.endnotes?.values() ?? []) visit(note);
  for (const comment of flow.comments?.values() ?? []) visit(comment.content);

  // Kana wins over Hangul: a Japanese document quotes Korean far less often
  // than the other way round, and either face draws the Han it shares.
  const hanFace: ScriptKey = scripts.has('jp') ? 'jp' : scripts.has('kr') ? 'kr' : 'sc';
  if (marks.has('han')) scripts.add(hanFace);
  return { scripts, hanFace };
}
