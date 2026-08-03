// ECMA-376 Part 1 §17.3.2 — Run Properties (rPr).

import type {
  FontFamilyMap,
  RunProperties,
  UnderlineStyle,
  VerticalAlign,
} from '@/core/document-model';
import { halfPtToPt, twipsToPt } from '@/core/ir';

import { asElement, getAttr, getVal, parseHalfPointAttr, parseToggle } from '@/word/xml-helpers';
import { shadingFillHex } from '@/word/shading';

const UNDERLINE_STYLES = new Set<UnderlineStyle>([
  'none',
  'single',
  'double',
  'thick',
  'dotted',
  'dottedHeavy',
  'dash',
  'dashHeavy',
  'wave',
]);

const VERTICAL_ALIGNS = new Set<VerticalAlign>(['baseline', 'superscript', 'subscript']);

// §17.18.40 ST_HighlightColor — the marker pen's seventeen colours. `none` is
// the absence of one, so it maps to nothing rather than to a colour.
const HIGHLIGHT_COLORS: ReadonlyMap<string, string> = new Map([
  ['black', '000000'],
  ['blue', '0000FF'],
  ['cyan', '00FFFF'],
  ['darkBlue', '000080'],
  ['darkCyan', '008080'],
  ['darkGray', '808080'],
  ['darkGreen', '008000'],
  ['darkMagenta', '800080'],
  ['darkRed', '800000'],
  ['darkYellow', '808000'],
  ['green', '00FF00'],
  ['lightGray', 'C0C0C0'],
  ['magenta', 'FF00FF'],
  ['red', 'FF0000'],
  ['white', 'FFFFFF'],
  ['yellow', 'FFFF00'],
]);

/**
 * Parse a run-properties element (`w:rPr`, ECMA-376 Part 1 §17.3.2) into the
 * typed {@link RunProperties}. Reads the character formatting the renderer needs:
 * style reference, bold/italic/strike toggles, underline style, size, colour,
 * font family map, vertical alignment, RTL, and language; sizes convert from
 * half-points to points and colours normalize to uppercase hex.
 *
 * @param rPr The flat-tree `w:rPr` element (or `undefined`/non-element).
 * @returns The parsed properties; an empty object when `rPr` is absent.
 */
export function parseRunProperties(rPr: unknown): RunProperties {
  const el = asElement(rPr);
  if (!el) return {};

  const out: Mutable<RunProperties> = {};

  if ('w:rStyle' in el) {
    const v = getVal(el['w:rStyle']);
    if (v) out.styleId = v;
  }

  if ('w:b' in el) {
    const v = parseToggle(el['w:b']);
    if (v !== undefined) out.bold = v;
  }
  if ('w:i' in el) {
    const v = parseToggle(el['w:i']);
    if (v !== undefined) out.italic = v;
  }
  if ('w:strike' in el) {
    const v = parseToggle(el['w:strike']);
    if (v !== undefined) out.strike = v;
  }

  // §17.3.2.5 / §17.3.2.33 — the run is DISPLAYED in capitals, whatever it
  // stores. Read nowhere, capitalized.docx printed its word in lower case
  // where every other reader shouts it.
  if ('w:caps' in el) {
    const v = parseToggle(el['w:caps']);
    if (v !== undefined) out.caps = v;
  }
  if ('w:smallCaps' in el) {
    const v = parseToggle(el['w:smallCaps']);
    if (v !== undefined) out.smallCaps = v;
  }

  if ('w:u' in el) {
    const v = getVal(el['w:u']);
    if (v && UNDERLINE_STYLES.has(v as UnderlineStyle)) {
      out.underline = v as UnderlineStyle;
    }
    // §17.3.2.40 — the rule under the text has a colour of its own. A themed
    // one is written alongside the RESOLVED hex it stands for (`w:color`), so
    // reading that is enough: Test_CharUnderlineThemeColor.docx asks for a gold
    // rule under black text and we drew it black.
    const c = getAttr(el['w:u'], 'color');
    if (c && /^[0-9A-Fa-f]{6}$/u.test(c)) out.underlineColorHex = c.toUpperCase();
  }

  if ('w:sz' in el) {
    const v = parseHalfPointAttr(el['w:sz'], 'val');
    if (v !== undefined) out.fontSizePt = halfPtToPt(v);
  }

  if ('w:color' in el) {
    const v = getVal(el['w:color']);
    if (v && /^[0-9A-Fa-f]{6}$/.test(v)) {
      out.colorHex = v.toUpperCase();
    } else if (v === 'auto') {
      // §17.3.2.6 — `auto` is a COLOUR, the automatic one, and a run that names
      // it overrides whatever its style lends: fdo77887.docx writes its Normal
      // style in blue and every run back to auto, and we printed the whole
      // form blue where the reference prints it black.
      out.colorHex = '000000';
    }
  }

  if ('w:rFonts' in el) {
    const ff = parseFontFamily(el['w:rFonts']);
    if (ff) out.fontFamily = ff;
  }

  if ('w:vertAlign' in el) {
    const v = getVal(el['w:vertAlign']);
    if (v && VERTICAL_ALIGNS.has(v as VerticalAlign)) {
      out.verticalAlign = v as VerticalAlign;
    }
  }

  // ECMA-376 §17.3.2.30 — w:rtl is a toggle property.
  if ('w:rtl' in el) {
    const v = parseToggle(el['w:rtl']);
    if (v !== undefined) out.rtl = v;
  }

  // §17.3.2.35 — the space added between characters, in twentieths of a point.
  // The paragraph's `w:spacing` is a different element with different
  // attributes; this one carries `w:val`. fdo71302.docx tracks out its
  // Subtitle style that way and we set it solid.
  if ('w:spacing' in el) {
    const v = getVal(el['w:spacing']);
    const n = v === undefined ? Number.NaN : Number(v);
    if (Number.isFinite(n) && n !== 0) out.letterSpacingPt = twipsToPt(n);
  }

  // §17.3.2.32 — the run's own background. Unlike a paragraph's, this one is
  // usually a PATTERN over a fill (`pct15` of black on white is Word's "light
  // shading"), so the two are blended rather than the fill taken alone:
  // fdo65400.docx shades two words that way and we painted neither.
  if ('w:shd' in el) {
    const shd = el['w:shd'];
    const hex = shadingFillHex(getVal(shd), getAttr(shd, 'color'), getAttr(shd, 'fill'));
    if (hex !== undefined && hex !== 'FFFFFF') out.shadingColorHex = hex;
  }

  // §17.3.2.15 `w:highlight` — the marker pen: one of seventeen named colours
  // painted behind the run. It goes in the same slot as the run's shading —
  // both are a filled box under the glyphs, and where a run states both, the
  // highlight is the one that shows. fdo76591.docx sets its "IMPORTANT NOTICE"
  // in white on a black highlight, and unread it was white on white.
  if ('w:highlight' in el) {
    const hex = HIGHLIGHT_COLORS.get(getVal(el['w:highlight']) ?? '');
    if (hex !== undefined) out.shadingColorHex = hex;
  }

  // ECMA-376 §17.3.2.20 — w:lang @w:val (the Latin language, e.g. "en-US").
  // Surfaced for the tagged-PDF per-element /Lang (a paragraph whose language
  // differs from the document default is tagged so AT switches pronunciation).
  if ('w:lang' in el) {
    const v = getAttr(el['w:lang'], 'val');
    if (v) out.lang = v;
  }

  return out;
}

function parseFontFamily(node: unknown): FontFamilyMap | undefined {
  const ff: Mutable<FontFamilyMap> = {};
  const ascii = getAttr(node, 'ascii');
  const hAnsi = getAttr(node, 'hAnsi');
  const cs = getAttr(node, 'cs');
  const eastAsia = getAttr(node, 'eastAsia');
  if (ascii) ff.ascii = ascii;
  if (hAnsi) ff.hAnsi = hAnsi;
  if (cs) ff.cs = cs;
  if (eastAsia) ff.eastAsia = eastAsia;
  return Object.keys(ff).length > 0 ? ff : undefined;
}

type Mutable<T> = { -readonly [K in keyof T]: T[K] };
