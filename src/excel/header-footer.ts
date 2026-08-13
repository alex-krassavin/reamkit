// Sheet header/footer expansion (E-SHEET W4). Excel stores a print header/footer
// as a single string in a `&`-code mini-language: `&L`/`&C`/`&R` switch the
// left / centre / right region, field codes (`&P` page, `&N` total, `&A` sheet
// name, …) inject values, and `&B`/`&I` toggle bold/italic. We expand one string
// into header/footer band content, with `&P`/`&N` as dynamic PAGE/NUMPAGES field
// runs the renderer resolves per page. Regions that stand together stand on ONE
// line, separated by tab stops at the middle and the far edge of the band — the
// way Excel draws them and the way Word writes the same header; a single region
// is a line of its own, aligned. `&B`/`&I`/`&U`/`&S` toggle bold/italic/underline/
// strike, `&nn` sets the point size, `&Krrggbb` the colour, and the style
// suffix of `&"family,style"` acts as another bold/italic toggle (the family
// itself is dropped — the renderer has one font set). `&F` and `&D`/`&T` are
// the caller's to supply (a file name, a reference date) and are dropped
// without one; `&Z` paths and `&G` pictures are dropped outright.

import type { Alignment, BodyElement, Run, RunProperties, TabStop } from '@/core/document-model';
import { pt } from '@/core/ir';

// §18.3.1.55 — the theme-colour order a header/footer's `&K` reference indexes,
// which is the workbook's own (§18.8.3): background first, then text.
const HEADER_THEME_SLOTS: ReadonlyArray<string> = [
  'lt1',
  'dk1',
  'lt2',
  'dk2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
];

/**
 * Lighten or darken a theme colour by a header/footer tint (§18.3.1.55): the
 * value is a fraction of the way to white, or to black when the sign is minus.
 *
 * @param hex     The theme colour, RRGGBB.
 * @param amount  0..1.
 * @param darken  True for a `-` sign.
 * @returns The tinted colour, RRGGBB.
 */
function applyHeaderTint(hex: string, amount: number, darken: boolean): string {
  if (!(amount > 0)) return hex.toUpperCase();
  const n = parseInt(hex, 16);
  const mix = (c: number): number =>
    darken ? Math.round(c * (1 - amount)) : Math.round(c + (255 - c) * amount);
  const r = mix((n >> 16) & 255);
  const g = mix((n >> 8) & 255);
  const b = mix(n & 255);
  return [r, g, b]
    .map((c) => c.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

interface Regions {
  readonly left: Array<Run>;
  readonly center: Array<Run>;
  readonly right: Array<Run>;
}

/**
 * Expand a print header/footer `&`-code string (E-SHEET W4) into header/footer band
 * content — one aligned paragraph per non-empty left/centre/right region — or `[]`
 * when every region is empty. `&P`/`&N` become dynamic PAGE/NUMPAGES field runs the
 * renderer resolves per page; non-deterministic or unsupported codes are dropped.
 *
 * @param formatString The raw `&`-code format string.
 * @param sheetName    The worksheet tab name, substituted for `&A`.
 * @param scale        The sheet's print scale.
 * @param basePt       The default header font size.
 * @param fileName     The workbook's file name, substituted for `&F`. A
 *                     byte-oriented API does not know it, so the caller supplies
 *                     it; absent, `&F` is dropped as before.
 * @param themePalette The workbook theme, resolving a `&K` theme reference.
 * @param now          The reference date `&D`/`&T` print. An explicit input,
 *                     never the wall clock; absent, both are dropped as before.
 */
export function buildHeaderFooterContent(
  formatString: string,
  sheetName: string,
  scale = 1,
  basePt = DEFAULT_HEADER_PT,
  fileName?: string,
  themePalette?: ReadonlyMap<string, string>,
  now?: Date,
): Array<BodyElement> {
  const regions = parseHeaderFooterString(
    formatString,
    sheetName,
    scale,
    basePt,
    fileName,
    themePalette,
    now,
  );
  const out: Array<BodyElement> = [];
  const lines = {
    left: splitLines(regions.left),
    center: splitLines(regions.center),
    right: splitLines(regions.right),
  };
  const filled = (region: keyof typeof lines): number =>
    lines[region].filter((line) => line.length > 0).length;
  // One region is a line of its own, aligned. Two or three share ONE line, the
  // way Excel and every reader draw them: the left flush, the centre on the
  // middle of the band and the right against its far edge (§17.3.1.38 tab
  // stops, which is how Word writes the same header). Stacked as a paragraph
  // apiece, 45540_classic_Header.xlsx came back three lines deep where the
  // sheet has one.
  const alone =
    (filled('left') > 0 ? 1 : 0) + (filled('center') > 0 ? 1 : 0) + (filled('right') > 0 ? 1 : 0) <=
    1;
  if (alone) {
    const para = (runs: ReadonlyArray<Run>, alignment: Alignment): void => {
      for (const line of splitLines(runs)) {
        if (line.length > 0) {
          out.push({ kind: 'paragraph', paragraph: { properties: { alignment }, runs: line } });
        }
      }
    };
    para(regions.left, 'left');
    para(regions.center, 'center');
    para(regions.right, 'right');
    return out;
  }
  // A region may itself be several lines, and the k-th line of each stands with
  // the k-th line of the others.
  const deep = Math.max(lines.left.length, lines.center.length, lines.right.length);
  for (let i = 0; i < deep; i++) {
    const left = lines.left[i] ?? [];
    const center = lines.center[i] ?? [];
    const right = lines.right[i] ?? [];
    if (left.length + center.length + right.length === 0) continue;
    const runs: Array<Run> = [...left];
    if (center.length > 0) runs.push(TAB_RUN, ...center);
    if (right.length > 0) runs.push(TAB_RUN, ...right);
    out.push({ kind: 'paragraph', paragraph: { properties: { tabs: HEADER_STOPS }, runs } });
  }
  return out;
}

/** The tab a header's regions are separated by. */
const TAB_RUN: Run = { text: '\t', properties: {} };

/**
 * §17.3.1.38 — the two stops a header's regions stand on: the middle of the
 * band and its far edge, both stated against the band itself rather than at a
 * distance, since only the layout knows how wide the band is.
 */
const HEADER_STOPS: ReadonlyArray<TabStop> = [
  { positionPt: pt(0), relativeTo: 'center', alignment: 'center' },
  { positionPt: pt(0), relativeTo: 'right', alignment: 'right' },
];

/**
 * Break a region's runs at the line breaks Excel allows inside one — a header
 * region is not necessarily one line.
 *
 * The break is a literal CR/LF in the format string, not a `&`-code, so it
 * arrives as text. Left as text it is drawn: tdf58243.xlsx puts a CR LF in the
 * middle of its centre header and we rendered the carriage return as a missing
 * glyph — a tofu box mid-title — with the rest of the title running on after it
 * where every other reader starts a second line.
 */
function splitLines(runs: ReadonlyArray<Run>): Array<Array<Run>> {
  const lines: Array<Array<Run>> = [[]];
  for (const run of runs) {
    const parts = run.text.split(/\r\n|\r|\n/);
    for (let i = 0; i < parts.length; i++) {
      if (i > 0) lines.push([]);
      const text = parts[i]!;
      if (text.length > 0) lines[lines.length - 1]!.push({ ...run, text });
    }
  }
  return lines;
}

// Single-pass scan of the &-code string. The default region (before any &L/&C/&R)
// is the centre, matching Excel.
/** The point size a header run takes when it names none — the layout's own. */
const DEFAULT_HEADER_PT = 11;

function parseHeaderFooterString(
  s: string,
  sheetName: string,
  scale: number,
  basePt: number,
  fileName?: string,
  themePalette?: ReadonlyMap<string, string>,
  now?: Date,
): Regions {
  const regions: Regions = { left: [], center: [], right: [] };
  let current: Array<Run> = regions.center;
  let bold = false;
  let italic = false;
  let underline = false;
  let strike = false;
  let sizePt: number | undefined;
  let colorHex: string | undefined;
  let buf = '';

  const runProps = (): RunProperties => {
    const effectivePt = (sizePt ?? basePt) * scale;
    return {
      ...(bold ? { bold: true } : {}),
      ...(italic ? { italic: true } : {}),
      ...(underline ? { underline: 'single' as const } : {}),
      ...(strike ? { strike: true } : {}),
      // The print scale shrinks everything printed on the page, headers included:
      // a header left at its literal size over a grid at half scale is twice as
      // tall as it should be — and a header band taller than the gap above the
      // top margin pushes the body down, which costs rows and, on tdf58243.xlsx,
      // a whole third page. A run that names no size takes the sheet's own.
      ...(effectivePt !== DEFAULT_HEADER_PT ? { fontSizePt: pt(effectivePt) } : {}),
      ...(colorHex !== undefined ? { colorHex } : {}),
    };
  };
  const flush = (): void => {
    if (buf.length > 0) {
      current.push({ text: buf, properties: runProps() });
      buf = '';
    }
  };
  const pushField = (field: 'PAGE' | 'NUMPAGES'): void => {
    flush();
    current.push({ text: '1', properties: runProps(), field });
  };

  let i = 0;
  while (i < s.length) {
    const ch = s[i]!;
    if (ch !== '&') {
      buf += ch;
      i++;
      continue;
    }
    const next = s[i + 1];
    if (next === undefined) {
      buf += '&'; // a trailing, code-less ampersand
      break;
    }
    if (next === '&') {
      buf += '&'; // escaped literal ampersand
      i += 2;
      continue;
    }
    if (next === 'L' || next === 'C' || next === 'R') {
      flush();
      current = next === 'L' ? regions.left : next === 'R' ? regions.right : regions.center;
      // Each region starts from the sheet's own header font. Carrying the state
      // across the switch puts the left region's 16pt bold onto the page number
      // in the right one, which is not what the file means or what LibreOffice
      // prints — tdf171828's footer sets Bold Italic 12pt for its centre and
      // then writes a plain "Seite &P" on the right.
      bold = false;
      italic = false;
      underline = false;
      strike = false;
      sizePt = undefined;
      colorHex = undefined;
      i += 2;
      continue;
    }
    if (next === 'P') {
      pushField('PAGE');
      i += 2;
      continue;
    }
    if (next === 'N') {
      pushField('NUMPAGES');
      i += 2;
      continue;
    }
    if (next === 'A') {
      buf += sheetName; // the worksheet tab name
      i += 2;
      continue;
    }
    if (next === 'B') {
      flush();
      bold = !bold;
      i += 2;
      continue;
    }
    if (next === 'I') {
      flush();
      italic = !italic;
      i += 2;
      continue;
    }
    if (next === 'U') {
      flush();
      underline = !underline;
      i += 2;
      continue;
    }
    if (next === 'S') {
      flush();
      strike = !strike;
      i += 2;
      continue;
    }
    if (next === 'K') {
      // §18.3.1.55 — `&K` takes SIX hex digits, or a theme reference written as
      // two digits of theme-colour index, a sign and three of tint: `&K01+000`
      // is the theme's first colour at no tint. Read as hex, that reference
      // consumed the "01" and left "+000" on the page as text — and the colour
      // it was meant to reset never reset, so HeaderFooterComplexFormats.xlsx
      // ran red from its "RedUnderlined" to the end of the line.
      i += 2;
      const themed = /^(\d\d)([+-])(\d\d\d)/.exec(s.slice(i));
      if (themed) {
        i += themed[0].length;
        const slot = HEADER_THEME_SLOTS[Number(themed[1])];
        const base = slot ? themePalette?.get(slot) : undefined;
        flush();
        colorHex = base
          ? applyHeaderTint(base, Number(themed[3]) / 1000, themed[2] === '-')
          : undefined;
        continue;
      }
      let hex = '';
      while (hex.length < 6 && i < s.length && /[0-9A-Fa-f]/.test(s[i]!)) {
        hex += s[i]!;
        i++;
      }
      if (hex.length === 6) {
        flush();
        colorHex = hex.toUpperCase();
      }
      continue;
    }
    if (next === '"') {
      // &"font,style" — the family is dropped (the renderer has one font set),
      // but the style suffix is a bold/italic toggle like &B and &I are, and
      // dropping it silently unstyles a whole header region.
      i += 2;
      let spec = '';
      while (i < s.length && s[i] !== '"') {
        spec += s[i]!;
        i++;
      }
      if (i < s.length) i++; // closing quote
      const style = (spec.split(',')[1] ?? '').toLowerCase();
      if (style.length > 0) {
        flush();
        bold = style.includes('bold');
        italic = style.includes('italic') || style.includes('oblique');
      }
      continue;
    }
    if (next >= '0' && next <= '9') {
      // &nn — the point size for what follows. A header written at &16 and
      // rendered at the body size is the wrong size on every page.
      i += 1;
      let digits = '';
      while (i < s.length && s[i]! >= '0' && s[i]! <= '9') {
        digits += s[i]!;
        i++;
      }
      const n = Number(digits);
      if (Number.isFinite(n) && n > 0 && n <= 409) {
        flush();
        sizePt = n;
      }
      continue;
    }
    // §18.3.1.34 `&F` — the workbook's file name. A byte-oriented API does not
    // know it, so it arrives from the caller; without one the code drops as
    // before. Five of the first forty POI workbooks head every page with it.
    if (next === 'F') {
      if (fileName !== undefined && fileName.length > 0) buf += fileName;
      i += 2;
      continue;
    }
    // §18.3.1.35 `&D` / §18.3.1.48 `&T` — the date and the time the page is
    // printed. The wall clock is not ours to read, so these resolve against the
    // caller's reference date and are dropped without one, exactly as `&F` is.
    // customIndexedColors.xlsx heads every page with `&D - - &T`.
    if ((next === 'D' || next === 'T') && now) {
      buf += next === 'D' ? shortDate(now) : clockTime(now);
      i += 2;
      continue;
    }
    // Any other single-letter code (&Z &G &E &X &Y &O &H …) is dropped as
    // unsupported styling.
    i += 2;
  }
  flush();
  return regions;
}

// §18.3.1.35 — `&D` prints the date in the system's short form, which for a
// library has to be one fixed spelling; this is the one Excel and Calc write
// under en-US (and the one the reference render prints). Read in UTC so the
// same reference date gives the same header on every host.
function shortDate(d: Date): string {
  const p = (n: number): string => String(n).padStart(2, '0');
  return `${p(d.getUTCMonth() + 1)}/${p(d.getUTCDate())}/${String(d.getUTCFullYear())}`;
}

// §18.3.1.48 — `&T`, the clock time, likewise in UTC and to the second.
function clockTime(d: Date): string {
  const p = (n: number): string => String(n).padStart(2, '0');
  return `${p(d.getUTCHours())}:${p(d.getUTCMinutes())}:${p(d.getUTCSeconds())}`;
}
