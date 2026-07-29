// Sheet header/footer expansion (E-SHEET W4). Excel stores a print header/footer
// as a single string in a `&`-code mini-language: `&L`/`&C`/`&R` switch the
// left / centre / right region, field codes (`&P` page, `&N` total, `&A` sheet
// name, …) inject values, and `&B`/`&I` toggle bold/italic. We expand one string
// into one aligned paragraph per non-empty region, with `&P`/`&N` as dynamic
// PAGE/NUMPAGES field runs the renderer resolves per page. The header/footer band
// layout draws paragraphs, so each region is its own paragraph (a left+right
// header therefore stacks rather than sharing one line — the common single-region
// case stays on one line). `&B`/`&I`/`&U`/`&S` toggle bold/italic/underline/
// strike, `&nn` sets the point size, `&Krrggbb` the colour, and the style
// suffix of `&"family,style"` acts as another bold/italic toggle (the family
// itself is dropped — the renderer has one font set). Non-deterministic codes
// (&D date, &T time, &F file, &Z path) and &G pictures are dropped.

import type { Alignment, BodyElement, Run, RunProperties } from '@/core/document-model';
import { pt } from '@/core/ir';

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
 */
export function buildHeaderFooterContent(
  formatString: string,
  sheetName: string,
  scale = 1,
  basePt = DEFAULT_HEADER_PT,
): Array<BodyElement> {
  const regions = parseHeaderFooterString(formatString, sheetName, scale, basePt);
  const out: Array<BodyElement> = [];
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
      // &Krrggbb — six hex digits, or a theme spec we cannot resolve here.
      i += 2;
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
    // Any other single-letter code (&D &T &F &Z &G &E &X &Y &O &H …) is dropped:
    // non-deterministic (date/time/file) or unsupported styling.
    i += 2;
  }
  flush();
  return regions;
}
