// Sheet header/footer expansion (E-SHEET W4). Excel stores a print header/footer
// as a single string in a `&`-code mini-language: `&L`/`&C`/`&R` switch the
// left / centre / right region, field codes (`&P` page, `&N` total, `&A` sheet
// name, …) inject values, and `&B`/`&I` toggle bold/italic. We expand one string
// into one aligned paragraph per non-empty region, with `&P`/`&N` as dynamic
// PAGE/NUMPAGES field runs the renderer resolves per page. The header/footer band
// layout draws paragraphs, so each region is its own paragraph (a left+right
// header therefore stacks rather than sharing one line — the common single-region
// case stays on one line). Non-deterministic or unsupported codes (&D date,
// &T time, &F file, &Z path, &G picture, font/size/colour/underline) are dropped.

import type { Alignment, BodyElement, Run } from '@/core/document-model';

interface Regions {
  readonly left: Array<Run>;
  readonly center: Array<Run>;
  readonly right: Array<Run>;
}

// Expand a header/footer format string into header/footer band content — one
// aligned paragraph per non-empty region — or [] when every region is empty.
export function buildHeaderFooterContent(
  formatString: string,
  sheetName: string,
): Array<BodyElement> {
  const regions = parseHeaderFooterString(formatString, sheetName);
  const out: Array<BodyElement> = [];
  const para = (runs: ReadonlyArray<Run>, alignment: Alignment): void => {
    if (runs.length > 0) {
      out.push({ kind: 'paragraph', paragraph: { properties: { alignment }, runs: [...runs] } });
    }
  };
  para(regions.left, 'left');
  para(regions.center, 'center');
  para(regions.right, 'right');
  return out;
}

// Single-pass scan of the &-code string. The default region (before any &L/&C/&R)
// is the centre, matching Excel.
function parseHeaderFooterString(s: string, sheetName: string): Regions {
  const regions: Regions = { left: [], center: [], right: [] };
  let current: Array<Run> = regions.center;
  let bold = false;
  let italic = false;
  let buf = '';

  const runProps = () => ({ ...(bold ? { bold: true } : {}), ...(italic ? { italic: true } : {}) });
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
    if (next === 'K') {
      // &Krrggbb (or a theme-colour spec) — drop the colour, skip its hex digits.
      i += 2;
      let n = 0;
      while (n < 6 && i < s.length && /[0-9A-Fa-f]/.test(s[i]!)) {
        i++;
        n++;
      }
      continue;
    }
    if (next === '"') {
      // &"font,style" — drop the font selection.
      i += 2;
      while (i < s.length && s[i] !== '"') i++;
      if (i < s.length) i++; // closing quote
      continue;
    }
    if (next >= '0' && next <= '9') {
      // &nn font size — drop it.
      i += 1;
      while (i < s.length && s[i]! >= '0' && s[i]! <= '9') i++;
      continue;
    }
    // Any other single-letter code (&D &T &F &Z &G &U &E &S &X &Y &O &H …) is
    // dropped: non-deterministic (date/time/file) or unsupported styling.
    i += 2;
  }
  flush();
  return regions;
}
