import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildXlsx } from './fixtures/build-xlsx';
import type { TextLineItem, TextToken } from '@/layout/page-doc';
import type { LayoutProfile } from '@/layout/styled-layout';
import { convertXlsxToPdfSync } from '@/core/converter';
import { Ream } from '@/core/converter/ream';
import { FontRegistry, createFontMeasure, parseTtf } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

// One long paragraph wraps into many single-spaced lines on the page, so the
// gap between consecutive baselines IS the line height — no inter-paragraph
// spacing in the way. (E-PARITY FP2.)
const docx = buildDocxFromBody(`<w:p><w:r><w:t>${'word '.repeat(200)}</w:t></w:r></w:p>`);

function lineCmds(profile?: LayoutProfile): ReadonlyArray<TextLineItem> {
  const flow = Ream.parse(docx).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
    ...(profile ? { layoutProfile: profile } : {}),
  });
  return laid.pages[0]!.commands.filter((c): c is TextLineItem => c.type === 'line');
}

// Line height = the gap between the 2nd and 3rd baselines (skip line 0, which
// also carries the block's leading-edge offset).
function lineHeightOf(profile?: LayoutProfile): number {
  const lines = lineCmds(profile);
  return Math.abs(lines[2]!.baselineY - lines[1]!.baselineY);
}

describe('layoutProfile — metric-derived leading (E-PARITY FP2)', () => {
  const parsed = parseTtf(FONTS.regular);
  const upem = parsed.unitsPerEm;
  const size = lineCmds()[1]!.line.maxFontSizePt; // body font size (same for all)

  it("default is the 'ream' profile — flat 1.2× leading, byte-identical", () => {
    const def = lineCmds().map((c) => c.baselineY);
    const ream = lineCmds('ream').map((c) => c.baselineY);
    expect(def).toEqual(ream);
    // Flat model: a single-spaced line is exactly 1.2× the font size.
    expect(lineHeightOf()).toBeCloseTo(size * 1.2, 4);
  });

  it("'word' derives leading from the font's OS/2 usWin metrics", () => {
    const vm = parsed.vmetrics;
    const expected = ((vm.winAscent + vm.winDescent) / upem) * size;
    expect(lineHeightOf('word')).toBeCloseTo(expected, 4);
    // Roboto's win box differs from the flat 1.2×, so the profile moves leading.
    expect(lineHeightOf('word')).not.toBeCloseTo(size * 1.2, 2);
  });

  it("'libreoffice' derives leading from the hhea triple", () => {
    const expected = ((parsed.ascender - parsed.descender + parsed.lineGap) / upem) * size;
    expect(lineHeightOf('libreoffice')).toBeCloseTo(expected, 4);
  });
});

describe('ttf-parser exposes vertical metrics (E-PARITY FP2)', () => {
  it('reads the hhea line gap and the OS/2 win/typo verticals', () => {
    const p = parseTtf(FONTS.regular);
    expect(p.lineGap).toBeGreaterThanOrEqual(0);
    const vm = p.vmetrics;
    expect(vm.winAscent).toBeGreaterThan(0);
    expect(vm.winDescent).toBeGreaterThan(0);
    expect(vm.typoAscent).toBeGreaterThan(0);
    expect(vm.typoDescent).toBeLessThan(0); // stored negative
    expect(typeof vm.useTypoMetrics).toBe('boolean');
  });
});

// A paragraph of varied-length words: Knuth-Plass rebalances it differently than
// greedy first-fit, so the two break it into different lines. (E-PARITY FP3.)
const varied = Array.from({ length: 160 }, (_, i) => 'x'.repeat(2 + ((i * 7) % 11))).join(' ');
// A narrow column makes each break decision count, so greedy first-fit and
// total-fit diverge (on a wide page they often coincide).
const variedDocx = buildDocxFromBody(
  `<w:p><w:r><w:t>${varied}</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="3200" w:h="20000"/>` +
    `<w:pgMar w:top="720" w:right="360" w:bottom="720" w:left="360"/></w:sectPr>`,
);

function variedLines(profile?: LayoutProfile): ReadonlyArray<TextLineItem> {
  const flow = Ream.parse(variedDocx).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
    ...(profile ? { layoutProfile: profile } : {}),
  });
  return laid.pages.flatMap((p) => p.commands.filter((c): c is TextLineItem => c.type === 'line'));
}

// Per-line text-token count across the whole paragraph — a fingerprint of where
// the lines break, independent of pagination.
function breakShape(profile?: LayoutProfile): ReadonlyArray<number> {
  return variedLines(profile).map((l) => l.line.tokens.filter((t) => t.kind === 'text').length);
}

// The words, in order, with breaking collapsed away — must be identical whatever
// the breaker does.
function wordsInOrder(profile?: LayoutProfile): string {
  return variedLines(profile)
    .flatMap((l) =>
      l.line.tokens
        .filter((t): t is TextToken => t.kind === 'text' && !t.isSpace)
        .map((t) => t.text),
    )
    .join('|');
}

describe('layoutProfile — greedy line breaking (E-PARITY FP3)', () => {
  it('breaks lines greedily under a profile, unlike default Knuth-Plass', () => {
    // Default stays total-fit…
    expect(breakShape()).toEqual(breakShape('ream'));
    // …and a profile diverges from it (greedy first-fit packs lines differently).
    expect(breakShape('word')).not.toEqual(breakShape('ream'));
  });

  it("'word' and 'libreoffice' share the greedy break structure (leading aside)", () => {
    // Both profiles break greedily; only their leading differs, and leading does
    // not move horizontal break points.
    expect(breakShape('word')).toEqual(breakShape('libreoffice'));
  });

  it('preserves every word and its order regardless of the breaker', () => {
    expect(wordsInOrder('word')).toBe(wordsInOrder('ream'));
    expect(wordsInOrder('libreoffice')).toBe(wordsInOrder('ream'));
  });
});

// A paragraph full of classic kern pairs (A-V, W-A, T-o, Y-a, V-o, …). (FP4.)
const kernRich = Array.from({ length: 40 }, () => 'WAVE AVATAR Tomato Yawn Active Vote Wax').join(
  ' ',
);
const kernDocx = buildDocxFromBody(`<w:p><w:r><w:t>${kernRich}</w:t></w:r></w:p>`);

// Total width of the word (non-space) tokens — independent of where lines break
// (each word is measured once), so it isolates the kerning difference.
function sumWordWidth(profile?: LayoutProfile): number {
  const flow = Ream.parse(kernDocx).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
    ...(profile ? { layoutProfile: profile } : {}),
  });
  let sum = 0;
  for (const page of laid.pages) {
    for (const c of page.commands.filter((x): x is TextLineItem => x.type === 'line')) {
      for (const t of c.line.tokens) if (t.kind === 'text' && !t.isSpace) sum += t.widthPt;
    }
  }
  return sum;
}

describe('layoutProfile — kerning (E-PARITY FP4)', () => {
  it("'word' measures kern-free; the default keeps kerning", () => {
    const p = parseTtf(FONTS.regular);
    const kerned = createFontMeasure(p, true).textWidthPt('AVATAR', 100);
    const unkerned = createFontMeasure(p, false).textWidthPt('AVATAR', 100);
    expect(unkerned).toBeGreaterThan(kerned); // dropping the negative pair kerns widens it
    expect(createFontMeasure(p).textWidthPt('AVATAR', 100)).toBe(kerned); // default = kerned
  });

  it("a kern-rich paragraph measures wider under 'word' than the kerned profiles", () => {
    // 'word' drops kerning, so its word widths are larger; 'ream' and
    // 'libreoffice' keep it and agree (breaking does not move word widths).
    expect(sumWordWidth('word')).toBeGreaterThan(sumWordWidth('libreoffice'));
    expect(sumWordWidth('libreoffice')).toBeCloseTo(sumWordWidth(), 6);
  });
});

describe('layoutProfile is a flowing-text knob — xlsx ignores it (E-PARITY FP0)', () => {
  it('xlsx output is byte-identical with and without a layoutProfile', () => {
    // Spreadsheet row heights follow the Calc print model, not font-metric
    // leading, so the profile must not touch xlsx geometry. (Corpus FP0 showed
    // a metric leading regresses Calc parity.)
    const xlsx = buildXlsx({
      rows: [
        ['Name', 'Score'],
        ['Ada', 42],
        ['Lin', 7],
      ],
    });
    const plain = convertXlsxToPdfSync(xlsx, { fonts: FONTS });
    const profiled = convertXlsxToPdfSync(xlsx, { fonts: FONTS, layoutProfile: 'libreoffice' });
    expect(profiled).toEqual(plain);
  });
});
