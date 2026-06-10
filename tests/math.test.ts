import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { MathNode } from '@/core/document-model';
import { convertDocxToPdfSync } from '@/core/converter';
import { OpcPackage } from '@/core/opc';
import { parseDocument } from '@/word';
import { layoutMath } from '@/pdf/math-layout';

const here = dirname(fileURLToPath(import.meta.url));
const FONTS = {
  regular: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf'))),
};
const latin1 = new TextDecoder('latin1');
const asLatin1 = (b: Uint8Array): string => latin1.decode(b);

const M_NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';

// a/b fraction wrapped in an oMath.
const FRACTION = `<m:oMath ${M_NS}>
  <m:f>
    <m:num><m:r><m:t>a</m:t></m:r></m:num>
    <m:den><m:r><m:t>b</m:t></m:r></m:den>
  </m:f>
</m:oMath>`;

const measure = (t: string, sz: number): number => [...t].length * sz * 0.5;

describe('OMML parser', () => {
  it('parses an inline fraction into a math run', () => {
    const docx = buildDocxFromBody(`<w:p>${FRACTION}</w:p>`);
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    expect(parsed[0]!.kind).toBe('paragraph');
    if (parsed[0]!.kind !== 'paragraph') throw new Error('unreachable');
    const run = parsed[0]!.paragraph.runs[0]!;
    expect(run.math).toMatchObject({
      type: 'row',
      children: [
        {
          type: 'fraction',
          num: { type: 'row', children: [{ type: 'run', text: 'a' }] },
          den: { type: 'row', children: [{ type: 'run', text: 'b' }] },
        },
      ],
    });
  });

  it('reads run style from m:rPr (m:sty / m:nor)', () => {
    const docx = buildDocxFromBody(
      `<w:p><m:oMath ${M_NS}><m:r><m:rPr><m:nor/></m:rPr><m:t>sin</m:t></m:r></m:oMath></w:p>`,
    );
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    if (parsed[0]!.kind !== 'paragraph') throw new Error('unreachable');
    expect(parsed[0]!.paragraph.runs[0]!.math).toMatchObject({
      type: 'row',
      children: [{ type: 'run', text: 'sin', nor: true }],
    });
  });
});

describe('math layout engine', () => {
  it('auto-italicises letters but leaves digits upright', () => {
    const run: MathNode = { type: 'run', text: 'x2' };
    const box = layoutMath(run, { sizePt: 12 }, measure);
    const glyphs = box.items.filter((it) => it.kind === 'glyph');
    expect(glyphs.map((g) => g.variant)).toEqual(['italic', 'regular']);
    expect(glyphs.map((g) => g.text)).toEqual(['x', '2']);
  });

  it('stacks a fraction with a rule and straddles the baseline', () => {
    const frac: MathNode = {
      type: 'fraction',
      num: { type: 'row', children: [{ type: 'run', text: 'a' }] },
      den: { type: 'row', children: [{ type: 'run', text: 'bc' }] },
    };
    const box = layoutMath(frac, { sizePt: 12 }, measure);
    expect(box.width).toBeGreaterThan(0);
    expect(box.ascent).toBeGreaterThan(0); // numerator above baseline
    expect(box.descent).toBeGreaterThan(0); // denominator below baseline
    expect(box.items.some((it) => it.kind === 'rule')).toBe(true); // the bar
    const glyphText = box.items
      .filter((it) => it.kind === 'glyph')
      .map((it) => it.text)
      .join('');
    expect(glyphText).toContain('a');
    expect(glyphText).toContain('bc');
  });

  it('omits the rule for a barless (noBar) fraction', () => {
    const frac: MathNode = {
      type: 'fraction',
      barless: true,
      num: { type: 'row', children: [{ type: 'run', text: 'a' }] },
      den: { type: 'row', children: [{ type: 'run', text: 'b' }] },
    };
    const box = layoutMath(frac, { sizePt: 12 }, measure);
    expect(box.items.some((it) => it.kind === 'rule')).toBe(false);
  });
});

const oMath = (inner: string): string => `<m:oMath ${M_NS}>${inner}</m:oMath>`;
const SUP = `<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>`;
const RAD = `<m:rad><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>`;

describe('scripts', () => {
  it('parses a superscript', () => {
    const docx = buildDocxFromBody(`<w:p>${oMath(SUP)}</w:p>`);
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    if (parsed[0]!.kind !== 'paragraph') throw new Error('unreachable');
    expect(parsed[0]!.paragraph.runs[0]!.math).toMatchObject({
      type: 'row',
      children: [
        {
          type: 'script',
          base: { type: 'row', children: [{ type: 'run', text: 'x' }] },
          sup: { type: 'row', children: [{ type: 'run', text: '2' }] },
        },
      ],
    });
  });

  it('lays out a superscript smaller and raised', () => {
    const node: MathNode = {
      type: 'script',
      base: { type: 'row', children: [{ type: 'run', text: 'x' }] },
      sup: { type: 'row', children: [{ type: 'run', text: '2' }] },
    };
    const box = layoutMath(node, { sizePt: 12 }, measure);
    const glyphs = box.items.filter((it) => it.kind === 'glyph');
    const sizes = glyphs.map((g) => g.sizePt);
    expect(sizes).toContain(12); // base
    expect(Math.min(...sizes)).toBeLessThan(12); // smaller script
    const script = box.items.find((it) => it.kind === 'glyph' && it.sizePt < 12);
    expect(script && script.kind === 'glyph' ? script.y : 0).toBeGreaterThan(0); // raised
    const baseBox = layoutMath({ type: 'run', text: 'x' }, { sizePt: 12 }, measure);
    expect(box.ascent).toBeGreaterThan(baseBox.ascent);
  });
});

describe('radicals', () => {
  it('lays out a drawn surd path over the radicand', () => {
    const node: MathNode = {
      type: 'radical',
      radicand: { type: 'row', children: [{ type: 'run', text: 'x' }] },
    };
    const box = layoutMath(node, { sizePt: 12 }, measure);
    expect(box.items.some((it) => it.kind === 'path')).toBe(true);
    expect(box.items.some((it) => it.kind === 'glyph')).toBe(true);
    expect(box.ascent).toBeGreaterThan(0);
  });
});

const EQARR = `<m:eqArr>
  <m:e><m:r><m:t>a=1</m:t></m:r></m:e>
  <m:e><m:r><m:t>b=2</m:t></m:r></m:e>
</m:eqArr>`;

describe('equation arrays (m:eqArr)', () => {
  it('parses an equation array into stacked rows', () => {
    const docx = buildDocxFromBody(`<w:p>${oMath(EQARR)}</w:p>`);
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    if (parsed[0]!.kind !== 'paragraph') throw new Error('unreachable');
    expect(parsed[0]!.paragraph.runs[0]!.math).toMatchObject({
      type: 'row',
      children: [
        {
          type: 'eqArr',
          rows: [
            { type: 'row', children: [{ type: 'run', text: 'a=1' }] },
            { type: 'row', children: [{ type: 'run', text: 'b=2' }] },
          ],
        },
      ],
    });
  });

  it('stacks the rows vertically, flush-left, straddling the axis', () => {
    const node: MathNode = {
      type: 'eqArr',
      rows: [
        { type: 'row', children: [{ type: 'run', text: 'a' }] },
        { type: 'row', children: [{ type: 'run', text: 'bb' }] },
      ],
    };
    const box = layoutMath(node, { sizePt: 12 }, measure);
    const glyphs = box.items.filter((it) => it.kind === 'glyph');
    expect(glyphs.length).toBeGreaterThanOrEqual(2); // one glyph run per row
    // Two distinct vertical bands → the rows are stacked.
    const ys = new Set(glyphs.map((g) => Math.round(g.y)));
    expect(ys.size).toBeGreaterThanOrEqual(2);
    // Both rows start flush-left at x ≈ 0.
    const minX = Math.min(...glyphs.map((g) => g.x));
    expect(minX).toBeCloseTo(0, 1);
    expect(box.ascent).toBeGreaterThan(0);
    expect(box.descent).toBeGreaterThan(0);
  });
});

const SUM = `<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr>
  <m:sub><m:r><m:t>k=0</m:t></m:r></m:sub>
  <m:sup><m:r><m:t>n</m:t></m:r></m:sup>
  <m:e><m:r><m:t>k</m:t></m:r></m:e></m:nary>`;

describe('n-ary operators & functions', () => {
  it('parses an n-ary operator with its character and limits (m:val via m: namespace)', () => {
    const docx = buildDocxFromBody(`<w:p>${oMath(SUM)}</w:p>`);
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    if (parsed[0]!.kind !== 'paragraph') throw new Error('unreachable');
    expect(parsed[0]!.paragraph.runs[0]!.math).toMatchObject({
      type: 'row',
      children: [
        {
          type: 'nary',
          op: '∑',
          sub: { type: 'row', children: [{ type: 'run', text: 'k=0' }] },
          sup: { type: 'row', children: [{ type: 'run', text: 'n' }] },
          body: { type: 'row', children: [{ type: 'run', text: 'k' }] },
        },
      ],
    });
  });

  it('parses a function application', () => {
    const fn = `<m:func><m:fName><m:r><m:t>sin</m:t></m:r></m:fName><m:e><m:r><m:t>x</m:t></m:r></m:e></m:func>`;
    const docx = buildDocxFromBody(`<w:p>${oMath(fn)}</w:p>`);
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    if (parsed[0]!.kind !== 'paragraph') throw new Error('unreachable');
    expect(parsed[0]!.paragraph.runs[0]!.math).toMatchObject({
      type: 'row',
      children: [
        {
          type: 'func',
          name: { children: [{ text: 'sin' }] },
          body: { children: [{ text: 'x' }] },
        },
      ],
    });
  });

  it('draws the big operator as a path and is taller than ordinary text', () => {
    const nary: MathNode = {
      type: 'nary',
      op: '∑',
      body: { type: 'row', children: [{ type: 'run', text: 'k' }] },
      sub: { type: 'row', children: [{ type: 'run', text: 'k=0' }] },
      sup: { type: 'row', children: [{ type: 'run', text: 'n' }] },
    };
    const box = layoutMath(nary, { sizePt: 12 }, measure);
    expect(box.items.some((it) => it.kind === 'path')).toBe(true); // drawn ∑
    expect(box.ascent).toBeGreaterThan(12 * 0.72); // taller than a plain glyph
  });
});

describe('display equations', () => {
  it('centres a paragraph holding an m:oMathPara', () => {
    const body = `<w:p><m:oMathPara ${M_NS}><m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara></w:p>`;
    const parsed = parseDocument(OpcPackage.open(buildDocxFromBody(body)).getMainDocument().data);
    expect(parsed[0]!.kind).toBe('paragraph');
    if (parsed[0]!.kind !== 'paragraph') throw new Error('unreachable');
    expect(parsed[0]!.paragraph.properties.alignment).toBe('center');
    expect(parsed[0]!.paragraph.runs[0]!.math).toBeDefined();
  });
});

describe('delimiters, matrices & accents', () => {
  it('parses delimiters with custom brackets', () => {
    const d = `<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e><m:r><m:t>a+b</m:t></m:r></m:e></m:d>`;
    const docx = buildDocxFromBody(`<w:p>${oMath(d)}</w:p>`);
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    if (parsed[0]!.kind !== 'paragraph') throw new Error('unreachable');
    expect(parsed[0]!.paragraph.runs[0]!.math).toMatchObject({
      type: 'row',
      children: [
        {
          type: 'delimiter',
          begChr: '[',
          endChr: ']',
          children: [{ children: [{ text: 'a+b' }] }],
        },
      ],
    });
  });

  it('parses a 2×2 matrix', () => {
    const m = `<m:m><m:mr><m:e><m:r><m:t>1</m:t></m:r></m:e><m:e><m:r><m:t>2</m:t></m:r></m:e></m:mr><m:mr><m:e><m:r><m:t>3</m:t></m:r></m:e><m:e><m:r><m:t>4</m:t></m:r></m:e></m:mr></m:m>`;
    const docx = buildDocxFromBody(`<w:p>${oMath(m)}</w:p>`);
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    if (parsed[0]!.kind !== 'paragraph') throw new Error('unreachable');
    const math = parsed[0]!.paragraph.runs[0]!.math as MathNode;
    expect(math).toMatchObject({ type: 'row', children: [{ type: 'matrix' }] });
  });

  it('draws delimiters as stretchy bracket paths around the content', () => {
    const node: MathNode = {
      type: 'delimiter',
      begChr: '(',
      endChr: ')',
      children: [{ type: 'row', children: [{ type: 'run', text: 'a' }] }],
    };
    const box = layoutMath(node, { sizePt: 12 }, measure);
    expect(box.items.filter((it) => it.kind === 'path')).toHaveLength(2); // open + close
    expect(box.width).toBeGreaterThan(measure('a', 12, 'italic')); // brackets add width
  });

  it('lays out a matrix straddling the axis with all cells', () => {
    const node: MathNode = {
      type: 'matrix',
      rows: [
        [
          { type: 'row', children: [{ type: 'run', text: '1' }] },
          { type: 'row', children: [{ type: 'run', text: '2' }] },
        ],
        [
          { type: 'row', children: [{ type: 'run', text: '3' }] },
          { type: 'row', children: [{ type: 'run', text: '4' }] },
        ],
      ],
    };
    const box = layoutMath(node, { sizePt: 12 }, measure);
    expect(box.ascent).toBeGreaterThan(0);
    expect(box.descent).toBeGreaterThan(0);
    const glyphs = box.items
      .filter((it) => it.kind === 'glyph')
      .map((it) => it.text)
      .join('');
    expect(glyphs).toContain('1');
    expect(glyphs).toContain('4');
  });

  it('draws an accent over the base', () => {
    const node: MathNode = {
      type: 'accent',
      char: '̂',
      base: { type: 'row', children: [{ type: 'run', text: 'x' }] },
    };
    const box = layoutMath(node, { sizePt: 12 }, measure);
    expect(box.items.some((it) => it.kind === 'path')).toBe(true); // the hat
    expect(box.ascent).toBeGreaterThan(layoutMath(node.base, { sizePt: 12 }, measure).ascent);
  });
});

describe('math rendering (end-to-end)', () => {
  it('renders an inline fraction with a bar rule and glyphs', () => {
    const docx = buildDocxFromBody(
      `<w:p><w:r><w:t xml:space="preserve">x = </w:t></w:r>${FRACTION}</w:p>`,
    );
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    // The fraction bar is a filled rule (re … f).
    expect(text).toMatch(/[\d.]+ [\d.]+ [\d.]+ [\d.]+ re\nf/);
    // Glyphs are rendered (numerator / denominator / surrounding text).
    expect(text).toMatch(/<[0-9A-F]+> Tj/);
  });

  it('renders a superscript and a radical (stroked surd)', () => {
    const docx = buildDocxFromBody(`<w:p>${oMath(SUP + RAD)}</w:p>`);
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    expect(text).toMatch(/<[0-9A-F]+> Tj/); // glyphs (x, 2)
    expect(text).toMatch(/ l\nS\nQ/); // surd: a stroked open polyline
  });

  it('renders a summation with a drawn operator and limits', () => {
    const docx = buildDocxFromBody(`<w:p>${oMath(SUM)}</w:p>`);
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    expect(text).toMatch(/\nS\nQ/); // the ∑ stroke
    expect(text).toMatch(/<[0-9A-F]+> Tj/); // limit / body glyphs
  });
});
