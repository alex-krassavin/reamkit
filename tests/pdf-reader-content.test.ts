// E-PDF EP2 — content-stream interpreter. Hand-built content streams exercise
// the text/line matrices, the show operators (Tj / TJ / ' / "), the CTM (cm /
// q / Q) and the per-glyph advance. With no font map every font falls back to a
// Latin-1 decode and a half-em (500/1000) advance.

import { describe, expect, it } from 'vitest';

import type { ContentFont, TextRun } from '@/pdf-reader/content';
import { interpretContent } from '@/pdf-reader/content';

const run = (cs: string, fonts = new Map<string, ContentFont>()): Array<TextRun> =>
  interpretContent(new TextEncoder().encode(cs), fonts).texts;

describe('content-stream interpreter (E-PDF EP2)', () => {
  it('extracts a positioned string from Td + Tj', () => {
    const runs = run('BT /F1 12 Tf 100 700 Td (Hello) Tj ET');
    expect(runs).toHaveLength(1);
    expect(runs[0]).toMatchObject({ text: 'Hello', fontKey: 'F1' });
    expect(runs[0]!.x).toBeCloseTo(100);
    expect(runs[0]!.y).toBeCloseTo(700);
    expect(runs[0]!.fontSizePt).toBeCloseTo(12);
  });

  it('honours an explicit text matrix (Tm)', () => {
    const runs = run('BT /F1 24 Tf 1 0 0 1 50 500 Tm (Hi) Tj ET');
    expect(runs[0]!.x).toBeCloseTo(50);
    expect(runs[0]!.y).toBeCloseTo(500);
    expect(runs[0]!.fontSizePt).toBeCloseTo(24);
  });

  it('concatenates a TJ array into one run at its origin', () => {
    const runs = run('BT /F1 12 Tf 10 20 Td [(Wo) -30 (rld)] TJ ET');
    expect(runs).toHaveLength(1);
    expect(runs[0]).toMatchObject({ text: 'World' });
    expect(runs[0]!.x).toBeCloseTo(10);
    expect(runs[0]!.y).toBeCloseTo(20);
  });

  it('advances the line matrix across Td-separated lines', () => {
    const runs = run('BT /F1 12 Tf 0 100 Td (A) Tj 0 -14 Td (B) Tj ET');
    expect(runs.map((r) => r.text)).toEqual(['A', 'B']);
    expect(runs[0]!.y).toBeCloseTo(100);
    expect(runs[1]!.x).toBeCloseTo(0); // a fresh Td resets x to the line start
    expect(runs[1]!.y).toBeCloseTo(86); // 100 − 14
  });

  it('composes the CTM (cm) with the text matrix', () => {
    const runs = run('q 1 0 0 1 0 50 cm BT /F1 12 Tf 0 0 Td (X) Tj ET Q');
    expect(runs[0]!.x).toBeCloseTo(0);
    expect(runs[0]!.y).toBeCloseTo(50); // shifted up by the cm translation
  });

  it('scales the effective font size by the matrix', () => {
    const runs = run('BT /F1 10 Tf 2 0 0 2 0 0 Tm (S) Tj ET');
    expect(runs[0]!.fontSizePt).toBeCloseTo(20); // 10 × matrix y-scale 2
  });

  it('decodes two-byte codes through a Type0 font', () => {
    // A trivial Identity-style font: 2-byte codes, code → char of that value.
    const font: ContentFont = {
      bytesPerCode: 2,
      decode: (codes) => codes.map((c) => String.fromCharCode(c)).join(''),
      width: () => 1000,
    };
    // <00480049> = U+0048 U+0049 = "HI"
    const runs = run('BT /F1 12 Tf 0 0 Td <00480049> Tj ET', new Map([['F1', font]]));
    expect(runs[0]!.text).toBe('HI');
  });

  it('tags runs with the enclosing BDC marked-content id (E-PDF EP3)', () => {
    const runs = run('/P <</MCID 2>> BDC BT /F1 12 Tf 0 0 Td (Tagged) Tj ET EMC');
    expect(runs[0]).toMatchObject({ text: 'Tagged', mcid: 2 });
  });

  it('leaves artifact text untagged (E-PDF EP3)', () => {
    const runs = run('/Artifact <</Type /Pagination>> BDC BT /F1 12 Tf 0 0 Td (Footer) Tj ET EMC');
    expect(runs[0]!.text).toBe('Footer');
    expect(runs[0]!.mcid).toBeUndefined();
  });
});
