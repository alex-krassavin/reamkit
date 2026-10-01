// §18.8.19 — a colour's tint, as Excel works it. Its colour picker offers five
// shades of each theme colour; for the Office 2013–2022 theme these are the
// shades it names, and the tints it writes for them (in 32767ths) are the
// shades it shows again.

import { describe, expect, it } from 'vitest';

import { applyTint } from '@/excel/tint';
import { workbookColorHex } from '@/excel/styles-parser';

describe('applyTint (§18.8.19)', () => {
  it('gives the shades Excel picks, lighter and darker', () => {
    const shades: ReadonlyArray<[string, number, string]> = [
      ['4472C4', 0.8, 'D9E1F2'],
      ['4472C4', 0.6, 'B4C6E7'],
      ['4472C4', 0.4, '8EA9DB'],
      ['4472C4', -0.25, '305496'],
      ['4472C4', -0.5, '203764'],
      ['ED7D31', 0.8, 'FCE4D6'],
      ['ED7D31', -0.5, '833C0C'],
      ['44546A', 0.6, 'ACB9CA'],
      ['E7E6E6', -0.9, '161616'],
      ['FFC000', 0.4, 'FFD966'],
      ['70AD47', -0.25, '548235'],
      ['5B9BD5', 0.8, 'DDEBF7'],
    ];
    for (const [base, tint, shade] of shades) expect(applyTint(base, tint)).toBe(shade);
  });

  it('gives the greys of white and black', () => {
    expect(applyTint('FFFFFF', -0.05)).toBe('F2F2F2');
    expect(applyTint('FFFFFF', -0.15)).toBe('D9D9D9');
    expect(applyTint('FFFFFF', -0.5)).toBe('808080');
    expect(applyTint('000000', 0.5)).toBe('808080');
    expect(applyTint('000000', 0.35)).toBe('595959');
    expect(applyTint('000000', 0.05)).toBe('0D0D0D');
  });

  it('reads the tint Excel writes as the shade it was picked as', () => {
    // 80% lighter, 25% and 50% darker, as Excel stores them.
    expect(applyTint('ED7D31', 0.79998168889431442)).toBe('FCE4D6');
    expect(applyTint('FFFFFF', -0.249977111117893)).toBe('BFBFBF');
    expect(applyTint('000000', 0.499984740745262)).toBe('808080');
  });

  it("gives what Excel's own PDF paints, on the 2023 theme and off the picker", () => {
    // Measured 2026-10-01: Excel's PDF of a probe filling cells with theme and
    // rgb colours under tints of its own and others.
    expect(applyTint('156082', 0.8)).toBe('C0E6F5');
    expect(applyTint('156082', 0.6)).toBe('83CCEB');
    expect(applyTint('156082', 0.4)).toBe('44B3E1');
    expect(applyTint('0E2841', 0.3)).toBe('2467AA');
    expect(applyTint('E97132', 0.7)).toBe('F9D3BF');
    expect(applyTint('A02B93', -0.37)).toBe('651B5E');
    expect(applyTint('FFFFFF', -0.75)).toBe('404040');
    expect(applyTint('E8E8E8', -0.9)).toBe('161616');
  });

  it('tints an rgb colour as it does a theme slot', () => {
    const colors = { indexed: [] };
    expect(workbookColorHex({ '@_rgb': 'FF4472C4', '@_tint': '0.8' }, colors)).toBe('D9E1F2');
    expect(workbookColorHex({ '@_rgb': 'FF7F7F7F', '@_tint': '-0.5' }, colors)).toBe('404040');
  });

  it('leaves a colour alone at no tint, or none it can read', () => {
    expect(applyTint('4472C4', 0)).toBe('4472C4');
    expect(applyTint('4472C4', Number.NaN)).toBe('4472C4');
  });
});
