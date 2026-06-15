// E-PPTX PX0 — the reader seam. A .pptx is sniffed and read into a FlowDoc whose
// page geometry is the deck's slide size, with one page per slide. Slide content
// arrives in PX1; this pins the wiring (sniff → reader → facade → Ream → PDF).

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildPptx } from './fixtures/build-pptx';
import { Ream } from '@/core/converter/ream';
import { PdfFile } from '@/pdf-reader/document';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

describe('pptx reader seam (E-PPTX PX0)', () => {
  it('sniffs a pptx and reports the format', () => {
    expect(Ream.parse(buildPptx(['', ''])).format).toBe('pptx');
  });

  it('renders one PDF page per slide at the 16:9 deck size', async () => {
    const pptx = buildPptx(['', '', '']); // 3 slides, default 16:9 deck
    const pdf = await Ream.parse(pptx).convert('pdf', { fonts: FONTS });
    const pages = PdfFile.parse(pdf).pages();
    expect(pages.length).toBe(3);
    // 12192000 × 6858000 EMU = 960 × 540 pt.
    const [x0, y0, x1, y1] = pages[0]!.mediaBox;
    expect(Math.round(x1 - x0)).toBe(960);
    expect(Math.round(y1 - y0)).toBe(540);
  });

  it('honours a 4:3 deck size', async () => {
    const pptx = buildPptx([''], { cx: 9144000, cy: 6858000 }); // 10" × 7.5"
    const pdf = await Ream.parse(pptx).convert('pdf', { fonts: FONTS });
    const [x0, y0, x1, y1] = PdfFile.parse(pdf).pages()[0]!.mediaBox;
    expect(Math.round(x1 - x0)).toBe(720); // 10in
    expect(Math.round(y1 - y0)).toBe(540); // 7.5in
  });
});
