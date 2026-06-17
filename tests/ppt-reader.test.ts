// PPT-1 — legacy `.ppt` text extraction. A PowerPoint 97–2003 binary file is a
// CFB holding a `PowerPoint Document` stream whose records are reached through the
// Current User → UserEditAtom → PersistDirectoryAtom indirection; each slide's
// text lives in TextCharsAtom (UTF-16) / TextBytesAtom (cp1252) atoms. These build
// a fixture `.ppt` and assert the reader walks that indirection back to the slide
// text, makes one page per slide, and renders to PDF / HTML.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildPpt } from './fixtures/build-ppt';
import { buildCfb } from './fixtures/build-cfb';
import type { FlowDoc } from '@/core/ir/flow';

import { extractPptContent } from '@/pptx/ppt/ppt-text';
import { pptReader, readPpt } from '@/pptx/ppt/ppt-reader';
import { Ream } from '@/core/converter/ream';
import { createConverter } from '@/core/converter/facade';

const ZERO_WIDTH_SPACE = '​';

const fonts = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

// The visible (non-empty, non-anchor) paragraph texts of the FlowDoc body.
function visibleTexts(doc: FlowDoc): Array<string> {
  return doc.body
    .filter((el) => el.kind === 'paragraph')
    .map((el) => el.paragraph.runs.map((r) => r.text).join(''))
    .filter((t) => t.length > 0 && t !== ZERO_WIDTH_SPACE);
}

// One page per slide: the page count is one plus the number of page breaks.
function pageCount(doc: FlowDoc): number {
  return (
    1 +
    doc.body.filter((el) => el.kind === 'paragraph' && el.paragraph.properties.pageBreakBefore)
      .length
  );
}

describe('ppt reader (PPT-1)', () => {
  it('reads inline slide text and makes one page per slide', () => {
    const doc = readPpt(buildPpt([{ text: 'First slide' }, { text: 'Second slide' }])).doc;
    expect(visibleTexts(doc)).toEqual(['First slide', 'Second slide']);
    expect(pageCount(doc)).toBe(2);
  });

  it('splits a slide text run into paragraphs at the CR mark', () => {
    const doc = readPpt(buildPpt([{ text: 'Title\rBullet one\rBullet two' }])).doc;
    expect(visibleTexts(doc)).toEqual(['Title', 'Bullet one', 'Bullet two']);
    expect(pageCount(doc)).toBe(1);
  });

  it('reads a TextBytesAtom (cp1252) as well as TextCharsAtom (UTF-16)', () => {
    // 0x97 is an em dash in cp1252; the high-range bytes must decode, not pass through.
    const doc = readPpt(buildPpt([{ textBytes: 'Plain text — dash' }])).doc;
    expect(visibleTexts(doc)).toEqual(['Plain text — dash']);
  });

  it('finds text nested inside a PPDrawing container (recursive descent)', () => {
    const doc = readPpt(buildPpt([{ text: 'Nested in a drawing', nested: true }])).doc;
    expect(visibleTexts(doc)).toEqual(['Nested in a drawing']);
  });

  it('falls back to the slide-list outline text when a slide stores none inline', () => {
    const doc = readPpt(buildPpt([{ outline: 'Outline only' }])).doc;
    expect(visibleTexts(doc)).toEqual(['Outline only']);
  });

  it('prefers a slide container inline text over the outline text', () => {
    const doc = readPpt(buildPpt([{ text: 'Inline wins', outline: 'Outline loses' }])).doc;
    expect(visibleTexts(doc)).toEqual(['Inline wins']);
  });

  it('emits an empty page for a slide with no text', () => {
    const doc = readPpt(buildPpt([{ text: 'Has text' }, {}])).doc;
    expect(visibleTexts(doc)).toEqual(['Has text']);
    expect(pageCount(doc)).toBe(2);
  });

  it('reads the slide size into the page size (master units → points)', () => {
    const def = readPpt(buildPpt([{ text: 'a' }])).doc;
    // Default 10in × 7.5in deck → 720 × 540 pt.
    expect(def.section.pageSize).toEqual({ width: 720, height: 540 });
    // 16:9 on-screen deck: 13⅓in × 7.5in → 960 × 540 pt.
    const wide = readPpt(buildPpt([{ text: 'a' }], { slideSizeInches: { w: 40 / 3, h: 7.5 } })).doc;
    expect(wide.section.pageSize).toEqual({ width: 960, height: 540 });
  });

  it('reads slides when the Current User stream is missing (scan fallback)', () => {
    const content = extractPptContent(
      buildPpt([{ text: 'Found by scan' }], { omitCurrentUser: true }),
    );
    expect(content.slides.flatMap((s) => s.paragraphs.map((p) => p.text))).toContain(
      'Found by scan',
    );
  });

  it('reports an encrypted deck as a dropped loss and reads no text', () => {
    const { doc, losses } = readPpt(buildPpt([{ text: 'secret' }], { encrypted: true }));
    expect(visibleTexts(doc)).toEqual([]);
    expect(losses.some((l) => l.severity === 'dropped')).toBe(true);
  });

  it('records the text-only loss for a normal deck', () => {
    const { losses } = readPpt(buildPpt([{ text: 'hi' }]));
    expect(losses.some((l) => l.severity === 'degraded' && /legacy \.ppt/.test(l.detail))).toBe(
      true,
    );
  });

  it('sniffs a .ppt and is detected by the converter', () => {
    const ppt = buildPpt([{ text: 'hi' }]);
    expect(pptReader.sniff(ppt)).toBe(true);
    expect(createConverter().detect(ppt)?.id).toBe('ppt');
    expect(Ream.parse(ppt).format).toBe('ppt');
  });

  it('does not sniff a non-PowerPoint compound file', () => {
    const docLike = buildCfb([{ name: 'WordDocument', data: new Uint8Array(2000) }]);
    expect(pptReader.sniff(docLike)).toBe(false);
  });

  it('converts a .ppt to PDF through the public API', async () => {
    const pdf = await Ream.parse(buildPpt([{ text: 'Slide text' }, { text: 'More' }])).convert(
      'pdf',
      { fonts },
    );
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });

  it('converts a .ppt to HTML carrying the slide text', async () => {
    const html = await Ream.parse(buildPpt([{ text: 'Hello deck' }])).convert('html');
    expect(new TextDecoder().decode(html)).toContain('Hello deck');
  });
});
