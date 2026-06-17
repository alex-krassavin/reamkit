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

import { extractPptContent, paragraphText } from '@/pptx/ppt/ppt-text';
import { pptReader, readPpt } from '@/pptx/ppt/ppt-reader';
import { Ream } from '@/core/converter/ream';
import { createConverter } from '@/core/converter/facade';

const ZERO_WIDTH_SPACE = '​';

const fonts = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

// The smallest image the decoder accepts — a 1×1 transparent PNG.
const PNG_1x1 = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
    'base64',
  ),
);

// The image blocks in the FlowDoc body.
function imageBlocks(doc: FlowDoc) {
  return doc.body.filter((el) => el.kind === 'image');
}

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

// The first body paragraph (the first slide's first line).
function firstParagraph(doc: FlowDoc) {
  const p = doc.body.find((el) => el.kind === 'paragraph');
  if (!p) throw new Error('no paragraph');
  return p.paragraph;
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
    expect(
      content.slides.flatMap((s) =>
        s.shapes.flatMap((sh) => sh.paragraphs ?? []).map(paragraphText),
      ),
    ).toContain('Found by scan');
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

describe('ppt reader formatting (PPT-2)', () => {
  it('reads bold / italic / underline from the StyleTextPropAtom', () => {
    const doc = readPpt(
      buildPpt([
        { text: 'Styled', charRuns: [{ length: 6, bold: true, italic: true, underline: true }] },
      ]),
    ).doc;
    const run = firstParagraph(doc).runs[0]!;
    expect(run.text).toBe('Styled');
    expect(run.properties.bold).toBe(true);
    expect(run.properties.italic).toBe(true);
    expect(run.properties.underline).toBe('single');
  });

  it('reads the font size (points) of a run', () => {
    const doc = readPpt(buildPpt([{ text: 'Big', charRuns: [{ length: 3, sizePt: 40 }] }])).doc;
    expect(firstParagraph(doc).runs[0]!.properties.fontSizePt).toBe(40);
  });

  it('reads an explicit RGB run colour (ColorIndexStruct index 0xFE)', () => {
    const doc = readPpt(
      buildPpt([{ text: 'Red', charRuns: [{ length: 3, colorHex: 'FF0000' }] }]),
    ).doc;
    expect(firstParagraph(doc).runs[0]!.properties.colorHex).toBe('FF0000');
  });

  it('splits a line into runs at character-run boundaries', () => {
    const doc = readPpt(
      buildPpt([{ text: 'AB', charRuns: [{ length: 1, bold: true }, { length: 1 }] }]),
    ).doc;
    const runs = firstParagraph(doc).runs;
    expect(runs.map((r) => r.text)).toEqual(['A', 'B']);
    expect(runs[0]!.properties.bold).toBe(true);
    expect(runs[1]!.properties.bold).toBeFalsy();
  });

  it('reads paragraph alignment from the paragraph run', () => {
    const doc = readPpt(buildPpt([{ text: 'Centered', paraRuns: [{ length: 8, align: 1 }] }])).doc;
    expect(firstParagraph(doc).properties.alignment).toBe('center');
  });

  it('indents a paragraph by its outline level', () => {
    const doc = readPpt(buildPpt([{ text: 'Indented', paraRuns: [{ length: 8, level: 2 }] }])).doc;
    expect(firstParagraph(doc).properties.indentLeft).toBe(36); // 2 × 18pt
  });

  it('keeps per-paragraph formatting across a CR-split run', () => {
    // "Title\rBody": a centered title then a left body, bold title run.
    const doc = readPpt(
      buildPpt([
        {
          text: 'Title\rBody',
          charRuns: [{ length: 6, bold: true }, { length: 4 }],
          paraRuns: [{ length: 6, align: 1 }, { length: 4 }],
        },
      ]),
    ).doc;
    const paras = doc.body.filter((el) => el.kind === 'paragraph');
    expect(paras[0]!.paragraph.runs.map((r) => r.text).join('')).toBe('Title');
    expect(paras[0]!.paragraph.properties.alignment).toBe('center');
    expect(paras[0]!.paragraph.runs[0]!.properties.bold).toBe(true);
    expect(paras[1]!.paragraph.runs.map((r) => r.text).join('')).toBe('Body');
    expect(paras[1]!.paragraph.runs[0]!.properties.bold).toBeFalsy();
  });
});

describe('ppt reader images (PPT-3)', () => {
  it('reads an embedded picture referenced by a slide shape', () => {
    const doc = readPpt(buildPpt([{ imageRef: 1 }], { images: [PNG_1x1] })).doc;
    const imgs = imageBlocks(doc);
    expect(imgs).toHaveLength(1);
    expect(imgs[0]!.image.resource).toBeDefined();
    expect(imgs[0]!.image.width).toBeGreaterThan(0);
    expect(imgs[0]!.image.height).toBeGreaterThan(0);
    // The bytes round-trip through the ResourceStore.
    const stored = doc.resources.get(imgs[0]!.image.resource!);
    expect(stored && stored[0]).toBe(0x89); // PNG signature
  });

  it('emits a slide image after the slide text', () => {
    const doc = readPpt(buildPpt([{ text: 'Caption', imageRef: 1 }], { images: [PNG_1x1] })).doc;
    const kinds = doc.body.map((el) => el.kind);
    expect(kinds).toEqual(['paragraph', 'image']);
  });

  it('places each slide image on its own page', () => {
    const doc = readPpt(
      buildPpt([{ imageRef: 1 }, { imageRef: 2 }], { images: [PNG_1x1, PNG_1x1] }),
    ).doc;
    expect(imageBlocks(doc)).toHaveLength(2);
    expect(pageCount(doc)).toBe(2);
  });

  it('skips a picture whose bytes are not a decodable image', () => {
    const doc = readPpt(
      buildPpt([{ imageRef: 1 }], { images: [Uint8Array.of(1, 2, 3, 4, 5, 6, 7, 8)] }),
    ).doc;
    expect(imageBlocks(doc)).toHaveLength(0);
  });

  it('converts a .ppt with an image to PDF', async () => {
    const pdf = await Ream.parse(
      buildPpt([{ text: 'Pic', imageRef: 1 }], { images: [PNG_1x1] }),
    ).convert('pdf', { fonts });
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });
});

describe('ppt reader placement (PPT-4)', () => {
  it('positions an anchored text box as a floating shape at its rectangle', () => {
    const doc = readPpt(
      buildPpt([{ boxes: [{ anchor: { x: 100, y: 50, w: 200, h: 80 }, text: 'Positioned' }] }]),
    ).doc;
    const shapes = doc.body.filter((el) => el.kind === 'shape');
    expect(shapes).toHaveLength(1);
    const shape = shapes[0]!.shape;
    expect(shape.float?.posH?.offsetPt).toBe(100);
    expect(shape.float?.posV?.offsetPt).toBe(50);
    expect(shape.width).toBe(200);
    expect(shape.height).toBe(80);
    const text = (shape.text?.content ?? [])
      .filter((el) => el.kind === 'paragraph')
      .map((el) => el.paragraph.runs.map((r) => r.text).join(''))
      .join('');
    expect(text).toBe('Positioned');
  });

  it('positions an anchored picture as a floating image at its rectangle', () => {
    const doc = readPpt(
      buildPpt([{ boxes: [{ anchor: { x: 10, y: 20, w: 300, h: 200 }, imageRef: 1 }] }], {
        images: [PNG_1x1],
      }),
    ).doc;
    const imgs = imageBlocks(doc);
    expect(imgs).toHaveLength(1);
    const image = imgs[0]!.image;
    expect(image.float?.posH?.offsetPt).toBe(10);
    expect(image.float?.posV?.offsetPt).toBe(20);
    expect(image.width).toBe(300);
    expect(image.height).toBe(200);
  });

  it('leaves an un-anchored shape in reading-order flow (no float)', () => {
    const doc = readPpt(buildPpt([{ boxes: [{ text: 'No anchor' }] }])).doc;
    expect(doc.body.filter((el) => el.kind === 'shape')).toHaveLength(0);
    expect(visibleTexts(doc)).toContain('No anchor');
  });

  it('gives each floating-only slide its own page', () => {
    const doc = readPpt(
      buildPpt([
        { boxes: [{ anchor: { x: 0, y: 0, w: 100, h: 50 }, text: 'A' }] },
        { boxes: [{ anchor: { x: 0, y: 0, w: 100, h: 50 }, text: 'B' }] },
      ]),
    ).doc;
    expect(doc.body.filter((el) => el.kind === 'shape')).toHaveLength(2);
    expect(pageCount(doc)).toBe(2);
  });

  it('converts a .ppt with positioned shapes to PDF', async () => {
    const pdf = await Ream.parse(
      buildPpt([{ boxes: [{ anchor: { x: 50, y: 50, w: 200, h: 100 }, text: 'Slide' }] }]),
    ).convert('pdf', { fonts });
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });
});
