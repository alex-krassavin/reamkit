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
import { pt } from '@/core/ir/units';

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
    expect(def.section!.pageSize).toEqual({ width: 720, height: 540 });
    // 16:9 on-screen deck: 13⅓in × 7.5in → 960 × 540 pt.
    const wide = readPpt(buildPpt([{ text: 'a' }], { slideSizeInches: { w: 40 / 3, h: 7.5 } })).doc;
    expect(wide.section!.pageSize).toEqual({ width: 960, height: 540 });
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

describe('ppt reader autoshapes (PPT-5)', () => {
  it('reads a filled autoshape as a positioned vector shape', () => {
    const doc = readPpt(
      buildPpt([
        {
          boxes: [
            { anchor: { x: 10, y: 10, w: 100, h: 50 }, shapeType: 1, fillColorHex: 'FF8800' },
          ],
        },
      ]),
    ).doc;
    const shapes = doc.body.filter((el) => el.kind === 'shape');
    expect(shapes).toHaveLength(1);
    const shape = shapes[0]!.shape;
    expect(shape.geometry.preset).toBe('rect');
    expect(shape.fill).toEqual({ kind: 'solid', colorHex: 'FF8800' });
    expect(shape.float?.posH?.offsetPt).toBe(10);
    expect(shape.width).toBe(100);
  });

  it('reads the line colour of an autoshape', () => {
    const doc = readPpt(
      buildPpt([
        { boxes: [{ anchor: { x: 0, y: 0, w: 100, h: 2 }, shapeType: 1, lineColorHex: '0000FF' }] },
      ]),
    ).doc;
    const shape = doc.body.find((el) => el.kind === 'shape')!.shape;
    expect(shape.line?.colorHex).toBe('0000FF');
  });

  it('maps the MSOSPT shape type to a preset geometry', () => {
    const doc = readPpt(
      buildPpt([
        { boxes: [{ anchor: { x: 0, y: 0, w: 80, h: 80 }, shapeType: 3, fillColorHex: '00FF00' }] },
      ]),
    ).doc;
    expect(doc.body.find((el) => el.kind === 'shape')!.shape.geometry.preset).toBe('ellipse');
  });

  it('skips an autoshape with no literal fill or line colour', () => {
    const doc = readPpt(
      buildPpt([
        {
          boxes: [
            { anchor: { x: 0, y: 0, w: 100, h: 50 }, shapeType: 1 }, // no colour → skipped
            { anchor: { x: 0, y: 60, w: 100, h: 50 }, shapeType: 1, fillColorHex: 'FF0000' },
          ],
        },
      ]),
    ).doc;
    expect(doc.body.filter((el) => el.kind === 'shape')).toHaveLength(1);
  });

  it('converts a .ppt with an autoshape to PDF', async () => {
    const pdf = await Ream.parse(
      buildPpt([
        {
          boxes: [
            {
              anchor: { x: 50, y: 50, w: 200, h: 100 },
              shapeType: 2,
              fillColorHex: '3366CC',
              lineColorHex: '000000',
            },
          ],
        },
      ]),
    ).convert('pdf', { fonts });
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });
});

describe('ppt reader scheme colours (PPT-6)', () => {
  // Distinct values per slot so a wrong index is detectable; slot 4 = fills.
  const SCHEME = ['000000', '111111', '222222', '808080', '336699', '444444', '555555', '666666'];

  it('resolves a shape fill from the slide colour scheme', () => {
    const doc = readPpt(
      buildPpt([
        {
          colorScheme: SCHEME,
          boxes: [{ anchor: { x: 10, y: 10, w: 100, h: 50 }, shapeType: 1, fillSchemeIndex: 4 }],
        },
      ]),
    ).doc;
    const shape = doc.body.find((el) => el.kind === 'shape')!.shape;
    expect(shape.fill).toEqual({ kind: 'solid', colorHex: '336699' });
  });

  it('resolves a shape line colour from the scheme (slot 1)', () => {
    const doc = readPpt(
      buildPpt([
        {
          colorScheme: SCHEME,
          boxes: [{ anchor: { x: 0, y: 0, w: 100, h: 2 }, shapeType: 1, lineSchemeIndex: 1 }],
        },
      ]),
    ).doc;
    expect(doc.body.find((el) => el.kind === 'shape')!.shape.line?.colorHex).toBe('111111');
  });

  it('follows the master colour scheme when fMasterScheme is set', () => {
    // The master's slot 4 differs from the slide's own — the master must win.
    const master = ['000000', '111111', '222222', '808080', 'aa1122', '444444', '555555', '666666'];
    const doc = readPpt(
      buildPpt(
        [
          {
            colorScheme: SCHEME, // a decoy own scheme
            followMasterScheme: true,
            boxes: [{ anchor: { x: 5, y: 5, w: 90, h: 40 }, shapeType: 1, fillSchemeIndex: 4 }],
          },
        ],
        { masters: [{ colorScheme: master }] },
      ),
    ).doc;
    expect(doc.body.find((el) => el.kind === 'shape')!.shape.fill).toEqual({
      kind: 'solid',
      colorHex: 'AA1122',
    });
  });

  it('drops a scheme colour when no scheme is present (no wrong colour)', () => {
    const doc = readPpt(
      buildPpt([
        { boxes: [{ anchor: { x: 0, y: 0, w: 100, h: 50 }, shapeType: 1, fillSchemeIndex: 4 }] },
      ]),
    ).doc;
    expect(doc.body.filter((el) => el.kind === 'shape')).toHaveLength(0);
  });

  it('converts a .ppt with a scheme-coloured shape to PDF', async () => {
    const pdf = await Ream.parse(
      buildPpt([
        {
          colorScheme: SCHEME,
          boxes: [
            {
              anchor: { x: 50, y: 50, w: 200, h: 100 },
              shapeType: 2,
              fillSchemeIndex: 4,
              lineSchemeIndex: 1,
            },
          ],
        },
      ]),
    ).convert('pdf', { fonts });
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });
});

describe('ppt reader custom geometry (PPT-7)', () => {
  // The segment opcodes (top 3 bits = MSOPATHTYPE): moveTo 0x4000, lineTo 0x0001,
  // curveTo 0x2001, close 0x6001.
  const MOVE = 0x4000;
  const LINE = 0x0001;
  const CURVE = 0x2001;
  const CLOSE = 0x6001;

  it('reads a freeform triangle as exact custom geometry', () => {
    const doc = readPpt(
      buildPpt([
        {
          boxes: [
            {
              anchor: { x: 10, y: 10, w: 200, h: 100 },
              shapeType: 0,
              lineColorHex: '0000FF',
              freeform: {
                geoRight: 200,
                geoBottom: 100,
                vertices: [
                  [0, 0],
                  [200, 0],
                  [100, 100],
                ],
                segments: [MOVE, LINE, LINE, CLOSE],
              },
            },
          ],
        },
      ]),
    ).doc;
    const shape = doc.body.find((el) => el.kind === 'shape')!.shape;
    expect(shape.geometry.kind).toBe('custom');
    expect(shape.geometry.custom).toEqual({
      pathWidth: 200,
      pathHeight: 100,
      commands: [
        { cmd: 'move', x: 0, y: 0 },
        { cmd: 'line', x: 200, y: 0 },
        { cmd: 'line', x: 100, y: 100 },
        { cmd: 'close' },
      ],
    });
  });

  it('reads a curveTo segment as a cubic bezier (three points)', () => {
    const doc = readPpt(
      buildPpt([
        {
          boxes: [
            {
              anchor: { x: 0, y: 0, w: 100, h: 50 },
              shapeType: 0,
              lineColorHex: '000000',
              freeform: {
                geoRight: 100,
                geoBottom: 50,
                vertices: [
                  [0, 0],
                  [10, 50],
                  [90, 50],
                  [100, 0],
                ],
                segments: [MOVE, CURVE],
              },
            },
          ],
        },
      ]),
    ).doc;
    const shape = doc.body.find((el) => el.kind === 'shape')!.shape;
    expect(shape.geometry.custom?.commands).toEqual([
      { cmd: 'move', x: 0, y: 0 },
      { cmd: 'cubic', x1: 10, y1: 50, x2: 90, y2: 50, x: 100, y: 0 },
    ]);
  });

  it('skips a render-hint escape without consuming a vertex', () => {
    // 0xAC00 is an escape (top bits 101) with sub-code 0x0C (auto-line) — a hint
    // that pulls no point; the second lineTo must still land on the third vertex.
    const doc = readPpt(
      buildPpt([
        {
          boxes: [
            {
              anchor: { x: 0, y: 0, w: 100, h: 100 },
              shapeType: 0,
              lineColorHex: '000000',
              freeform: {
                geoRight: 100,
                geoBottom: 100,
                vertices: [
                  [0, 0],
                  [100, 0],
                  [100, 100],
                ],
                segments: [MOVE, LINE, 0xac00, LINE, CLOSE],
              },
            },
          ],
        },
      ]),
    ).doc;
    const shape = doc.body.find((el) => el.kind === 'shape')!.shape;
    expect(shape.geometry.custom?.commands).toEqual([
      { cmd: 'move', x: 0, y: 0 },
      { cmd: 'line', x: 100, y: 0 },
      { cmd: 'line', x: 100, y: 100 },
      { cmd: 'close' },
    ]);
  });

  it('falls back to preset geometry on an unmodelled arc escape (never wrong)', () => {
    // 0xA300 is an escape with sub-code 0x03 (arc-to) — it would consume points for
    // a curve we do not synthesize, so the whole custom path bails to the preset.
    const doc = readPpt(
      buildPpt([
        {
          boxes: [
            {
              anchor: { x: 0, y: 0, w: 100, h: 100 },
              shapeType: 1,
              fillColorHex: 'FF0000',
              freeform: {
                geoRight: 100,
                geoBottom: 100,
                vertices: [
                  [0, 0],
                  [100, 0],
                ],
                segments: [MOVE, 0xa300],
              },
            },
          ],
        },
      ]),
    ).doc;
    const shape = doc.body.find((el) => el.kind === 'shape')!.shape;
    expect(shape.geometry.kind).toBe('preset');
    expect(shape.geometry.preset).toBe('rect');
  });

  it('converts a .ppt freeform to PDF', async () => {
    const pdf = await Ream.parse(
      buildPpt([
        {
          boxes: [
            {
              anchor: { x: 50, y: 50, w: 200, h: 120 },
              shapeType: 0,
              fillColorHex: '3366CC',
              freeform: {
                geoRight: 200,
                geoBottom: 120,
                vertices: [
                  [0, 0],
                  [200, 60],
                  [0, 120],
                ],
                segments: [MOVE, LINE, LINE, CLOSE],
              },
            },
          ],
        },
      ]),
    ).convert('pdf', { fonts });
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });
});

describe('ppt reader system colours (PPT-8)', () => {
  it('resolves a shape fill from a Windows system colour (windowText → black)', () => {
    const doc = readPpt(
      buildPpt([
        { boxes: [{ anchor: { x: 0, y: 0, w: 80, h: 40 }, shapeType: 1, fillSysColor: 8 }] },
      ]),
    ).doc;
    expect(doc.body.find((el) => el.kind === 'shape')!.shape.fill).toEqual({
      kind: 'solid',
      colorHex: '000000',
    });
  });

  it('skips a procedural system colour (0xF0–0xF7) rather than guess', () => {
    const doc = readPpt(
      buildPpt([
        { boxes: [{ anchor: { x: 0, y: 0, w: 80, h: 40 }, shapeType: 1, fillSysColor: 0xf0 }] },
      ]),
    ).doc;
    // No resolvable colour ⇒ no autoshape is emitted (a shape needs a fill or line).
    expect(doc.body.filter((el) => el.kind === 'shape')).toHaveLength(0);
  });
});

describe('ppt outline text (PPT-10)', () => {
  // PowerPoint stores a slide's TITLE AND BODY in the document's slide list, not
  // in the shape that draws it: the placeholder holds an OutlineTextRefAtom
  // naming which of the slide's texts is its own. Read without following it, a
  // whole deck came out as one un-anchored blob at the top-left corner.
  const deck = (): Uint8Array =>
    buildPpt([
      {
        outlineTexts: [
          { textType: 0, text: 'The title' },
          { textType: 1, text: 'The body' },
        ],
        boxes: [
          { anchor: { x: 60, y: 40, w: 500, h: 80 }, outlineRef: 0 },
          { anchor: { x: 60, y: 160, w: 500, h: 300 }, outlineRef: 1 },
        ],
      },
    ]);

  it('puts each outline text in the shape that references it', () => {
    const slide = extractPptContent(deck()).slides[0]!;
    expect(slide.shapes.map((s) => s.paragraphs?.map(paragraphText).join(''))).toEqual([
      'The title',
      'The body',
    ]);
    // …with the shape's own rectangle, instead of falling back to the flow.
    expect(slide.shapes.map((s) => s.rectPt?.y)).toEqual([40, 160]);
  });

  it('renders both as positioned shapes on one page', () => {
    const doc = readPpt(deck()).doc;
    const boxes = doc.body.filter((el) => el.kind === 'shape');
    expect(boxes).toHaveLength(2);
    expect(boxes.map((b) => b.shape.float?.posV?.offsetPt)).toEqual([pt(40), pt(160)]);
  });
});

describe('ppt master text styles (PPT-11)', () => {
  // A `.ppt` title is 44pt because the MASTER says so — the slide's own
  // StyleTextPropAtom usually states no size at all. §2.9.42 TextMasterStyleAtom
  // holds one style per indent level, per text type.
  const deck = (): Uint8Array =>
    buildPpt(
      [
        {
          masterIndex: 0,
          outlineTexts: [
            { textType: 0, text: 'Title' },
            { textType: 1, text: 'First\rSecond' },
          ],
          boxes: [
            { anchor: { x: 10, y: 10, w: 400, h: 60 }, outlineRef: 0 },
            { anchor: { x: 10, y: 90, w: 400, h: 200 }, outlineRef: 1 },
          ],
        },
      ],
      {
        masters: [
          {
            colorScheme: [
              'FFFFFF',
              '000000',
              '808080',
              '000000',
              'FF0000',
              '00FF00',
              '0000FF',
              'FFFF00',
            ],
            textStyles: [
              { textType: 0, sizesPt: [44] },
              { textType: 1, sizesPt: [32, 28] },
            ],
          },
        ],
      },
    );

  it('gives a run the size its level inherits from the master', () => {
    const slide = extractPptContent(deck()).slides[0]!;
    const sizes = slide.shapes.map((s) => s.paragraphs?.map((p) => p.runs[0]?.sizePt));
    expect(sizes).toEqual([[44], [32, 32]]);
  });

  it('takes the centre-title variant from the plain title style it varies', () => {
    // A master that states only the alignment of `centerTitle` still lends the
    // title's size: the styles merge property by property, not whole.
    const slide = extractPptContent(
      buildPpt(
        [
          {
            masterIndex: 0,
            outlineTexts: [{ textType: 5, text: 'Centred' }],
            boxes: [{ anchor: { x: 10, y: 10, w: 400, h: 60 }, outlineRef: 0 }],
          },
        ],
        {
          masters: [
            {
              colorScheme: [
                'FFFFFF',
                '000000',
                '808080',
                '000000',
                'FF0000',
                '00FF00',
                '0000FF',
                'FFFF00',
              ],
              textStyles: [{ textType: 0, sizesPt: [44] }],
            },
          ],
        },
      ),
    ).slides[0]!;
    expect(slide.shapes[0]?.paragraphs?.[0]?.runs[0]?.sizePt).toBe(44);
  });
});

describe('ppt slide background (PPT-12)', () => {
  // MS-ODRAW marks the background with `fBackground` on a shape like any other.
  // Read as content it is a rectangle among the slide's shapes; dropped, as it
  // was, a slide PowerPoint paints black comes out white — and on 41246-2 that
  // is 95 % of the page.
  const deck = (): Uint8Array =>
    buildPpt([
      {
        boxes: [
          { shapeType: 1, background: true, fillColorHex: '000000' },
          { anchor: { x: 100, y: 80, w: 200, h: 60 }, text: 'Fin' },
        ],
      },
    ]);

  it('reads it as the slide’s background, not as a shape on the slide', () => {
    const slide = extractPptContent(deck()).slides[0]!;
    expect(slide.background).toEqual({ fillColorHex: '000000' });
    expect(slide.shapes.map((s) => s.paragraphs?.map(paragraphText).join(''))).toEqual(['Fin']);
  });

  it('paints it behind the content, over the whole page', () => {
    const doc = readPpt(deck()).doc;
    const backdrop = doc.body.find((el) => el.kind === 'shape' && el.shape.float?.behind === true);
    expect(backdrop?.kind === 'shape' ? backdrop.shape.fill : undefined).toEqual({
      kind: 'solid',
      colorHex: '000000',
    });
    expect(backdrop?.kind === 'shape' ? backdrop.shape.width : undefined).toEqual(
      doc.section?.pageSize?.width,
    );
  });
});

describe('ppt shape fills the file says are there and are not (PPT-13)', () => {
  const MOVE = 0x4000;
  const LINE = 0x0001;
  const CLOSE = 0x6001;

  it('leaves a shape hollow when its boolean set clears fFilled', () => {
    // A shape states a fill colour whether or not it is filled. 37625's chart
    // states a red one on the outline every reader leaves hollow, and painting
    // it put a red slab over half the plot.
    const doc = readPpt(
      buildPpt([
        {
          boxes: [
            {
              anchor: { x: 0, y: 0, w: 80, h: 40 },
              shapeType: 1,
              fillColorHex: 'FF0000',
              lineColorHex: '000080',
              noFill: true,
            },
          ],
        },
      ]),
    ).doc;
    const shape = doc.body.find((el) => el.kind === 'shape')!.shape;
    expect(shape.fill).toEqual({ kind: 'none' });
    expect(shape.line?.colorHex).toBe('000080'); // …and the outline still draws
  });

  it('draws neither when the set clears fLine too', () => {
    const doc = readPpt(
      buildPpt([
        {
          boxes: [
            {
              anchor: { x: 0, y: 0, w: 80, h: 40 },
              shapeType: 1,
              fillColorHex: 'FF0000',
              lineColorHex: '000080',
              noFill: true,
              noLine: true,
            },
          ],
        },
      ]),
    ).doc;
    // Nothing left to draw ⇒ no shape at all, as for a colour that will not resolve.
    expect(doc.body.filter((el) => el.kind === 'shape')).toHaveLength(0);
  });

  it('reads a freeform whose arrays state their length without the header', () => {
    const geometry = (arrayLenExcludesHeader: boolean): number | undefined => {
      const slide = extractPptContent(
        buildPpt([
          {
            boxes: [
              {
                anchor: { x: 0, y: 0, w: 200, h: 120 },
                shapeType: 0,
                lineColorHex: '000080',
                freeform: {
                  geoRight: 200,
                  geoBottom: 120,
                  vertices: [
                    [0, 0],
                    [200, 60],
                    [0, 120],
                  ],
                  segments: [MOVE, LINE, LINE, CLOSE],
                  arrayLenExcludesHeader,
                },
              },
            ],
          },
        ]),
      ).slides[0]!;
      return slide.shapes[0]?.autoShape?.geometry?.commands.length;
    };
    // Stated either way, the path is the same four commands: the segment array
    // is found where the vertex array actually ends, not where its length says.
    expect(geometry(false)).toBe(4);
    expect(geometry(true)).toBe(4);
  });
});

describe('ppt master decoration (PPT-14)', () => {
  // §2.4.24 fMasterObjects — the rules, logos and footer band a master draws on
  // every slide that follows it. 37625's "teri" logo is one, and no slide of
  // that deck carried it.
  const scheme = ['FFFFFF', '000000', '808080', '000000', 'FF0000', '00FF00', '0000FF', 'FFFF00'];
  const deck = (followMasterObjects: boolean): Uint8Array =>
    buildPpt(
      [
        {
          masterIndex: 0,
          followMasterObjects,
          boxes: [{ anchor: { x: 10, y: 10, w: 300, h: 40 }, text: 'The slide' }],
        },
      ],
      {
        masters: [
          {
            colorScheme: scheme,
            boxes: [
              // Decoration: drawn on the slide.
              {
                anchor: { x: 600, y: 480, w: 60, h: 40 },
                shapeType: 1,
                fillColorHex: 'FF6600',
              },
              // A prototype: its prompt text belongs to no slide.
              { anchor: { x: 10, y: 10, w: 300, h: 40 }, text: 'Click to edit', placeholder: true },
            ],
          },
        ],
      },
    );

  it('draws the master’s shapes under the slide’s own, and not its placeholders', () => {
    const slide = extractPptContent(deck(true)).slides[0]!;
    expect(
      slide.shapes.map((s) => s.autoShape?.fillColorHex ?? paragraphText(s.paragraphs![0]!)),
    ).toEqual(['FF6600', 'The slide']);
  });

  it('…and none of them when the slide does not follow the master', () => {
    const slide = extractPptContent(deck(false)).slides[0]!;
    expect(slide.shapes).toHaveLength(1);
  });
});
