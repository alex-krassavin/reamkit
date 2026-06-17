// DOC-1 — legacy `.doc` text extraction. A Word 97–2003 binary file is a CFB
// holding a `WordDocument` stream; its text lives in pieces (the piece table /
// CLX) that are either 16-bit Unicode or 8-bit Windows-1252. These build a
// fixture `.doc` and assert the reader walks the FIB → piece table → pieces back
// to the original text, splits it into paragraphs, and renders to PDF.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDoc, buildPicf } from './fixtures/build-doc';
import type { BodyElement } from '@/core/document-model';
import type { FlowDoc } from '@/core/ir/flow';

import { docReader, readDoc } from '@/word/doc/doc-reader';
import { Ream } from '@/core/converter/ream';

const CM = String.fromCharCode(0x07); // table cell mark
const PIC = String.fromCharCode(0x01); // picture placeholder char
// The smallest image the decoder accepts — a 1×1 transparent PNG.
const PNG_1x1 = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
    'base64',
  ),
);

// The visible text of each paragraph in the FlowDoc body.
function paragraphTexts(doc: FlowDoc): Array<string> {
  return doc.body
    .filter((el) => el.kind === 'paragraph')
    .map((el) => el.paragraph.runs.map((r) => r.text).join(''));
}

// The runs (text + resolved properties) of the document's first paragraph.
function firstParagraphRuns(doc: FlowDoc) {
  const p = doc.body.find((el) => el.kind === 'paragraph');
  return p ? p.paragraph.runs : [];
}

// The paragraph elements of the FlowDoc body.
function paragraphs(doc: FlowDoc) {
  return doc.body.filter((el) => el.kind === 'paragraph');
}

// The concatenated text of a table cell's paragraphs.
function cellText(content: ReadonlyArray<BodyElement>): string {
  return content
    .flatMap((el) => (el.kind === 'paragraph' ? el.paragraph.runs.map((r) => r.text) : []))
    .join('');
}

describe('doc reader (DOC-1)', () => {
  it('reads UTF-16 piece text and splits paragraphs at the CR mark', () => {
    const doc = readDoc(buildDoc([{ text: 'Hello world\rSecond line\r', compressed: false }])).doc;
    expect(paragraphTexts(doc)).toEqual(['Hello world', 'Second line']);
  });

  it('decodes a compressed (Windows-1252) piece, not Latin-1', () => {
    // 0x97 → em dash (—), 0x92 → right single quote (’) — the cp1252 high range.
    const doc = readDoc(buildDoc([{ text: 'Café — déjà’\r', compressed: true }])).doc;
    expect(paragraphTexts(doc)).toEqual(['Café — déjà’']);
  });

  it('concatenates a compressed and an uncompressed piece in CP order', () => {
    const doc = readDoc(
      buildDoc([
        { text: 'ASCII part ', compressed: true },
        { text: 'Unicode π\r', compressed: false }, // π is only encodable as UTF-16
      ]),
    ).doc;
    expect(paragraphTexts(doc)).toEqual(['ASCII part Unicode π']);
  });

  it('reads the piece table from 0Table when fWhichTblStm is clear', () => {
    const doc = readDoc(
      buildDoc([{ text: 'Zero table stream\r', compressed: false }], { whichTable: '0Table' }),
    ).doc;
    expect(paragraphTexts(doc)).toEqual(['Zero table stream']);
  });

  it('drops control characters and turns tabs into spaces', () => {
    const doc = readDoc(buildDoc([{ text: 'a\tbc\r', compressed: false }])).doc;
    expect(paragraphTexts(doc)).toEqual(['a bc']); // tab → space
  });

  it('reports the formatting loss and reads no text from an encrypted file', () => {
    const result = readDoc(
      buildDoc([{ text: 'secret\r', compressed: false }], { encrypted: true }),
    );
    expect(paragraphTexts(result.doc)).toEqual(['​']); // one empty paragraph (ZWSP)
    expect(result.losses[0]).toMatchObject({ severity: 'dropped', feature: 'text' });
    expect(result.losses[0]?.detail).toContain('encrypted');
  });

  it('records a degraded text loss for a normal .doc', () => {
    const result = readDoc(buildDoc([{ text: 'plain\r', compressed: false }]));
    expect(result.losses[0]).toMatchObject({ severity: 'degraded', feature: 'text' });
  });

  it('sniffs a .doc but not a docx zip or random bytes', () => {
    expect(docReader.sniff(buildDoc([{ text: 'x\r', compressed: false }]))).toBe(true);
    expect(docReader.sniff(new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 4]))).toBe(false);
    expect(docReader.sniff(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]))).toBe(false);
  });

  it('exposes the .doc format id through Ream', () => {
    expect(Ream.parse(buildDoc([{ text: 'hi\r', compressed: false }])).format).toBe('doc');
  });

  it('reads bold/italic run formatting and splits runs at the boundaries', () => {
    // "Bold " bold, "then " plain, "italic" italic, then the paragraph mark.
    const doc = readDoc(
      buildDoc([{ text: 'Bold then italic\r', compressed: false }], {
        formatRuns: [
          { length: 5, bold: true },
          { length: 5 },
          { length: 6, italic: true },
          { length: 1 },
        ],
      }),
    ).doc;
    const runs = firstParagraphRuns(doc);
    expect(runs.map((r) => r.text)).toEqual(['Bold ', 'then ', 'italic']);
    // Properties are post-cascade, so unset toggles resolve to false (not absent).
    expect(runs[0]?.properties.bold).toBe(true);
    expect(runs[0]?.properties.italic).toBeFalsy();
    expect(runs[1]?.properties.bold).toBeFalsy();
    expect(runs[1]?.properties.italic).toBeFalsy();
    expect(runs[2]?.properties.italic).toBe(true);
    expect(runs[2]?.properties.bold).toBeFalsy();
  });

  it('reads font size (half-points) and underline from the CHPX', () => {
    const doc = readDoc(
      buildDoc([{ text: 'x\r', compressed: false }], {
        formatRuns: [{ length: 1, sizeHalfPts: 48, underlineKul: 1 }, { length: 1 }],
      }),
    ).doc;
    const runs = firstParagraphRuns(doc);
    expect(runs).toHaveLength(1);
    expect(runs[0]?.properties.fontSizePt).toBe(24); // 48 half-points
    expect(runs[0]?.properties.underline).toBe('single');
  });

  it('reads paragraph alignment from the PAPX', () => {
    const doc = readDoc(
      buildDoc([{ text: 'Centered\r', compressed: false }], {
        paraRuns: [{ length: 9, jc: 1 }], // jc 1 = center
      }),
    ).doc;
    expect(paragraphs(doc)[0]?.paragraph.properties.alignment).toBe('center');
  });

  it('reads paragraph indents and spacing (twips → points) from the PAPX', () => {
    const doc = readDoc(
      buildDoc([{ text: 'Indented\r', compressed: false }], {
        paraRuns: [
          {
            length: 9,
            indentLeftTwips: 720, // 36pt
            indentFirstTwips: -360, // -18pt (hanging)
            spaceBeforeTwips: 240, // 12pt
            spaceAfterTwips: 120, // 6pt
          },
        ],
      }),
    ).doc;
    const props = paragraphs(doc)[0]?.paragraph.properties;
    expect(props?.indentLeft).toBe(36);
    expect(props?.indentFirstLine).toBe(-18);
    expect(props?.spacingBefore).toBe(12);
    expect(props?.spacingAfter).toBe(6);
  });

  it('applies CHPX run formatting and PAPX paragraph formatting together', () => {
    const doc = readDoc(
      buildDoc([{ text: 'Title\r', compressed: false }], {
        formatRuns: [{ length: 6, bold: true }],
        paraRuns: [{ length: 6, jc: 1 }],
      }),
    ).doc;
    expect(firstParagraphRuns(doc)[0]?.properties.bold).toBe(true);
    expect(paragraphs(doc)[0]?.paragraph.properties.alignment).toBe('center');
  });

  it('gives each paragraph its own PAPX alignment', () => {
    const doc = readDoc(
      buildDoc([{ text: 'Left\rRight\r', compressed: false }], {
        paraRuns: [
          { length: 5, jc: 0 }, // left (default)
          { length: 6, jc: 2 }, // right
        ],
      }),
    ).doc;
    const ps = paragraphs(doc);
    expect(ps[0]?.paragraph.properties.alignment).not.toBe('right');
    expect(ps[1]?.paragraph.properties.alignment).toBe('right');
  });

  it('groups in-table paragraphs into a 2×2 Table (cell mark 0x07, TTP row end)', () => {
    const doc = readDoc(
      buildDoc([{ text: `A${CM}B${CM}C${CM}D${CM}`, compressed: false }], {
        paraRuns: [
          { length: 2, inTable: true }, // "A" + cell mark
          { length: 2, inTable: true, rowEnd: true }, // "B" + cell mark, ends row 1
          { length: 2, inTable: true }, // "C"
          { length: 2, inTable: true, rowEnd: true }, // "D", ends row 2
        ],
      }),
    ).doc;
    const tableEl = doc.body.find((el) => el.kind === 'table');
    expect(tableEl).toBeDefined();
    const t = tableEl?.table;
    expect(t?.rows).toHaveLength(2);
    expect(t?.rows[0]?.cells).toHaveLength(2);
    expect(cellText(t!.rows[0]!.cells[0]!.content)).toBe('A');
    expect(cellText(t!.rows[0]!.cells[1]!.content)).toBe('B');
    expect(cellText(t!.rows[1]!.cells[0]!.content)).toBe('C');
    expect(cellText(t!.rows[1]!.cells[1]!.content)).toBe('D');
  });

  it('keeps the paragraphs before and after a table', () => {
    const doc = readDoc(
      buildDoc([{ text: `Before\rA${CM}B${CM}After\r`, compressed: false }], {
        paraRuns: [
          { length: 7 }, // "Before" + CR — not in a table
          { length: 2, inTable: true }, // "A"
          { length: 2, inTable: true, rowEnd: true }, // "B", ends the row
          { length: 6 }, // "After" + CR — not in a table
        ],
      }),
    ).doc;
    expect(doc.body.map((el) => el.kind)).toEqual(['paragraph', 'table', 'paragraph']);
  });

  it('reads an inline picture into an image block (DOC-5)', () => {
    const doc = readDoc(
      buildDoc([{ text: `${PIC}\r`, compressed: false }], {
        formatRuns: [
          { length: 1, picOffset: 0 }, // the picture placeholder char
          { length: 1 }, // the CR
        ],
        data: buildPicf(PNG_1x1, 1440, 720), // 72pt × 36pt
      }),
    ).doc;
    const img = doc.body.find((el) => el.kind === 'image');
    expect(img).toBeDefined();
    expect(doc.resources.get(img!.image.resource!)).toEqual(PNG_1x1);
    expect(img!.image.width).toBe(72); // 1440 twips / 20
    expect(img!.image.height).toBe(36); // 720 twips / 20
  });

  it('drops a picture whose Data stream is absent, without failing', () => {
    const doc = readDoc(
      buildDoc([{ text: `${PIC}Body\r`, compressed: false }], {
        formatRuns: [{ length: 1, picOffset: 0 }, { length: 5 }],
        // no `data` — the picture cannot be resolved
      }),
    ).doc;
    expect(doc.body.find((el) => el.kind === 'image')).toBeUndefined();
    expect(paragraphTexts(doc)).toEqual(['Body']);
  });

  it('renders a .doc with an image to a valid PDF', async () => {
    const pdf = await Ream.parse(
      buildDoc([{ text: `${PIC}\r`, compressed: false }], {
        formatRuns: [{ length: 1, picOffset: 0 }, { length: 1 }],
        data: buildPicf(PNG_1x1, 1440, 1440),
      }),
    ).convert('pdf', {
      fonts: {
        regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
        bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
      },
    });
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });

  it('renders a .doc with a table to a valid PDF', async () => {
    const pdf = await Ream.parse(
      buildDoc([{ text: `A${CM}B${CM}C${CM}D${CM}`, compressed: false }], {
        paraRuns: [
          { length: 2, inTable: true },
          { length: 2, inTable: true, rowEnd: true },
          { length: 2, inTable: true },
          { length: 2, inTable: true, rowEnd: true },
        ],
      }),
    ).convert('pdf', {
      fonts: {
        regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
        bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
      },
    });
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });

  it('renders a formatted .doc to a valid PDF', async () => {
    const pdf = await Ream.parse(
      buildDoc([{ text: 'Bold title\rbody text\r', compressed: false }], {
        formatRuns: [{ length: 11, bold: true, sizeHalfPts: 36 }, { length: 10 }],
        paraRuns: [
          { length: 11, jc: 1 }, // centered title
          { length: 10, indentLeftTwips: 360 },
        ],
      }),
    ).convert('pdf', {
      fonts: {
        regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
        bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
      },
    });
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });

  it('renders a .doc to a valid PDF', async () => {
    const pdf = await Ream.parse(
      buildDoc([{ text: 'Hello from a legacy doc\rSecond paragraph\r', compressed: false }]),
    ).convert('pdf', {
      fonts: {
        regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
        bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
      },
    });
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });
});
