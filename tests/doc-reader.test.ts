// DOC-1 — legacy `.doc` text extraction. A Word 97–2003 binary file is a CFB
// holding a `WordDocument` stream; its text lives in pieces (the piece table /
// CLX) that are either 16-bit Unicode or 8-bit Windows-1252. These build a
// fixture `.doc` and assert the reader walks the FIB → piece table → pieces back
// to the original text, splits it into paragraphs, and renders to PDF.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDoc } from './fixtures/build-doc';
import type { FlowDoc } from '@/core/ir/flow';

import { docReader, readDoc } from '@/word/doc/doc-reader';
import { Ream } from '@/core/converter/ream';

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

  it('renders a formatted .doc to a valid PDF', async () => {
    const pdf = await Ream.parse(
      buildDoc([{ text: 'Bold title\rbody text\r', compressed: false }], {
        formatRuns: [{ length: 11, bold: true, sizeHalfPts: 36 }, { length: 10 }],
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
