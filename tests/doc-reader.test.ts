// DOC-1 — legacy `.doc` text extraction. A Word 97–2003 binary file is a CFB
// holding a `WordDocument` stream; its text lives in pieces (the piece table /
// CLX) that are either 16-bit Unicode or 8-bit Windows-1252. These build a
// fixture `.doc` and assert the reader walks the FIB → piece table → pieces back
// to the original text, splits it into paragraphs, and renders to PDF.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDoc, buildPicf } from './fixtures/build-doc';
import type { DocSepMargins } from './fixtures/build-doc';
import type { BodyElement } from '@/core/document-model';
import type { FlowDoc } from '@/core/ir/flow';
import type { TextLineItem } from '@/layout/page-doc';

import { docReader, readDoc } from '@/word/doc/doc-reader';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';

const CM = String.fromCharCode(0x07); // table cell mark
const PIC = String.fromCharCode(0x01); // picture placeholder char
const F13 = String.fromCharCode(0x13); // field begin
const F14 = String.fromCharCode(0x14); // field separator
const F15 = String.fromCharCode(0x15); // field end
const PB = String.fromCharCode(0x0c); // page break — or, closing a section, its section mark
const COL = String.fromCharCode(0x0e); // column break
const ZWSP = '\u200b'; // the text of an empty paragraph
const FONTS = {
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

// The text of each line on each page, laid out the way the PDF is.
function pageTexts(doc: FlowDoc): Array<Array<string>> {
  const laid = layoutStyledDocument(doc.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(doc),
  });
  return laid.pages.map((page) =>
    page.commands
      .filter((c): c is TextLineItem => c.type === 'line')
      .map((l) => l.line.tokens.map((t) => (t.kind === 'text' ? t.text : '')).join('')),
  );
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

  it('reads the section page size from the SEP (landscape Letter, F4)', () => {
    // sprmSXaPage/sprmSYaPage in the first section's Sepx → the FlowDoc section.
    // Landscape Letter: width 15840 twips (11"), height 12240 (8.5") → 792×612pt.
    const doc = readDoc(
      buildDoc([{ text: 'Hi\r', compressed: true }], {
        pageSizeTwips: { width: 15840, height: 12240 },
      }),
    ).doc;
    expect(Math.round(doc.section!.pageSize!.width)).toBe(792);
    expect(Math.round(doc.section!.pageSize!.height)).toBe(612);
    expect(doc.section!.pageSize!.orientation).toBe('landscape');
  });

  it('defaults to US Letter when the section gives no page size', () => {
    const doc = readDoc(buildDoc([{ text: 'Hi\r', compressed: true }])).doc;
    expect(Math.round(doc.section!.pageSize!.width)).toBe(612);
    expect(Math.round(doc.section!.pageSize!.height)).toBe(792);
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

  it('reads a line break (0x0B) as the end of a line within the paragraph', () => {
    // Word: "Latin" and "Second" on two lines of one paragraph; two breaks leave
    // an empty line between them.
    const LB = String.fromCharCode(0x0b);
    const one = readDoc(buildDoc([{ text: `Latin${LB}Second\r`, compressed: false }])).doc;
    expect(paragraphTexts(one)).toEqual(['Latin\nSecond']);
    expect(pageTexts(one)).toEqual([['Latin', 'Second']]);
    const two = readDoc(buildDoc([{ text: `Latin${LB}${LB}Second\r`, compressed: false }])).doc;
    expect(pageTexts(two)).toEqual([['Latin', '', 'Second']]);
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

  it('reads widow control from the PAPX, on where it says nothing', () => {
    // [MS-DOC] sprmPFWidowControl — Word's Normal sets it, and a paragraph that
    // turns it off says so with a zero.
    const doc = readDoc(
      buildDoc([{ text: 'Off\rOn\r', compressed: false }], {
        paraRuns: [{ length: 4, widowControl: false }, { length: 3 }],
      }),
    ).doc;
    const [off, on] = paragraphs(doc);
    expect(off?.paragraph.properties.widowControl).toBe(false);
    expect(on?.paragraph.properties.widowControl).toBe(true);
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

  it('reads per-cell borders from the TC80 array (DOC-11)', () => {
    const doc = readDoc(
      buildDoc([{ text: `A${CM}B${CM}`, compressed: false }], {
        paraRuns: [
          { length: 2, inTable: true }, // cell A
          {
            length: 2,
            inTable: true,
            rowEnd: true,
            cellEdgesTwips: [0, 1440, 2880], // 2 cells
            // cell A: a 2pt (16/8) red (ico 6) single top border; cell B: no borders.
            cellTc: [{ borders: { top: { widthEighthPt: 16, brcType: 1, ico: 6 } } }, {}],
          },
        ],
      }),
    ).doc;
    const t = doc.body.find((el) => el.kind === 'table')?.table;
    expect(t?.rows[0]?.cells[0]?.properties.borders?.top).toEqual({
      style: 'single',
      width: 2,
      colorHex: 'FF0000',
    });
    expect(t?.rows[0]?.cells[1]?.properties.borders).toBeUndefined();
  });

  it('reads per-cell background shading from sprmTDefTableShd (DOC-11)', () => {
    const doc = readDoc(
      buildDoc([{ text: `A${CM}B${CM}`, compressed: false }], {
        paraRuns: [
          { length: 2, inTable: true },
          {
            length: 2,
            inTable: true,
            rowEnd: true,
            cellEdgesTwips: [0, 1440, 2880],
            cellShadings: ['FFFF00', undefined], // cell A yellow fill, cell B no fill
          },
        ],
      }),
    ).doc;
    const t = doc.body.find((el) => el.kind === 'table')?.table;
    expect(t?.rows[0]?.cells[0]?.properties.shading).toEqual({ colorHex: 'FFFF00' });
    expect(t?.rows[0]?.cells[1]?.properties.shading).toBeUndefined();
  });

  it('resolves a vertical merge into start/end cell roles (DOC-11)', () => {
    const doc = readDoc(
      buildDoc([{ text: `A${CM}B${CM}`, compressed: false }], {
        paraRuns: [
          {
            length: 2,
            inTable: true,
            rowEnd: true,
            cellEdgesTwips: [0, 1440],
            cellTc: [{ vMerge: 'restart' }],
          },
          {
            length: 2,
            inTable: true,
            rowEnd: true,
            cellEdgesTwips: [0, 1440],
            cellTc: [{ vMerge: 'continue' }],
          },
        ],
      }),
    ).doc;
    const t = doc.body.find((el) => el.kind === 'table')?.table;
    expect(t?.rows[0]?.cells[0]?.properties.merge).toBe('start');
    expect(t?.rows[1]?.cells[0]?.properties.merge).toBe('end');
  });

  it('suppresses field codes and keeps the cached result (DOC-6)', () => {
    // "Page {PAGE→1} of {NUMPAGES→3}" — the codes are dropped, the results kept.
    const doc = readDoc(
      buildDoc([
        {
          text: `Page ${F13} PAGE ${F14}1${F15} of ${F13}NUMPAGES${F14}3${F15}\r`,
          compressed: false,
        },
      ]),
    ).doc;
    expect(paragraphTexts(doc)).toEqual(['Page 1 of 3']);
  });

  it('suppresses the whole code of a nested field', () => {
    // An outer field whose code contains a nested field; only the outer result shows.
    const doc = readDoc(
      buildDoc([{ text: `${F13}IF ${F13}A${F14}B${F15} ${F14}RESULT${F15}\r`, compressed: false }]),
    ).doc;
    expect(paragraphTexts(doc)).toEqual(['RESULT']);
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

  it('reads cell widths from sprmTDefTable and stays aligned past it (DOC-7)', () => {
    const doc = readDoc(
      buildDoc([{ text: `A${CM}B${CM}`, compressed: false }], {
        paraRuns: [
          { length: 2, inTable: true },
          // The TTP carries the long sprmTDefTable before the table flags.
          { length: 2, inTable: true, rowEnd: true, cellEdgesTwips: [0, 2880, 4320] },
        ],
      }),
    ).doc;
    const t = doc.body.find((el) => el.kind === 'table')?.table;
    expect(t).toBeDefined();
    // The row was still detected — sprmPFTtp is read after the 2-byte-length sprm.
    expect(t!.rows).toHaveLength(1);
    expect(t!.rows[0]!.cells).toHaveLength(2);
    // 2880 and 1440 twips → 144 and 72 points.
    expect(t!.grid).toEqual([144, 72]);
    expect(t!.rows[0]!.cells[0]!.properties.width).toBe(144);
    expect(t!.rows[0]!.cells[1]!.properties.width).toBe(72);
  });

  it('reads the default header and footer stories (DOC-8)', () => {
    const doc = readDoc(
      buildDoc([{ text: 'Body\r', compressed: false }], {
        headerFooter: { defaultHeader: 'My header', defaultFooter: 'Page footer' },
      }),
    ).doc;
    expect(doc.section?.headers.map((h) => h.type)).toEqual(['default']);
    expect(doc.section?.footers.map((f) => f.type)).toEqual(['default']);
    const hdrId = doc.section!.headers[0]!.relationshipId;
    const ftrId = doc.section!.footers[0]!.relationshipId;
    expect(cellText(doc.headersFooters!.get(hdrId)!)).toBe('My header');
    expect(cellText(doc.headersFooters!.get(ftrId)!)).toBe('Page footer');
    expect(paragraphTexts(doc)).toEqual(['Body']); // the body is unaffected
  });

  it('classifies a first-page header as type "first"', () => {
    const doc = readDoc(
      buildDoc([{ text: 'Body\r', compressed: false }], {
        headerFooter: { firstHeader: 'Cover header' },
      }),
    ).doc;
    expect(doc.section?.headers.map((h) => h.type)).toEqual(['first']);
    const id = doc.section!.headers[0]!.relationshipId;
    expect(cellText(doc.headersFooters!.get(id)!)).toBe('Cover header');
  });

  it('has no header/footer references when the doc has none', () => {
    const doc = readDoc(buildDoc([{ text: 'Body\r', compressed: false }])).doc;
    expect(doc.section?.headers).toEqual([]);
    expect(doc.headersFooters).toBeUndefined();
  });

  it('renders list items with a bullet marker and per-level indent (DOC-9)', () => {
    const doc = readDoc(
      buildDoc([{ text: 'First\rSecond\r', compressed: false }], {
        paraRuns: [
          { length: 6, listIlfo: 1, listIlvl: 0 }, // "First" — list item, level 0
          { length: 7, listIlfo: 1, listIlvl: 1 }, // "Second" — list item, level 1
        ],
      }),
    ).doc;
    const ps = paragraphs(doc);
    expect(ps[0]?.paragraph.runs.map((r) => r.text)).toEqual(['• ', 'First']);
    expect(ps[0]?.paragraph.properties.indentLeft).toBe(18); // level 0 → 18pt
    expect(ps[1]?.paragraph.runs.map((r) => r.text)).toEqual(['◦ ', 'Second']);
    expect(ps[1]?.paragraph.properties.indentLeft).toBe(36); // level 1 → 36pt
  });

  it('renders an arabic numbered list with a running counter (DOC-10)', () => {
    const doc = readDoc(
      buildDoc([{ text: 'A\rB\rC\r', compressed: false }], {
        paraRuns: [
          { length: 2, listIlfo: 1, listIlvl: 0 },
          { length: 2, listIlfo: 1, listIlvl: 0 },
          { length: 2, listIlfo: 1, listIlvl: 0 },
        ],
        // lsid 100: a simple list, level 0 = arabic, template "%0." → 0x0000 + '.'.
        lists: {
          lfos: [100],
          lstfs: [
            { lsid: 100, simple: true, levels: [{ nfc: 0, iStartAt: 1, xst: [0x0000, 0x2e] }] },
          ],
        },
      }),
    ).doc;
    const ps = paragraphs(doc);
    expect(ps[0]?.paragraph.runs.map((r) => r.text)).toEqual(['1. ', 'A']);
    expect(ps[1]?.paragraph.runs.map((r) => r.text)).toEqual(['2. ', 'B']);
    expect(ps[2]?.paragraph.runs.map((r) => r.text)).toEqual(['3. ', 'C']);
  });

  it('renders a lower-letter list (nfc 4) (DOC-10)', () => {
    const doc = readDoc(
      buildDoc([{ text: 'A\rB\r', compressed: false }], {
        paraRuns: [
          { length: 2, listIlfo: 1, listIlvl: 0 },
          { length: 2, listIlfo: 1, listIlvl: 0 },
        ],
        lists: {
          lfos: [7],
          lstfs: [
            { lsid: 7, simple: true, levels: [{ nfc: 4, iStartAt: 1, xst: [0x0000, 0x29] }] },
          ],
        },
      }),
    ).doc;
    const ps = paragraphs(doc);
    expect(ps[0]?.paragraph.runs[0]?.text).toBe('a) ');
    expect(ps[1]?.paragraph.runs[0]?.text).toBe('b) ');
  });

  it('renders a multi-level "1.1." numbered list (DOC-10)', () => {
    const doc = readDoc(
      buildDoc([{ text: 'A\rB\r', compressed: false }], {
        paraRuns: [
          { length: 2, listIlfo: 1, listIlvl: 0 },
          { length: 2, listIlfo: 1, listIlvl: 1 },
        ],
        lists: {
          lfos: [5],
          lstfs: [
            {
              lsid: 5,
              simple: false,
              levels: [
                { nfc: 0, iStartAt: 1, xst: [0x0000, 0x2e] }, // "%0."
                { nfc: 0, iStartAt: 1, xst: [0x0000, 0x2e, 0x0001, 0x2e] }, // "%0.%1."
              ],
            },
          ],
        },
      }),
    ).doc;
    const ps = paragraphs(doc);
    expect(ps[0]?.paragraph.runs[0]?.text).toBe('1. ');
    expect(ps[1]?.paragraph.runs[0]?.text).toBe('1.1. ');
  });

  it('keeps a bullet (nfc 23) as its glyph, not a number (DOC-10)', () => {
    const doc = readDoc(
      buildDoc([{ text: 'A\r', compressed: false }], {
        paraRuns: [{ length: 2, listIlfo: 1, listIlvl: 0 }],
        lists: {
          lfos: [9],
          lstfs: [{ lsid: 9, simple: true, levels: [{ nfc: 23, xst: [0x2022] }] }], // '•'
        },
      }),
    ).doc;
    expect(paragraphs(doc)[0]?.paragraph.runs[0]?.text).toBe('• ');
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

describe('doc reader: page and column breaks', () => {
  // [MS-DOC] §2.8.26 — 0x0C is a manual page break, except as the last character
  // of a section (each but the last), where it is that section's mark. The pages
  // are Word for Mac's, counted on .doc files that carry these very characters.

  it('reads a page break as a run of its own, where it stands among the words', () => {
    const doc = readDoc(buildDoc([{ text: `Latin${PB}Second\r`, compressed: false }])).doc;
    expect(firstParagraphRuns(doc).map((r) => [r.text, r.pageBreak === true])).toEqual([
      ['Latin', false],
      ['', true],
      ['Second', false],
    ]);
    // Word: "Latin" on page 1, "Second" at the top of page 2.
    expect(pageTexts(doc)).toEqual([['Latin'], ['Second']]);
  });

  it('breaks the page before a paragraph that opens with one, and twice for two', () => {
    // Word: "Second" at the top of page 2; with two breaks in a row, three pages
    // and nothing but an empty line on the second.
    const opening = readDoc(buildDoc([{ text: `one\r${PB}Second\r`, compressed: false }])).doc;
    expect(pageTexts(opening)).toEqual([['one'], ['Second']]);
    const twice = readDoc(buildDoc([{ text: `Latin${PB}${PB}Second\r`, compressed: false }])).doc;
    expect(pageTexts(twice)).toEqual([['Latin'], [''], ['Second']]);
  });

  it('sets the mark of a paragraph that ends in a break at the head of the next page', () => {
    // Word's .doc, whatever Copts.fSplitPgBreakAndParaMark says: "one", a break,
    // then "two" on the second line of page 2, under the mark's line; the same
    // after a column break; and a last paragraph of nothing but a break makes
    // a page of its own.
    const pages = (text: string) => pageTexts(readDoc(buildDoc([{ text, compressed: false }])).doc);
    expect(pages(`one${PB}\rtwo\r`)).toEqual([['one'], [ZWSP, 'two']]);
    expect(pages(`one${COL}\rtwo\r`)).toEqual([['one'], [ZWSP, 'two']]);
    expect(pages(`one\r${PB}\r`)).toEqual([['one'], [ZWSP]]);
  });

  it('reads a column break, which starts the next page in a section of one column', () => {
    const doc = readDoc(buildDoc([{ text: `Latin${COL}Second\r`, compressed: false }])).doc;
    expect(firstParagraphRuns(doc).map((r) => [r.text, r.columnBreak === true])).toEqual([
      ['Latin', false],
      ['\n', true],
      ['Second', false],
    ]);
    // Word: "Second" at the top of page 2.
    expect(pageTexts(doc)).toEqual([['Latin'], ['Second']]);
  });

  it('ends the paragraph at a section mark, and breaks no page there', () => {
    // "one", the first section's mark, then "two" in a section that runs on
    // (sprmSBkc 0). Word: one page, "two" on the line under "one".
    const doc = readDoc(
      buildDoc([{ text: `one${PB}two\r`, compressed: false }], {
        sections: [{ length: 4 }, { length: 4, bkc: 0 }],
      }),
    ).doc;
    expect(paragraphTexts(doc)).toEqual(['one', 'two']);
    expect(paragraphs(doc).flatMap((el) => el.paragraph.runs.filter((r) => r.pageBreak))).toEqual(
      [],
    );
    expect(pageTexts(doc)).toEqual([['one', 'two']]);
  });

  it('reads a section mark alone after a paragraph mark as an empty paragraph', () => {
    const doc = readDoc(
      buildDoc([{ text: `one\r${PB}two\r`, compressed: false }], {
        sections: [{ length: 5 }, { length: 4, bkc: 0 }],
      }),
    ).doc;
    expect(paragraphTexts(doc)).toEqual(['one', ZWSP, 'two']);
  });

  it('starts a new page at a paragraph marked to break before, unless it is the first', () => {
    // sprmPFPageBreakBefore — how LibreOffice writes every page break into a .doc.
    // Word: "Second" at the top of page 2; on the first paragraph, one page.
    const doc = readDoc(
      buildDoc([{ text: 'one\rSecond\r', compressed: false }], {
        paraRuns: [{ length: 4 }, { length: 7, pageBreakBefore: true }],
      }),
    ).doc;
    expect(paragraphs(doc).map((el) => el.paragraph.properties.pageBreakBefore === true)).toEqual([
      false,
      true,
    ]);
    expect(pageTexts(doc)).toEqual([['one'], ['Second']]);
    const first = readDoc(
      buildDoc([{ text: 'one\rtwo\r', compressed: false }], {
        paraRuns: [{ length: 4, pageBreakBefore: true }, { length: 4 }],
      }),
    ).doc;
    expect(pageTexts(first)).toEqual([['one', 'two']]);
  });
});

describe('doc reader: sections', () => {
  // [MS-DOC] §2.8.26 — each section runs up to its section mark, and the next one
  // starts where its own sprmSBkc says. Pages as Word for Mac sets the same
  // sections in .doc files LibreOffice wrote.
  const sectioned = (
    text: string,
    sections: ReadonlyArray<{
      length: number;
      bkc?: number;
      pageSizeTwips?: { width: number; height: number };
      marginsTwips?: DocSepMargins;
    }>,
  ): FlowDoc => readDoc(buildDoc([{ text, compressed: false }], { sections })).doc;

  it('starts a section on a new page when its SEP says nothing', () => {
    // Word: "two" at the top of page 2.
    const doc = sectioned(`one${PB}two\r`, [{ length: 4 }, { length: 4 }]);
    expect(doc.sections.map((s) => s.endIndex)).toEqual([1, 2]);
    expect(pageTexts(doc)).toEqual([['one'], ['two']]);
  });

  it('starts it on the page in hand, on an odd page or on an even one, as bkc says', () => {
    // Word: "two" under "one" (0); on page 3 after a blank page 2 (4, odd); on
    // page 2 (3, even).
    const at = (bkc: number) =>
      pageTexts(sectioned(`one${PB}two\r`, [{ length: 4 }, { length: 4, bkc }]));
    expect(at(0)).toEqual([['one', 'two']]);
    expect(at(4)).toEqual([['one'], [], ['two']]);
    expect(at(3)).toEqual([['one'], ['two']]);
  });

  it('lays out three sections each as its own bkc says', () => {
    // Word: "one" on page 1; "two" and "three" on page 2, the third section
    // running on from the second.
    const doc = sectioned(`one${PB}two${PB}three\r`, [
      { length: 4 },
      { length: 4 },
      { length: 6, bkc: 0 },
    ]);
    expect(pageTexts(doc)).toEqual([['one'], ['two', 'three']]);
  });

  it('gives each section its own page size', () => {
    // Word: page 1 is 612×792, page 2 (the landscape section) 792×612.
    const doc = sectioned(`one${PB}two\r`, [
      { length: 4 },
      { length: 4, pageSizeTwips: { width: 15840, height: 12240 } },
    ]);
    const laid = layoutStyledDocument(doc.body, {
      registry: FontRegistry.fromBytes(FONTS),
      ...flowRenderOptions(doc),
    });
    expect(laid.pages.map((p) => [p.width, p.height])).toEqual([
      [612, 792],
      [792, 612],
    ]);
  });

  it('sets each section within its own margins', () => {
    // Word: "one" at x 162.2 (a two-inch margin and a quarter-inch gutter), its
    // line's top at 108.4; "two", in the next section, at 36.4 and 36.4. A top
    // margin stored negative is the same size.
    const doc = sectioned(`one${PB}two\r`, [
      { length: 4, marginsTwips: { left: 2880, right: 720, top: 2160, bottom: 1440, gutter: 360 } },
      { length: 4, marginsTwips: { left: 720, right: 2880, top: -720, bottom: 720 } },
    ]);
    expect(doc.sections.map((s) => s.properties.margins)).toEqual([
      { top: 108, right: 36, bottom: 72, left: 144, gutter: 18 },
      { top: 36, right: 144, bottom: 36, left: 36 },
    ]);
    const laid = layoutStyledDocument(doc.body, {
      registry: FontRegistry.fromBytes(FONTS),
      ...flowRenderOptions(doc),
    });
    const first = laid.pages.map((p) => p.commands.find((c) => c.type === 'line')!);
    expect(first.map((l) => l.originX)).toEqual([162, 36]);
    expect(first[0]!.baselineY - first[1]!.baselineY).toBeCloseTo(72, 3);
  });

  it('gives each section its own header and footer, or the one before', () => {
    // Word: "Head A" and "Foot A" on page 1; "Head B" on page 2, over the
    // footer the second section takes from the first ("same as previous", which
    // Word writes as a story of no characters).
    const doc = readDoc(
      buildDoc([{ text: `one${PB}two\r`, compressed: false }], {
        sections: [{ length: 4 }, { length: 4 }],
        headerFooter: [
          { defaultHeader: 'Head A', defaultFooter: 'Foot A' },
          { defaultHeader: 'Head B' },
        ],
      }),
    ).doc;
    expect(pageTexts(doc)).toEqual([
      ['Head A', 'one', 'Foot A'],
      ['Head B', 'two', 'Foot A'],
    ]);
  });

  it('shows no line for the paragraph mark that closes every story', () => {
    // LibreOffice writes a one-line header as "Head", CR, CR; Word shows one line.
    const doc = readDoc(
      buildDoc([{ text: 'one\r', compressed: false }], {
        headerFooter: { defaultHeader: 'Head\r' },
      }),
    ).doc;
    expect(pageTexts(doc)).toEqual([['Head', 'one']]);
  });

  it('shows the first page its own header when the section asks', () => {
    // sprmSFTitlePage. Word: "Cover" over page 1 and no footer under it, "Head"
    // and "Foot" on page 2.
    const doc = readDoc(
      buildDoc([{ text: `one${PB}two\r`, compressed: false }], {
        sections: [{ length: 8, titlePage: true }],
        headerFooter: { firstHeader: 'Cover', defaultHeader: 'Head', defaultFooter: 'Foot' },
      }),
    ).doc;
    expect(pageTexts(doc)).toEqual([
      ['Cover', 'one'],
      ['Head', 'two', 'Foot'],
    ]);
  });

  it('shows the even pages their own header when the document has facing pages', () => {
    // Dop.fFacingPages. Word: "Odd", "Even", "Odd" over three pages.
    const bands = { defaultHeader: 'Odd', evenHeader: 'Even' };
    const text = `one${PB}two${PB}three\r`;
    const facing = readDoc(
      buildDoc([{ text, compressed: false }], { headerFooter: bands, facingPages: true }),
    ).doc;
    expect(pageTexts(facing).map((lines) => lines[0])).toEqual(['Odd', 'Even', 'Odd']);
    const single = readDoc(buildDoc([{ text, compressed: false }], { headerFooter: bands })).doc;
    expect(pageTexts(single).map((lines) => lines[0])).toEqual(['Odd', 'Odd', 'Odd']);
  });

  it("numbers each page in a footer's PAGE and NUMPAGES fields, not as Word cached them", () => {
    // The codes LibreOffice writes; Word saved "9" and "9" as the results.
    // Word: "Page 1 of 3", "Page 2 of 3", "Page 3 of 3".
    const footer = `Page ${F13} PAGE ${F14}9${F15} of ${F13} NUMPAGES \\* MERGEFORMAT ${F14}9${F15}`;
    const doc = readDoc(
      buildDoc([{ text: `one${PB}two${PB}three\r`, compressed: false }], {
        headerFooter: { defaultFooter: footer },
      }),
    ).doc;
    expect(pageTexts(doc).map((lines) => lines[lines.length - 1])).toEqual([
      'Page 1 of 3',
      'Page 2 of 3',
      'Page 3 of 3',
    ]);
    const id = doc.section!.footers[0]!.relationshipId;
    const [band] = doc.headersFooters!.get(id)!;
    const runs = band?.kind === 'paragraph' ? band.paragraph.runs : [];
    expect(runs.map((r) => [r.text, r.field])).toEqual([
      ['Page ', undefined],
      ['9', 'PAGE'],
      [' of ', undefined],
      ['9', 'NUMPAGES'],
    ]);
  });

  it("numbers a section's pages again where its SEP starts the count again", () => {
    // LibreOffice writes `w:pgNumType w:start="1" w:fmt="lowerRoman"` on the
    // second section as sprmSFPgnRestart, sprmSPgnStart97 1 and sprmSNfcPgn 2.
    // Word: "p1", "p2", then "pi", "pii".
    const page = `p${F13} PAGE ${F14}9${F15}`;
    const doc = readDoc(
      buildDoc([{ text: `one${PB}two${PB}three${PB}four\r`, compressed: false }], {
        sections: [{ length: 8 }, { length: 11, pageNumbers: { nfc: 2, restartAt: 1 } }],
        headerFooter: [{ defaultFooter: page }, { defaultFooter: page }],
      }),
    ).doc;
    expect(pageTexts(doc).map((lines) => lines[lines.length - 1])).toEqual([
      'p1',
      'p2',
      'pi',
      'pii',
    ]);
  });

  it('keeps the cached result of every other field in a band', () => {
    // SECTIONPAGES, a date, and a PAGE inside another field's code: what Word
    // cached is all the reader has of them.
    const footer =
      `${F13} SECTIONPAGES ${F14}5${F15}/${F13} DATE ${F14}1.2.99${F15}/` +
      `${F13} IF ${F13} PAGE ${F14}1${F15} = 1 "a" "b" ${F14}a${F15}`;
    const doc = readDoc(
      buildDoc([{ text: `one${PB}two\r`, compressed: false }], {
        headerFooter: { defaultFooter: footer },
      }),
    ).doc;
    expect(pageTexts(doc).map((lines) => lines[lines.length - 1])).toEqual([
      '5/1.2.99/a',
      '5/1.2.99/a',
    ]);
    const id = doc.section!.footers[0]!.relationshipId;
    const [band] = doc.headersFooters!.get(id)!;
    expect(band?.kind === 'paragraph' && band.paragraph.runs.some((r) => r.field)).toBe(false);
  });

  it('marks a PAGE field in the body, in the formatting of its cached result', () => {
    // The body keeps the cached number, as it does from a docx.
    const doc = readDoc(
      buildDoc([{ text: `Page ${F13} page ${F14}7${F15}\r`, compressed: false }], {
        formatRuns: [{ length: 13 }, { length: 1, bold: true }, { length: 2 }],
      }),
    ).doc;
    expect(
      firstParagraphRuns(doc).map((r) => [r.text, r.field, r.properties.bold === true]),
    ).toEqual([
      ['Page ', undefined, false],
      ['7', 'PAGE', true],
    ]);
  });

  it('keeps an inch on every side where the SEP states no margin, as Word does', () => {
    const doc = sectioned(`one${PB}two\r`, [{ length: 4 }, { length: 4 }]);
    expect(doc.sections[0]!.properties.margins).toEqual({
      top: 72,
      right: 72,
      bottom: 72,
      left: 72,
    });
  });
});

describe('doc reader: styles', () => {
  // §2.9.271 STSH. Measured in Word for Mac on a .doc LibreOffice wrote from a
  // .docx of these styles: each word's size, weight and slant, and each
  // paragraph's alignment, spacing and keep, as Word reports them.
  const NORMAL = {
    istd: 0,
    kind: 'paragraph',
    para: { spaceAfterTwips: 240 },
    chars: { sizeHalfPts: 24 },
  } as const;
  const HEADING = {
    istd: 1,
    kind: 'paragraph',
    base: 0,
    para: { jc: 1, spaceBeforeTwips: 480, keepNext: true },
    chars: { bold: true, sizeHalfPts: 32 },
  } as const;
  const EMPH = { istd: 10, kind: 'character', chars: { italic: true, sizeHalfPts: 20 } } as const;
  const runsOf = (doc: FlowDoc) =>
    paragraphs(doc).map((el) =>
      el.paragraph.runs.map((r) => [
        r.text,
        r.properties.fontSizePt,
        r.properties.bold === true,
        r.properties.italic === true,
      ]),
    );

  it('gives a paragraph the properties of its style and of the one that is based on', () => {
    // Word: "Title" 16 pt bold, centred, 24 pt before and 12 after (Normal's),
    // kept with the next; "Body" 12 pt, 12 after.
    const doc = readDoc(
      buildDoc([{ text: 'Title\rBody\r', compressed: false }], {
        styles: [NORMAL, HEADING],
        paraRuns: [{ length: 6, istd: 1 }, { length: 5 }],
      }),
    ).doc;
    const [title, body] = paragraphs(doc).map((el) => el.paragraph.properties);
    expect(title).toMatchObject({
      alignment: 'center',
      spacingBefore: 24,
      spacingAfter: 12,
      keepNext: true,
    });
    expect(body).toMatchObject({ alignment: 'left', spacingAfter: 12 });
    expect(runsOf(doc)).toEqual([[['Title', 16, true, false]], [['Body', 12, false, false]]]);
  });

  it('lays a character style over the paragraph style, and the run formatting over both', () => {
    // Word: "Body " 12 pt, "emph" 10 pt italic, "bold" 12 pt bold.
    const doc = readDoc(
      buildDoc([{ text: 'Body emphbold\r', compressed: false }], {
        styles: [NORMAL, EMPH],
        paraRuns: [{ length: 14 }],
        formatRuns: [{ length: 5 }, { length: 4, charStyle: 10 }, { length: 5, bold: true }],
      }),
    ).doc;
    expect(runsOf(doc)).toEqual([
      [
        ['Body ', 12, false, false],
        ['emph', 10, false, true],
        ['bold', 12, true, false],
      ],
    ]);
  });

  it('reads a toggle against what the style says', () => {
    // [MS-DOC] ToggleOperand: 0x80 is the style's value, 0x81 its opposite, the
    // way Word 97's heading styles make themselves bold over Normal. "Head "
    // bold and "plain" not in a bold heading; "body" bold in a Normal paragraph.
    const doc = readDoc(
      buildDoc([{ text: 'Head plain\rbody\r', compressed: false }], {
        styles: [NORMAL, HEADING],
        paraRuns: [{ length: 11, istd: 1 }, { length: 5 }],
        formatRuns: [
          { length: 5, bold: 0x80 },
          { length: 6, bold: 0x81 },
          { length: 5, bold: 0x81 },
        ],
      }),
    ).doc;
    expect(runsOf(doc)).toEqual([
      [
        ['Head ', 16, true, false],
        ['plain', 16, false, false],
      ],
      [['body', 12, true, false]],
    ]);
  });

  it('keeps an empty paragraph in the spacing of its style and the size of its mark', () => {
    // Word: the empty paragraph between "one" and "two" has 12 pt after it, as
    // they do.
    const doc = readDoc(
      buildDoc([{ text: 'one\r\rtwo\r', compressed: false }], {
        styles: [NORMAL],
        paraRuns: [{ length: 4 }, { length: 1 }, { length: 4 }],
      }),
    ).doc;
    expect(paragraphs(doc).map((el) => el.paragraph.properties.spacingAfter)).toEqual([12, 12, 12]);
    expect(runsOf(doc)[1]).toEqual([[ZWSP, 12, false, false]]);
  });

  it('reads a list level past the ninth as the first', () => {
    // Word 97's "List Bullet" style says level 10 over a list of the old kind
    // (53446.doc); read as it stands, it indented the item eleven levels.
    const doc = readDoc(
      buildDoc([{ text: 'item\r', compressed: false }], {
        paraRuns: [{ length: 5, listIlfo: 1, listIlvl: 10 }],
      }),
    ).doc;
    expect(paragraphs(doc)[0]!.paragraph.properties.indentLeft).toBe(18);
  });
});
