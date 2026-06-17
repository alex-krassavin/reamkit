// XLS-4 — BIFF8 cell styling. A styled `.xls` (FONT / FORMAT / XF / PALETTE
// records) reads into the same XlsxStyles model the OOXML path uses, so the print
// model renders a legacy workbook's fonts, fills, borders, number formats and
// alignment. Asserts the parsed style table, the cell→XF link, and the end-to-end
// resolved run props + shading + formatted number through the projection.

import { describe, expect, it } from 'vitest';

import { buildXls, fontRec, formatRec, labelSstRec, numberRec, xfRec } from './fixtures/build-xls';
import type { BodyElement } from '@/core/document-model';
import { readXlsToSheetDoc } from '@/excel/xls/biff-reader';
import { projectSheetDoc } from '@/excel/sheet-to-flow';

// A styled workbook: font 1 = bold red 14pt; XF 1 = that font + a custom percent
// format + a yellow solid fill + a thin black bottom border + centre alignment.
// Palette indices: 10 = red (FF0000), 13 = yellow (FFFF00), 8 = black (000000).
const styledXls = (): Uint8Array =>
  buildXls({
    sst: ['Hi'],
    styleRecords: [
      fontRec({}), // font 0 (default)
      fontRec({ sizePt: 14, bold: true, colorIndex: 10 }), // font 1
      formatRec(164, '0.00%'),
      xfRec({}), // XF 0 (default)
      xfRec({
        fontId: 1,
        numFmtId: 164,
        halign: 2, // center
        fill: { pattern: 1, fg: 13 }, // solid yellow
        bottom: { style: 1, colorIndex: 8 }, // thin black
      }), // XF 1
    ],
    sheets: [{ name: 'S', records: [numberRec(0, 0, 0.5, 1), labelSstRec(1, 0, 0, 1)] }],
  });

describe('xls styling — parsed style table (XLS-4)', () => {
  const styles = readXlsToSheetDoc(styledXls()).styles;
  const xf = styles.cellXfs[1]!;

  it('reads the font (bold, colour, size) with the index-skip honoured', () => {
    expect(styles.fonts[1]).toMatchObject({ sizePt: 14, bold: true, colorHex: 'FFFF0000' });
  });

  it('reads a custom number format', () => {
    expect(styles.numFmts.get(164)).toBe('0.00%');
  });

  it('links the XF to its font, format, fill and border', () => {
    expect(xf.fontId).toBe(1);
    expect(xf.numFmtId).toBe(164);
    expect(styles.fills[xf.fillId]).toMatchObject({ patternType: 'solid', fgColorHex: 'FFFFFF00' });
    expect(styles.borders[xf.borderId]?.bottom).toMatchObject({
      style: 'thin',
      colorHex: 'FF000000',
    });
    expect(xf.alignment?.horizontal).toBe('center');
  });

  it('sets the cell style index to its XF', () => {
    const a1 = readXlsToSheetDoc(styledXls()).sheets[0]!.grid.cells.find(
      (c) => c.row === 0 && c.column === 0,
    );
    expect(a1?.styleIndex).toBe(1);
  });
});

describe('xls styling — resolved end to end (XLS-4)', () => {
  const a1 = (() => {
    const flow = projectSheetDoc(readXlsToSheetDoc(styledXls()));
    const table = flow.body.find((el) => el.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected a grid table');
    return table.table.rows[0]!.cells[0]!;
  })();

  const firstRun = (content: ReadonlyArray<BodyElement>) => {
    for (const el of content) {
      if (el.kind === 'paragraph' && el.paragraph.runs.length > 0) return el.paragraph.runs[0]!;
    }
    throw new Error('no run');
  };

  it('renders the cell fill', () => {
    expect(a1.properties.shading?.colorHex).toBe('FFFFFF00');
  });

  it('renders the bold red 14pt font and centres the cell', () => {
    const run = firstRun(a1.content);
    expect(run.properties.bold).toBe(true);
    expect(run.properties.colorHex).toBe('FFFF0000');
    expect(run.properties.fontSizePt).toBe(14);
    expect(
      a1.content.some(
        (el) => el.kind === 'paragraph' && el.paragraph.properties.alignment === 'center',
      ),
    ).toBe(true);
  });

  it('applies the custom number format to the value', () => {
    expect(firstRun(a1.content).text).toBe('50.00%');
  });
});
