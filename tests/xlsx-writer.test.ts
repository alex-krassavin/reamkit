// E-SHEET SD1 — xlsx-writer: SheetDoc → .xlsx through the core OPC writer.
// Structural checks + a basic read→write→read round-trip; the full IR-identity
// gate over a fixture matrix lives in xlsx-roundtrip.test.ts (SD2).

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { writeXlsx } from '@/excel/xlsx-writer';
import { Ream } from '@/core/converter/ream';
import { createConverter } from '@/core/converter/facade';

const STYLES = `
  <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="12"/><color rgb="FFFF0000"/><name val="Arial"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills>
  <borders count="1"><border/></borders>
  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="1" borderId="0" applyFont="1" applyFill="1"/></cellXfs>`;

describe('xlsx-writer — structure (E-SHEET SD1)', () => {
  it('emits a valid OPC package with the workbook parts', () => {
    const sheet = readXlsxToSheetDoc(buildXlsx({ rows: [[1, 'hi']] }));
    const bytes = writeXlsx(sheet).bytes;
    // PK zip magic.
    expect(bytes[0]).toBe(0x50);
    expect(bytes[1]).toBe(0x4b);
    // Re-readable by our own reader.
    const back = readXlsxToSheetDoc(bytes);
    expect(back.sheets).toHaveLength(1);
    expect(back.sheets[0]!.name).toBe('Sheet1');
  });

  it('Ream.parse(xlsx).convert("xlsx") yields a re-readable workbook', async () => {
    const xlsx = buildXlsx({ rows: [[42, 'x']], stylesXml: STYLES });
    const out = await Ream.parse(xlsx).convert('xlsx');
    const back = readXlsxToSheetDoc(out);
    expect(back.sheets[0]!.grid.cells.length).toBe(2);
  });

  it("the facade converts a spreadsheet input with to: 'xlsx'", async () => {
    const conv = createConverter();
    const xlsx = buildXlsx({ rows: [[1, 'x']] });
    const { bytes } = await conv.convert(xlsx, { to: 'xlsx' });
    const back = readXlsxToSheetDoc(bytes);
    expect(back.sheets[0]!.grid.cells.length).toBe(2);
  });
});

describe('xlsx-writer — round-trip basics (E-SHEET SD1)', () => {
  it('preserves cells, types, shared strings and merges', () => {
    const original = readXlsxToSheetDoc(
      buildXlsx({
        rows: [
          [1, 'alpha'],
          [2, 'beta'],
        ],
        mergeRefs: ['A1:B1'],
        stylesXml: STYLES,
      }),
    );
    const round = readXlsxToSheetDoc(writeXlsx(original).bytes);

    expect(round.sharedStrings).toEqual(original.sharedStrings);
    expect(round.sheets[0]!.grid.cells).toEqual(original.sheets[0]!.grid.cells);
    expect(round.sheets[0]!.grid.merges).toEqual(original.sheets[0]!.grid.merges);
    expect(round.sheets[0]!.grid.maxRow).toBe(original.sheets[0]!.grid.maxRow);
    expect(round.sheets[0]!.grid.maxColumn).toBe(original.sheets[0]!.grid.maxColumn);
  });

  it('round-trips the style table (fonts, fills, cellXfs)', () => {
    const original = readXlsxToSheetDoc(buildXlsx({ rows: [[1]], stylesXml: STYLES }));
    const round = readXlsxToSheetDoc(writeXlsx(original).bytes);
    expect(round.styles.fonts).toEqual(original.styles.fonts);
    expect(round.styles.fills).toEqual(original.styles.fills);
    expect(round.styles.cellXfs).toEqual(original.styles.cellXfs);
  });
});
