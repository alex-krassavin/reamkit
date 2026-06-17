// XLS-2 — the BIFF8 reader. A synthesized `.xls` (build-xls → CFB + BIFF record
// stream) reads back into the same SheetDoc the OOXML path produces, so the whole
// render pipeline works on a legacy binary workbook. Covers the cell record
// families (NUMBER / RK / MULRK / LABELSST+SST / BOOLERR / FORMULA), multi-sheet
// directories, merges and the 1904 date flag, plus direct unit tests of the two
// fiddliest pieces: SST CONTINUE-boundary strings and RK decoding.

import { describe, expect, it } from 'vitest';

import {
  boolRec,
  buildXls,
  errRec,
  formulaNumberRec,
  formulaStringRecs,
  labelSstRec,
  mergeCellsRec,
  mulRkRec,
  numberRec,
  rkRec,
} from './fixtures/build-xls';
import type { SheetDoc } from '@/core/ir/sheet';
import type { WorksheetCell } from '@/core/spreadsheet-model';
import { decodeRk, readSst, readXlsToSheetDoc } from '@/excel/xls/biff-reader';
import { projectSheetDoc } from '@/excel/sheet-to-flow';

const cellAt = (
  doc: SheetDoc,
  sheet: number,
  row: number,
  col: number,
): WorksheetCell | undefined =>
  doc.sheets[sheet]!.grid.cells.find((c) => c.row === row && c.column === col);

// The resolved display value: shared strings via the SST, inline text, or the raw value.
const valueAt = (doc: SheetDoc, sheet: number, row: number, col: number): string | undefined => {
  const c = cellAt(doc, sheet, row, col);
  if (!c) return undefined;
  if (c.type === 's') return doc.sharedStrings[Number(c.rawValue)];
  if (c.type === 'inlineStr') return c.inlineText;
  return c.rawValue;
};

describe('xls reader — end to end (XLS-2)', () => {
  it('reads numbers, shared strings, two sheets, merges and the 1904 flag', () => {
    const xls = buildXls({
      date1904: true,
      sst: ['Hello', 'World'],
      sheets: [
        {
          name: 'First',
          records: [
            numberRec(0, 0, 42),
            numberRec(0, 1, 3.5),
            labelSstRec(1, 0, 0), // "Hello"
            labelSstRec(1, 1, 1), // "World"
            mergeCellsRec([{ r0: 2, r1: 2, c0: 0, c1: 1 }]),
          ],
        },
        { name: 'Second', records: [numberRec(0, 0, 7)] },
      ],
    });
    const doc = readXlsToSheetDoc(xls);

    expect(doc.sheets.map((s) => s.name)).toEqual(['First', 'Second']);
    expect(doc.date1904).toBe(true);
    expect(valueAt(doc, 0, 0, 0)).toBe('42');
    expect(valueAt(doc, 0, 0, 1)).toBe('3.5');
    expect(valueAt(doc, 0, 1, 0)).toBe('Hello');
    expect(valueAt(doc, 0, 1, 1)).toBe('World');
    expect(valueAt(doc, 1, 0, 0)).toBe('7');
    expect(doc.sheets[0]!.grid.merges).toEqual([
      { startRow: 2, endRow: 2, startColumn: 0, endColumn: 1 },
    ]);
  });

  it('reads RK and MULRK numeric cells', () => {
    const xls = buildXls({
      sheets: [
        {
          name: 'Nums',
          records: [rkRec(0, 0, 100), mulRkRec(1, 0, [1, 2, 3])],
        },
      ],
    });
    const doc = readXlsToSheetDoc(xls);
    expect(valueAt(doc, 0, 0, 0)).toBe('100');
    expect(valueAt(doc, 0, 1, 0)).toBe('1');
    expect(valueAt(doc, 0, 1, 1)).toBe('2');
    expect(valueAt(doc, 0, 1, 2)).toBe('3');
  });

  it('reads boolean and error cells', () => {
    const xls = buildXls({
      sheets: [{ name: 'Be', records: [boolRec(0, 0, true), errRec(0, 1, 0x07)] }],
    });
    const doc = readXlsToSheetDoc(xls);
    expect(cellAt(doc, 0, 0, 0)).toMatchObject({ type: 'b', rawValue: '1' });
    expect(cellAt(doc, 0, 0, 1)).toMatchObject({ type: 'e', rawValue: '#DIV/0!' });
  });

  it('reads cached formula results (number and string)', () => {
    const xls = buildXls({
      sheets: [
        {
          name: 'Fx',
          records: [formulaNumberRec(0, 0, 12.25), ...formulaStringRecs(1, 0, 'computed')],
        },
      ],
    });
    const doc = readXlsToSheetDoc(xls);
    expect(valueAt(doc, 0, 0, 0)).toBe('12.25');
    expect(valueAt(doc, 0, 1, 0)).toBe('computed');
  });

  it('projects an .xls to a flow document with a grid table', () => {
    const xls = buildXls({
      sst: ['Title'],
      sheets: [{ name: 'S', records: [labelSstRec(0, 0, 0), numberRec(0, 1, 9)] }],
    });
    const flow = projectSheetDoc(readXlsToSheetDoc(xls));
    const table = flow.body.find((el) => el.kind === 'table');
    expect(table?.kind).toBe('table');
    if (table?.kind === 'table') expect(table.table.rows.length).toBeGreaterThan(0);
  });
});

describe('SST continuation strings (XLS-2)', () => {
  // Build raw SST blocks: [cstTotal][cstUnique][cch][flags][chars…] + CONTINUEs.
  const u32 = (n: number): Array<number> => [
    n & 0xff,
    (n >> 8) & 0xff,
    (n >> 16) & 0xff,
    (n >> 24) & 0xff,
  ];
  const u16 = (n: number): Array<number> => [n & 0xff, (n >> 8) & 0xff];
  const ascii = (s: string): Array<number> => [...s].map((c) => c.charCodeAt(0));
  const utf16 = (s: string): Array<number> => [...s].flatMap((c) => u16(c.charCodeAt(0)));

  it('joins a string split across a CONTINUE boundary', () => {
    const head = [...u32(1), ...u32(1)]; // cstTotal, cstUnique
    const block0 = Uint8Array.from([...head, ...u16(10), 0x00, ...ascii('HELLO')]); // cch=10, compressed
    const block1 = Uint8Array.from([0x00, ...ascii('WORLD')]); // CONTINUE: grbit=compressed
    expect(readSst([block0, block1])).toEqual(['HELLOWORLD']);
  });

  it('re-reads the compression flag at the boundary (compressed → uncompressed)', () => {
    const head = [...u32(1), ...u32(1)];
    const block0 = Uint8Array.from([...head, ...u16(4), 0x00, ...ascii('AB')]); // 2 compressed chars
    const block1 = Uint8Array.from([0x01, ...utf16('CD')]); // CONTINUE: grbit=uncompressed
    expect(readSst([block0, block1])).toEqual(['ABCD']);
  });

  it('reads an uncompressed (UTF-16) string', () => {
    const head = [...u32(1), ...u32(1)];
    const block = Uint8Array.from([...head, ...u16(3), 0x01, ...utf16('Ünï')]);
    expect(readSst([block])).toEqual(['Ünï']);
  });
});

describe('RK number decoding (XLS-2)', () => {
  // rk encodings: bit0 = ×1/100, bit1 = integer.
  it('decodes integer, ×100 and IEEE RK values', () => {
    expect(decodeRk(((100 << 2) | 0x02) >>> 0)).toBe(100); // integer
    expect(decodeRk(((1234 << 2) | 0x03) >>> 0)).toBe(12.34); // integer ×1/100
    // IEEE: 1.5 = 0x3FF8000000000000; the high 30 bits go into rk (low 2 bits = flags 0).
    const high = 0x3ff80000;
    expect(decodeRk(high >>> 0)).toBeCloseTo(1.5, 10);
  });
});
