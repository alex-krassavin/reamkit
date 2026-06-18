// XLS-2 — the BIFF8 reader. A synthesized `.xls` (build-xls → CFB + BIFF record
// stream) reads back into the same SheetDoc the OOXML path produces, so the whole
// render pipeline works on a legacy binary workbook. Covers the cell record
// families (NUMBER / RK / MULRK / LABELSST+SST / BOOLERR / FORMULA), multi-sheet
// directories, merges and the 1904 date flag, plus direct unit tests of the two
// fiddliest pieces: SST CONTINUE-boundary strings and RK decoding.

import { describe, expect, it } from 'vitest';

import {
  boolRec,
  bottomMarginRec,
  buildXls,
  cf12ColorScaleRec,
  cf12DataBarRec,
  cf12IconSetRec,
  cfRec,
  commentObjRec,
  condFmt12Rec,
  condFmtRec,
  dvRec,
  errRec,
  footerRec,
  formulaNumberRec,
  formulaStringRecs,
  hCenterRec,
  hPageBreakRec,
  headerRec,
  hlinkRec,
  labelSstRec,
  leftMarginRec,
  mergeCellsRec,
  mulRkRec,
  nameRec,
  noteRec,
  numberRec,
  printGridlinesRec,
  rec,
  rightMarginRec,
  rkRec,
  setupRec,
  topMarginRec,
  txoRecs,
  vCenterRec,
  vPageBreakRec,
  wsBoolRec,
} from './fixtures/build-xls';
import type { FlowDoc } from '@/core/ir/flow';
import type { SheetDoc } from '@/core/ir/sheet';
import type { WorksheetCell } from '@/core/spreadsheet-model';
import { decodeRk, readSst, readXlsToSheetDoc } from '@/excel/xls/biff-reader';
import { projectSheetDoc } from '@/excel/sheet-to-flow';
import { resolvePrintArea, resolvePrintTitleRows } from '@/excel/print-model';

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

describe('xls hyperlinks (XLS-8)', () => {
  // The first run href anywhere in the projected grid table.
  const findHref = (flow: FlowDoc): string | undefined => {
    for (const el of flow.body) {
      if (el.kind !== 'table') continue;
      for (const row of el.table.rows) {
        for (const cell of row.cells) {
          for (const block of cell.content) {
            if (block.kind !== 'paragraph') continue;
            for (const run of block.paragraph.runs) if (run.href) return run.href;
          }
        }
      }
    }
    return undefined;
  };

  it('reads an external URL hyperlink (HLINK) into the SheetDoc', () => {
    const xls = buildXls({
      sst: ['Site'],
      sheets: [
        {
          name: 'S',
          records: [
            labelSstRec(1, 1, 0), // B2 = "Site"
            hlinkRec({
              firstRow: 1,
              lastRow: 1,
              firstCol: 1,
              lastCol: 1,
              url: 'https://example.com/',
            }),
          ],
        },
      ],
    });
    expect(readXlsToSheetDoc(xls).sheets[0]!.hyperlinks).toEqual([
      {
        ref: { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 },
        url: 'https://example.com/',
      },
    ]);
  });

  it('stamps the URL onto the covered cell run through projection', () => {
    const xls = buildXls({
      sst: ['Site'],
      sheets: [
        {
          name: 'S',
          records: [
            labelSstRec(0, 0, 0),
            hlinkRec({
              firstRow: 0,
              lastRow: 0,
              firstCol: 0,
              lastCol: 0,
              url: 'https://reamkit.dev/',
            }),
          ],
        },
      ],
    });
    expect(findHref(projectSheetDoc(readXlsToSheetDoc(xls)))).toBe('https://reamkit.dev/');
  });

  it('ignores a malformed/short HLINK record (no wrong content)', () => {
    const xls = buildXls({
      sheets: [{ name: 'S', records: [numberRec(0, 0, 1), rec(0x01b8, new Uint8Array(10))] }],
    });
    expect(readXlsToSheetDoc(xls).sheets[0]!.hyperlinks).toBeUndefined();
  });

  it('ignores an in-workbook location-only link (no external URL)', () => {
    // An HLink with HasMoniker clear: streamVersion 2, flags 0x08 (HasLocationStr).
    const data = new Uint8Array(32);
    const v = new DataView(data.buffer);
    v.setUint16(0, 0, true); // ref8 A1
    v.setUint32(24, 2, true); // streamVersion
    v.setUint32(28, 0x08, true); // flags: HasLocationStr only (no moniker)
    const xls = buildXls({
      sheets: [{ name: 'S', records: [numberRec(0, 0, 1), rec(0x01b8, data)] }],
    });
    expect(readXlsToSheetDoc(xls).sheets[0]!.hyperlinks).toBeUndefined();
  });
});

describe('xls print model (XLS-9)', () => {
  const sheetDoc = (records: ReadonlyArray<Uint8Array>): SheetDoc =>
    readXlsToSheetDoc(
      buildXls({
        sst: ['Hi'],
        sheets: [{ name: 'S', records: [labelSstRec(0, 0, 0), ...records] }],
      }),
    );
  const gridOf = (records: ReadonlyArray<Uint8Array>) => sheetDoc(records).sheets[0]!.grid;

  it('reads page setup (orientation, scale, fit) and fit-to-page', () => {
    const g = gridOf([
      setupRec({ landscape: true, scale: 80, fitWidth: 2, fitHeight: 3 }),
      wsBoolRec(true),
    ]);
    expect(g.pageSetup?.orientation).toBe('landscape');
    expect(g.pageSetup?.scale).toBe(80);
    expect(g.pageSetup?.fitToWidth).toBe(2);
    expect(g.pageSetup?.fitToHeight).toBe(3);
    expect(g.fitToPage).toBe(true);
  });

  it('reads print options (gridlines + horizontal/vertical centering)', () => {
    const g = gridOf([printGridlinesRec(true), hCenterRec(true), vCenterRec(true)]);
    expect(g.printOptions).toEqual({
      gridLines: true,
      horizontalCentered: true,
      verticalCentered: true,
    });
  });

  it('reads page margins plus header/footer margins from Setup', () => {
    const g = gridOf([
      leftMarginRec(0.5),
      rightMarginRec(0.5),
      topMarginRec(1),
      bottomMarginRec(1),
      setupRec({ headerMarginInches: 0.25, footerMarginInches: 0.25 }),
    ]);
    expect(g.pageMargins).toEqual({
      leftInches: 0.5,
      rightInches: 0.5,
      topInches: 1,
      bottomInches: 1,
      headerInches: 0.25,
      footerInches: 0.25,
    });
  });

  it('reads header and footer strings (an empty record yields none)', () => {
    const g = gridOf([headerRec('&CReport'), footerRec('')]);
    expect(g.headerFooter).toEqual({ oddHeader: '&CReport' });
  });

  it('reads manual row and column breaks', () => {
    const g = gridOf([hPageBreakRec([10, 20]), vPageBreakRec([5])]);
    expect(g.rowBreaks).toEqual([10, 20]);
    expect(g.colBreaks).toEqual([5]);
  });

  it('leaves print fields undefined when no print records are present', () => {
    const g = gridOf([]);
    expect(g.pageSetup).toBeUndefined();
    expect(g.printOptions).toBeUndefined();
    expect(g.pageMargins).toBeUndefined();
    expect(g.rowBreaks).toBeUndefined();
  });

  it('projects a landscape sheet to a wider-than-tall page', () => {
    const flow = projectSheetDoc(sheetDoc([setupRec({ landscape: true, paperSize: 1 })]));
    expect(flow.section.pageSize.width).toBeGreaterThan(flow.section.pageSize.height);
  });
});

describe('xls defined names (XLS-10)', () => {
  const docWith = (globalRecords: ReadonlyArray<Uint8Array>): SheetDoc =>
    readXlsToSheetDoc(
      buildXls({ sheets: [{ name: 'S', records: [numberRec(0, 0, 1)] }], globalRecords }),
    );

  it('reads a Print_Area built-in name scoped to its sheet', () => {
    const doc = docWith([
      nameRec({ builtinId: 0x06, itab: 1, areas: [{ r0: 0, r1: 9, c0: 0, c1: 3 }] }),
    ]);
    expect(doc.definedNames).toEqual([
      { name: '_xlnm.Print_Area', localSheetId: 0, value: 'A1:D10' },
    ]);
    expect(resolvePrintArea(doc.definedNames, 0)).toEqual({
      startColumn: 0,
      startRow: 0,
      endColumn: 3,
      endRow: 9,
    });
  });

  it('reads Print_Titles repeat-rows from the union formula', () => {
    const doc = docWith([
      nameRec({
        builtinId: 0x07,
        itab: 1,
        union: true,
        areas: [
          { r0: 0, r1: 1, c0: 0, c1: 255 }, // rows 1-2 (repeat rows; full column span)
          { r0: 0, r1: 65535, c0: 0, c1: 0 }, // column A (repeat cols; full row span)
        ],
      }),
    ]);
    expect(doc.definedNames[0]!.name).toBe('_xlnm.Print_Titles');
    expect(resolvePrintTitleRows(doc.definedNames, 0)).toEqual({ startRow: 0, endRow: 1 });
  });

  it('reads a regular (user) named range', () => {
    const doc = docWith([
      nameRec({ name: 'MyRange', itab: 0, areas: [{ r0: 1, r1: 4, c0: 1, c1: 2 }] }),
    ]);
    expect(doc.definedNames).toEqual([{ name: 'MyRange', value: 'B2:C5' }]);
  });

  it('skips a name with no resolvable reference (no wrong range)', () => {
    const doc = docWith([nameRec({ name: 'Empty', itab: 0, areas: [] })]);
    expect(doc.definedNames).toEqual([]);
  });
});

describe('xls cell comments (XLS-11)', () => {
  it('reads a cell comment (ref, author, text) into the SheetDoc', () => {
    const doc = readXlsToSheetDoc(
      buildXls({
        sheets: [
          {
            name: 'S',
            records: [
              numberRec(0, 0, 1),
              commentObjRec(1),
              ...txoRecs('Looks off'),
              noteRec({ row: 0, col: 0, idObj: 1, author: 'Alex' }),
            ],
          },
        ],
      }),
    );
    expect(doc.sheets[0]!.comments).toEqual([
      { ref: 'A1', author: 'Alex', text: 'Looks off', threaded: false },
    ]);
  });

  it('joins note → text by object id regardless of record order', () => {
    const doc = readXlsToSheetDoc(
      buildXls({
        sheets: [
          {
            name: 'S',
            records: [noteRec({ row: 2, col: 3, idObj: 7 }), commentObjRec(7), ...txoRecs('Later')],
          },
        ],
      }),
    );
    expect(doc.sheets[0]!.comments).toEqual([{ ref: 'D3', text: 'Later', threaded: false }]);
  });

  it('lists comments in a Comments section in the projection', () => {
    const doc = readXlsToSheetDoc(
      buildXls({
        sheets: [
          {
            name: 'S',
            records: [
              numberRec(0, 0, 1),
              commentObjRec(1),
              ...txoRecs('A remark'),
              noteRec({ row: 0, col: 0, idObj: 1, author: 'Q' }),
            ],
          },
        ],
      }),
    );
    const text = projectSheetDoc(doc)
      .body.filter((el) => el.kind === 'paragraph')
      .flatMap((el) => el.paragraph.runs.map((r) => r.text))
      .join(' ');
    expect(text).toContain('Comments');
    expect(text).toContain('A remark');
  });

  it('skips a note whose object has no text (no wrong content)', () => {
    const doc = readXlsToSheetDoc(
      buildXls({
        sheets: [
          { name: 'S', records: [numberRec(0, 0, 1), noteRec({ row: 0, col: 0, idObj: 9 })] },
        ],
      }),
    );
    expect(doc.sheets[0]!.comments).toBeUndefined();
  });
});

describe('xls data validation (XLS-12)', () => {
  const cellProps = (doc: SheetDoc): Record<string, unknown> | undefined => {
    const table = projectSheetDoc(doc).body.find((el) => el.kind === 'table');
    return table?.kind === 'table' ? table.table.rows[0]?.cells[0]?.properties : undefined;
  };

  it('reads a list validation into the grid', () => {
    const doc = readXlsToSheetDoc(
      buildXls({
        sheets: [
          {
            name: 'S',
            records: [
              numberRec(0, 0, 1),
              dvRec({
                valType: 3,
                listLiteral: 'Yes,No',
                ranges: [{ r0: 0, r1: 0, c0: 0, c1: 0 }],
              }),
            ],
          },
        ],
      }),
    );
    expect(doc.sheets[0]!.grid.dataValidations).toEqual([
      {
        type: 'list',
        ranges: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
        showDropDown: false,
        showInputMessage: true,
        showErrorMessage: true,
        formula1: '"Yes,No"',
      },
    ]);
  });

  it('paints an in-cell dropdown for a list validation cell', () => {
    const doc = readXlsToSheetDoc(
      buildXls({
        sst: ['x'],
        sheets: [
          {
            name: 'S',
            records: [
              labelSstRec(0, 0, 0),
              dvRec({ valType: 3, listLiteral: 'A,B', ranges: [{ r0: 0, r1: 0, c0: 0, c1: 0 }] }),
            ],
          },
        ],
      }),
    );
    expect(cellProps(doc)?.dropdown).toBe(true);
  });

  it('does not paint a dropdown when the combo is suppressed', () => {
    const doc = readXlsToSheetDoc(
      buildXls({
        sst: ['x'],
        sheets: [
          {
            name: 'S',
            records: [
              labelSstRec(0, 0, 0),
              dvRec({
                valType: 3,
                suppressCombo: true,
                listLiteral: 'A',
                ranges: [{ r0: 0, r1: 0, c0: 0, c1: 0 }],
              }),
            ],
          },
        ],
      }),
    );
    expect(doc.sheets[0]!.grid.dataValidations?.[0]?.showDropDown).toBe(true);
    expect(cellProps(doc)?.dropdown).toBeUndefined();
  });

  it('reads a whole-number validation with operator and range', () => {
    const doc = readXlsToSheetDoc(
      buildXls({
        sheets: [
          {
            name: 'S',
            records: [
              numberRec(0, 0, 1),
              dvRec({ valType: 1, operator: 4, ranges: [{ r0: 1, r1: 3, c0: 1, c1: 1 }] }),
            ],
          },
        ],
      }),
    );
    expect(doc.sheets[0]!.grid.dataValidations).toEqual([
      {
        type: 'whole',
        ranges: [{ startRow: 1, endRow: 3, startColumn: 1, endColumn: 1 }],
        operator: 'greaterThan',
        showDropDown: false,
        showInputMessage: true,
        showErrorMessage: true,
      },
    ]);
  });
});

describe('xls conditional formatting (XLS-13)', () => {
  // A cellIs rule over A1:A3: value > 5 → a solid fill from palette index 10 (red).
  const cfSheet = (cellValue: number): SheetDoc =>
    readXlsToSheetDoc(
      buildXls({
        sheets: [
          {
            name: 'S',
            records: [
              numberRec(0, 0, cellValue),
              condFmtRec({ ccf: 1, ranges: [{ r0: 0, r1: 2, c0: 0, c1: 0 }] }),
              cfRec({ operator: 5, intValue: 5, fillFgIcv: 10 }), // greaterThan 5, red fill
            ],
          },
        ],
      }),
    );

  it('reads a cellIs rule and its dxf fill', () => {
    const doc = cfSheet(10);
    expect(doc.sheets[0]!.grid.conditionalFormats).toEqual([
      {
        ranges: [{ startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 }],
        rules: [
          { type: 'cellIs', priority: 1, operator: 'greaterThan', formulas: ['5'], dxfId: 0 },
        ],
      },
    ]);
    expect(doc.styles.dxfs?.[0]?.fill?.patternType).toBe('solid');
    expect(doc.styles.dxfs?.[0]?.fill?.fgColorHex).toMatch(/^FF[0-9A-F]{6}$/);
  });

  it('shades a matching cell through the projection', () => {
    const table = projectSheetDoc(cfSheet(10)).body.find((el) => el.kind === 'table');
    const cell = table?.kind === 'table' ? table.table.rows[0]?.cells[0] : undefined;
    expect(cell?.properties.shading).toBeTruthy();
  });

  it('does not shade a non-matching cell', () => {
    const table = projectSheetDoc(cfSheet(1)).body.find((el) => el.kind === 'table');
    const cell = table?.kind === 'table' ? table.table.rows[0]?.cells[0] : undefined;
    expect(cell?.properties.shading).toBeUndefined();
  });

  it('ignores a CF12 rule that has no CondFmt12 header', () => {
    const doc = readXlsToSheetDoc(
      buildXls({
        sheets: [{ name: 'S', records: [numberRec(0, 0, 1), rec(0x087a, new Uint8Array(20))] }],
      }),
    );
    expect(doc.sheets[0]!.grid.conditionalFormats).toBeUndefined();
  });
});

describe('xls conditional formatting v12 (XLS-CF12)', () => {
  // Read a sheet whose A1:A3 carries the values 1/2/3 plus the given CF12 rule
  // (CondFmt12 header + one CF12 record), produced by the build-xls fixture.
  const cf12Sheet = (rule: Uint8Array): SheetDoc =>
    readXlsToSheetDoc(
      buildXls({
        sheets: [
          {
            name: 'S',
            records: [
              numberRec(0, 0, 1),
              numberRec(1, 0, 2),
              numberRec(2, 0, 3),
              condFmt12Rec({ numcf: 1, ranges: [{ r0: 0, r1: 2, c0: 0, c1: 0 }] }),
              rule,
            ],
          },
        ],
      }),
    );

  it('reads a 3-stop colour scale (cfvos + literal sRGB colours)', () => {
    const doc = cf12Sheet(
      cf12ColorScaleRec({
        priority: 1,
        stops: [
          { cfvo: { type: 2 }, rgb: [0xf8, 0x69, 0x6b] }, // min  → F8696B
          { cfvo: { type: 5, value: 50 }, rgb: [0xff, 0xeb, 0x84] }, // percentile 50 → FFEB84
          { cfvo: { type: 3 }, rgb: [0x63, 0xbe, 0x7b] }, // max  → 63BE7B
        ],
      }),
    );
    expect(doc.sheets[0]!.grid.conditionalFormats).toEqual([
      {
        ranges: [{ startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 }],
        rules: [
          {
            type: 'colorScale',
            priority: 1,
            cfvos: [{ type: 'min' }, { type: 'percentile', val: '50' }, { type: 'max' }],
            colorsHex: ['F8696B', 'FFEB84', '63BE7B'],
          },
        ],
      },
    ]);
  });

  it('reads a data bar (colour + min/max cfvos + percent clamps)', () => {
    const doc = cf12Sheet(
      cf12DataBarRec({
        priority: 2,
        rgb: [0x63, 0x8e, 0xc6],
        percentMin: 10,
        percentMax: 90,
        min: { type: 2 },
        max: { type: 3 },
      }),
    );
    expect(doc.sheets[0]!.grid.conditionalFormats?.[0]?.rules[0]).toEqual({
      type: 'dataBar',
      priority: 2,
      cfvos: [{ type: 'min' }, { type: 'max' }],
      colorHex: '638EC6',
      minLength: 10,
      maxLength: 90,
    });
  });

  it('reads an icon set (named set + per-icon cfvos + reverse)', () => {
    const doc = cf12Sheet(
      cf12IconSetRec({
        priority: 3,
        setId: 3, // 3TrafficLights1
        reverse: true,
        thresholds: [
          { type: 4, value: 0 }, // percent 0
          { type: 4, value: 33 },
          { type: 4, value: 67 },
        ],
      }),
    );
    expect(doc.sheets[0]!.grid.conditionalFormats?.[0]?.rules[0]).toEqual({
      type: 'iconSet',
      priority: 3,
      iconSet: '3TrafficLights1',
      cfvos: [
        { type: 'percent', val: '0' },
        { type: 'percent', val: '33' },
        { type: 'percent', val: '67' },
      ],
      reverse: true,
    });
  });

  it('shades a cell from a colour scale through the projection', () => {
    const doc = cf12Sheet(
      cf12ColorScaleRec({
        stops: [
          { cfvo: { type: 2 }, rgb: [0xff, 0x00, 0x00] },
          { cfvo: { type: 3 }, rgb: [0x00, 0xff, 0x00] },
        ],
      }),
    );
    const table = projectSheetDoc(doc).body.find((el) => el.kind === 'table');
    const cell = table?.kind === 'table' ? table.table.rows[2]?.cells[0] : undefined;
    expect(cell?.properties.shading).toBeTruthy();
  });

  it('reads two CF12 rules sharing one CondFmt12 header', () => {
    const doc = readXlsToSheetDoc(
      buildXls({
        sheets: [
          {
            name: 'S',
            records: [
              numberRec(0, 0, 1),
              condFmt12Rec({ numcf: 2, ranges: [{ r0: 0, r1: 0, c0: 0, c1: 0 }] }),
              cf12DataBarRec({ priority: 1, rgb: [1, 2, 3], min: { type: 2 }, max: { type: 3 } }),
              cf12IconSetRec({
                priority: 2,
                setId: 0,
                thresholds: [
                  { type: 4, value: 0 },
                  { type: 4, value: 33 },
                  { type: 4, value: 67 },
                ],
              }),
            ],
          },
        ],
      }),
    );
    const rules = doc.sheets[0]!.grid.conditionalFormats?.[0]?.rules;
    expect(rules?.map((r) => r.type)).toEqual(['dataBar', 'iconSet']);
  });

  it('skips a graphical rule whose colour is theme-relative, never guessing', () => {
    // A data bar whose ExtendedColor is themed (xclrType 3) cannot be resolved to a
    // concrete sRGB value, so the whole rule is dropped rather than mis-coloured.
    const themed = cf12DataBarRec({ rgb: [0, 0, 0], min: { type: 2 }, max: { type: 3 } });
    // The ExtendedColor sits after rec header(4) + CF12 header(48) + data-bar sub(6);
    // flip its xclrType from 2 (literal RGB) to 3 (themed) so it cannot resolve.
    new DataView(themed.buffer).setUint32(4 + 48 + 6, 3, true);
    const doc = cf12Sheet(themed);
    expect(doc.sheets[0]!.grid.conditionalFormats).toBeUndefined();
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
