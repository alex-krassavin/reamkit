// E-SHEET W6 — cell-format details beyond rich text: wrapText, non-solid /
// gradient fills (summarised to a representative solid), left indent, diagonal
// borders, textRotation (stacked vertical text) and shrinkToFit (the font scaled
// to the column width). All render; the alignment + border fields also round-trip.

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import type { TableCell } from '@/core/document-model';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { writeXlsx } from '@/excel/xlsx-writer';
import { Ream } from '@/core/converter/ream';

// A styles table with two extra fills (lightGray pattern, red→blue gradient), a
// diagonal border and four alignment xfs (wrap / indent×1 / indent×2 / rotation+
// shrink). cellXfs: 0 default, 1 wrap, 2 lightGray fill, 3 gradient fill,
// 4 diagonal border, 5 indent=1, 6 indent=2, 7 rotation+shrink.
const STYLES = `
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="lightGray"><fgColor rgb="FF000000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>
    <fill><gradientFill><stop position="0"><color rgb="FFFF0000"/></stop><stop position="1"><color rgb="FF0000FF"/></stop></gradientFill></fill>
  </fills>
  <borders count="2">
    <border/>
    <border diagonalDown="1"><diagonal style="thin"><color rgb="FF000000"/></diagonal></border>
  </borders>
  <cellXfs count="9">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment indent="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment indent="2"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment textRotation="90" shrinkToFit="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment shrinkToFit="1"/></xf>
  </cellXfs>`;

function firstCell(xlsx: Uint8Array, row = 0, col = 0): TableCell {
  const flow = Ream.parse(xlsx).flow;
  const table = flow.body.find((el) => el.kind === 'table');
  if (table?.kind !== 'table') throw new Error('expected a grid table');
  const cell = table.table.rows[row]?.cells[col];
  if (!cell) throw new Error('no cell');
  return cell;
}

function cellText(cell: TableCell): string {
  const para = cell.content[0];
  return para?.kind === 'paragraph' ? para.paragraph.runs.map((r) => r.text).join('') : '';
}

const LONG = 'a fairly long label that would normally be clipped';

describe('wrapText (E-SHEET W6)', () => {
  it('keeps the full text of a wrapText cell that an occupied neighbour blocks', () => {
    const xlsx = buildXlsx({
      rows: [[{ value: LONG, styleIndex: 1 }, 'X']],
      stylesXml: STYLES,
    });
    expect(cellText(firstCell(xlsx))).toBe(LONG);
  });

  it('still clips the same text without wrapText', () => {
    const xlsx = buildXlsx({ rows: [[{ value: LONG, styleIndex: 0 }, 'X']], stylesXml: STYLES });
    expect(cellText(firstCell(xlsx)).length).toBeLessThan(LONG.length);
  });
});

describe('non-solid + gradient fills (E-SHEET W6)', () => {
  it('blends a lightGray pattern (black over white) to a light grey solid', () => {
    const xlsx = buildXlsx({ rows: [[{ value: 1, styleIndex: 2 }]], stylesXml: STYLES });
    // 25% of black over 75% white → ~BF grey.
    expect(firstCell(xlsx).properties.shading?.colorHex).toBe('BFBFBF');
  });

  it('summarises a red→blue gradient to its mean (purple)', () => {
    const xlsx = buildXlsx({ rows: [[{ value: 1, styleIndex: 3 }]], stylesXml: STYLES });
    expect(firstCell(xlsx).properties.shading?.colorHex).toBe('800080');
  });
});

describe('left indent (E-SHEET W6)', () => {
  it('indents the cell paragraph, proportional to the indent level', () => {
    const i1 = firstCell(buildXlsx({ rows: [[{ value: 'a', styleIndex: 5 }]], stylesXml: STYLES }));
    const i2 = firstCell(buildXlsx({ rows: [[{ value: 'a', styleIndex: 6 }]], stylesXml: STYLES }));
    const para1 = i1.content[0];
    const para2 = i2.content[0];
    const indent1 = para1?.kind === 'paragraph' ? (para1.paragraph.properties.indentLeft ?? 0) : 0;
    const indent2 = para2?.kind === 'paragraph' ? (para2.paragraph.properties.indentLeft ?? 0) : 0;
    expect(indent1).toBeGreaterThan(0);
    expect(indent2).toBeCloseTo(indent1 * 2, 4);
  });
});

describe('diagonal borders (E-SHEET W6)', () => {
  it('carries a diagonalDown border onto the cell', () => {
    const cell = firstCell(buildXlsx({ rows: [[{ value: 1, styleIndex: 4 }]], stylesXml: STYLES }));
    expect(cell.properties.borders?.diagonalDown).toBeDefined();
    expect(cell.properties.borders?.diagonalUp).toBeUndefined();
  });
});

describe('textRotation — stacked vertical text (E-SHEET W6)', () => {
  it('stacks a rotated cell one centred glyph per line', () => {
    const cell = firstCell(
      buildXlsx({ rows: [[{ value: 'Hi', styleIndex: 7 }]], stylesXml: STYLES }),
    );
    expect(cell.content).toHaveLength(2);
    const chars = cell.content.map((el) =>
      el.kind === 'paragraph' ? el.paragraph.runs[0]?.text : '',
    );
    expect(chars).toEqual(['H', 'i']);
    const aligns = cell.content.map((el) =>
      el.kind === 'paragraph' ? el.paragraph.properties.alignment : undefined,
    );
    expect(aligns).toEqual(['center', 'center']);
  });

  it('leaves an unrotated cell as a single paragraph', () => {
    const cell = firstCell(
      buildXlsx({ rows: [[{ value: 'Hi', styleIndex: 0 }]], stylesXml: STYLES }),
    );
    expect(cell.content).toHaveLength(1);
  });
});

describe('shrinkToFit — font scaled to the column (E-SHEET W6)', () => {
  // A long label in a narrow (4-char) column shrinks its font to fit; the run's
  // explicit size drops well below the column's 11pt default.
  function runFontPt(styleIndex: number): number {
    const cell = firstCell(
      buildXlsx({
        rows: [[{ value: 'a really wide label', styleIndex }]],
        columns: [{ min: 1, max: 1, widthChars: 4 }],
        stylesXml: STYLES,
      }),
    );
    const para = cell.content[0];
    return para?.kind === 'paragraph' ? (para.paragraph.runs[0]?.properties.fontSizePt ?? 0) : 0;
  }

  it('scales a shrinkToFit cell’s font down to its column width', () => {
    const shrunk = runFontPt(8);
    const normal = runFontPt(0);
    expect(shrunk).toBeGreaterThan(0);
    expect(shrunk).toBeLessThan(normal); // ~2.5pt vs the 11pt default
  });
});

describe('cell-format round-trip (E-SHEET W6)', () => {
  it('preserves indent / textRotation / shrinkToFit / diagonal across read→write→read', () => {
    const xlsx = buildXlsx({
      rows: [
        [
          { value: 'a', styleIndex: 5 },
          { value: 'b', styleIndex: 7 },
          { value: 'c', styleIndex: 4 },
        ],
      ],
      stylesXml: STYLES,
    });
    const s1 = readXlsxToSheetDoc(xlsx);
    const b1 = writeXlsx(s1).bytes;
    const s2 = readXlsxToSheetDoc(b1);
    const b2 = writeXlsx(s2).bytes;
    expect(b2).toEqual(b1); // idempotent fixpoint
    expect(s2.styles.cellXfs).toEqual(s1.styles.cellXfs);
    expect(s2.styles.borders).toEqual(s1.styles.borders);
    // The new alignment + border fields survived.
    const align7 = s1.styles.cellXfs[7]?.alignment;
    expect(align7).toMatchObject({ textRotation: 90, shrinkToFit: true });
    expect(s1.styles.cellXfs[5]?.alignment).toMatchObject({ indent: 1 });
    expect(s1.styles.borders[1]?.diagonal?.style).toBe('thin');
    expect(s1.styles.borders[1]?.diagonalDown).toBe(true);
  });
});
