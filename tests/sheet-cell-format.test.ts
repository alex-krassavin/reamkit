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

  it('keeps the shrunk text on ONE line — shrinking is what it does instead', () => {
    // §18.8.1: `shrinkToFit` scales the font so the text fits the cell; it does
    // not wrap. Left wrapping, ShrinkToFit.xlsx broke its scaled label onto a
    // second line where both references keep it on one.
    const cell = firstCell(
      buildXlsx({
        rows: [
          [{ value: 'This text is too long for the cell and must be scaled.', styleIndex: 8 }],
        ],
        columns: [{ min: 1, max: 1, widthChars: 10 }],
        stylesXml: STYLES,
      }),
    );
    expect(cell.properties.noWrap).toBe(true);
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

describe('§18.8.22 <u> names an underline, it does not toggle one', () => {
  const runProps = (fontXml: string) => {
    const xlsx = buildXlsx({
      rows: [[{ value: 'Label', styleIndex: 1 }]],
      stylesXml:
        `<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>` +
        `<font><sz val="12"/><name val="DejaVu Sans"/>${fontXml}</font></fonts>` +
        `<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
        `<borders count="1"><border/></borders>` +
        `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
        `<xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/></cellXfs>`,
    });
    const cell = firstCell(xlsx);
    const para = cell.content[0];
    if (para?.kind !== 'paragraph') throw new Error('expected a paragraph');
    return para.paragraph.runs[0]?.properties;
  };

  it('reads val="none" as no underline', () => {
    // `u` is a CT_UnderlineProperty whose val is a NAME out of
    // ST_UnderlineValues, not the CT_BooleanProperty `b`/`i`/`strike` are. Read
    // with the boolean helper beside it, "none" is neither "false" nor "0" and
    // came out true — which put a rule under every text cell of 52348.xlsx,
    // whose two fonts both spell the default out as `<u val="none"/>`.
    expect(runProps('<u val="none"/>')?.underline).not.toBe('single');
    expect(runProps('<u val="false"/>')?.underline).not.toBe('single');
  });

  it('still underlines when the element says so, or says nothing at all', () => {
    expect(runProps('<u/>')?.underline).toBe('single');
    expect(runProps('<u val="single"/>')?.underline).toBe('single');
    expect(runProps('<u val="double"/>')?.underline).toBe('single');
  });
});

describe('§18.18.3 border weights are screen pixels, not eighth-points', () => {
  const widthOf = (style: string): number | undefined => {
    const xlsx = buildXlsx({
      rows: [[{ value: 'Box', styleIndex: 1 }]],
      stylesXml:
        `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
        `<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
        `<borders count="2"><border/>` +
        `<border><top style="${style}"><color rgb="FF000000"/></top></border></borders>` +
        `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
        `<xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1"/></cellXfs>`,
    });
    return firstCell(xlsx).properties.borders?.top?.width;
  };

  it('draws thin at 1px, medium at 2px and thick at 3px (96 DPI)', () => {
    // The names were mapped onto WordprocessingML's eighth-point scale, where
    // the same words mean 0.5 / 1 / 1.5 pt — so every rule on a spreadsheet came
    // out a third too light. 52348.xlsx strokes its red medium frame at 1pt
    // where Excel draws 1.5 and LibreOffice 1.75.
    expect(widthOf('thin')).toBeCloseTo(0.75);
    expect(widthOf('medium')).toBeCloseTo(1.5);
    expect(widthOf('thick')).toBeCloseTo(2.25);
    expect(widthOf('hair')).toBeCloseTo(0.375);
  });
});

describe('a full-width character is one em (UAX #11)', () => {
  it('costs two digit widths, not one', () => {
    // §18.3.1.13 measures a column in Maximum Digit Widths, and a CJK ideograph
    // is a full em — two of them. Falling through to the Latin default charged
    // 1.18 apiece, so 51519.xlsx's 201-character report was cut at 73 where
    // LibreOffice cuts at 47 and Excel at about 40: we were the only one of the
    // three showing nearly twice what the column holds.
    const clip = (text: string): number => {
      const xlsx = buildXlsx({
        rows: [[text, 'blocker']],
        columns: [{ min: 1, max: 1, widthChars: 10 }],
      });
      return cellText(firstCell(xlsx)).length;
    };
    const latin = clip('nnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnn');
    const cjk = clip('豊田製品戦略事業統括本部とよたかいしゃトヨタコメント');
    expect(latin).toBeGreaterThan(8);
    // Half the Latin count, give or take the bucket the Latin letters land in.
    expect(cjk).toBeLessThan(latin * 0.65);
    expect(cjk).toBeGreaterThan(3);
  });
});

describe('a style a whole column or row hands to its unwritten cells', () => {
  // fills[2] is solid green; cellXfs[1] paints it.
  const STYLE_TABLE =
    `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
    `<fills count="3"><fill><patternFill patternType="none"/></fill>` +
    `<fill><patternFill patternType="gray125"/></fill>` +
    `<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill></fills>` +
    `<borders count="1"><border/></borders>` +
    `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
    `<xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/></cellXfs>`;

  it('paints a cell the file never wrote a <c> for (§18.3.1.13 col@style)', () => {
    // 51710.xlsx paints its column A grey with one `<col style="1"/>`, and 588
    // of its rows carry no A cell at all — so the band stopped dead where the
    // cells did, twelve pages short of the end.
    const xlsx = buildXlsx({
      rows: [
        [null, 'b'],
        [null, 'b'],
      ],
      columns: [{ min: 1, max: 1, widthChars: 9, styleIndex: 1 }],
      stylesXml: STYLE_TABLE,
    });
    expect(firstCell(xlsx, 1, 0).properties.shading?.colorHex).toBe('00FF00');
    expect(firstCell(xlsx, 1, 1).properties.shading).toBeUndefined();
  });

  it('lets a row with customFormat do the same (§18.3.1.73 row@s)', () => {
    const xlsx = buildXlsx({
      rows: [
        ['a', 'b'],
        [null, 'x'],
      ],
      rowHeights: [{ row: 1, heightPt: 15, styleIndex: 1 }],
      stylesXml: STYLE_TABLE,
    });
    expect(firstCell(xlsx, 1, 0).properties.shading?.colorHex).toBe('00FF00');
    expect(firstCell(xlsx, 0, 0).properties.shading).toBeUndefined();
  });
});
