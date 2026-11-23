// Issue #145 — a CT_TblWidth written as a percentage WITH its sign (`w:w="100%"`),
// as the docx package wrote every percentage width up to 9.x, or with a unit
// (`"2in"`). The expectations are Word 16.114's readings of our own probes.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { CellProperties, TableProperties } from '@/core/document-model';
import type { PoNode } from '@/core/po-helpers';
import { Ream } from '@/core/converter/ream';
import { twipsToPt } from '@/core/ir';
import { OpcPackage } from '@/core/opc';
import { percentSignValue, poPercentAttr } from '@/core/po-helpers';
import { PdfFile } from '@/pdf-reader/document';
import { parseDocument, parseStyles } from '@/word';
import { cellMarginTwips, readTableMeasure } from '@/word/table-width';

const node = (attrs: Record<string, string>): PoNode => ({
  'a:x': [],
  ':@': Object.fromEntries(Object.entries(attrs).map(([k, v]) => [`@_${k}`, v])),
});

describe('percentSignValue (§22.9.2.9 ST_Percentage)', () => {
  it('reads a percentage written with its sign', () => {
    expect(percentSignValue('100%')).toBe(100);
    expect(percentSignValue('33.3%')).toBe(33.3);
    expect(percentSignValue('-50%')).toBe(-50);
    // Word takes a leading plus and leading zeros as well.
    expect(percentSignValue('+50%')).toBe(50);
    expect(percentSignValue('050%')).toBe(50);
  });

  it('refuses every spelling Word refuses, and bare numbers', () => {
    // " 50%", "50 %" and "50%%" each gave Word's table no width at all.
    for (const raw of [' 50%', '50 %', '50%%', '%', '.5%', '50', '2in', '']) {
      expect(percentSignValue(raw), raw).toBeUndefined();
    }
  });
});

describe('poPercentAttr', () => {
  it('scales the percent form to the unit of the bare-number form', () => {
    // DrawingML counts thousandths of a percent (§20.1.10.40).
    expect(poPercentAttr(node({ val: '75%' }), 'val', 1000)).toBe(75000);
    expect(poPercentAttr(node({ val: '75000' }), 'val', 1000)).toBe(75000);
    // A chart counts whole percents.
    expect(poPercentAttr(node({ val: '150%' }), 'val', 1)).toBe(150);
    expect(poPercentAttr(node({ val: '150' }), 'val', 1)).toBe(150);
  });

  it('reads nothing that is not a number in either spelling', () => {
    expect(poPercentAttr(node({}), 'val', 1000)).toBeUndefined();
    expect(poPercentAttr(node({ val: 'abc' }), 'val', 1000)).toBeUndefined();
    expect(poPercentAttr(undefined, 'val', 1000)).toBeUndefined();
  });
});

describe('readTableMeasure (§17.4.88 CT_TblWidth)', () => {
  it('lets a percent sign overrule the type, as Word does', () => {
    for (const type of ['pct', 'dxa', 'auto', 'nil', undefined]) {
      expect(readTableMeasure('50%', type), String(type)).toEqual({
        kind: 'pct',
        fiftieths: 2500,
        spelled: true,
      });
    }
  });

  it('drops the fraction of a percent, as Word does', () => {
    const fiftieths = (w: string): number | undefined => {
      const m = readTableMeasure(w, 'pct');
      return m?.kind === 'pct' ? m.fiftieths : undefined;
    };
    expect(fiftieths('33.6%')).toBe(33 * 50);
    expect(fiftieths('99.9%')).toBe(99 * 50);
    expect(fiftieths('66.7%')).toBe(66 * 50);
    expect(fiftieths('0.5%')).toBe(0);
    // Fiftieths keep their own precision, but not a decimal place.
    expect(fiftieths('1666')).toBe(1666);
    expect(fiftieths('1666.6')).toBe(1666);
  });

  it('lets a unit overrule the type', () => {
    expect(readTableMeasure('2in', 'pct')).toEqual({ kind: 'dxa', twips: 2880, spelled: true });
    expect(readTableMeasure('72pt', 'pct')).toEqual({ kind: 'dxa', twips: 1440, spelled: true });
    expect(readTableMeasure('0.5in', 'auto')).toEqual({ kind: 'dxa', twips: 720, spelled: true });
  });

  it('reads a bare number by its type', () => {
    expect(readTableMeasure('5000', 'pct')).toEqual({
      kind: 'pct',
      fiftieths: 5000,
      spelled: false,
    });
    expect(readTableMeasure('1440', 'dxa')).toEqual({ kind: 'dxa', twips: 1440, spelled: false });
    expect(readTableMeasure('1440', undefined)).toEqual({
      kind: 'dxa',
      twips: 1440,
      spelled: false,
    });
    expect(readTableMeasure('0', 'auto')).toEqual({ kind: 'auto', twips: 0, spelled: false });
    expect(readTableMeasure('720', 'nil')).toEqual({ kind: 'nil', twips: 720, spelled: false });
  });

  it('reads nothing from a value Word cannot read', () => {
    expect(readTableMeasure(undefined, 'pct')).toBeUndefined();
    expect(readTableMeasure(' 50%', 'pct')).toBeUndefined();
    expect(readTableMeasure('50%%', 'pct')).toBeUndefined();
  });
});

describe('cellMarginTwips', () => {
  it('reads a table default given as a share, auto or nil as no inset', () => {
    // Word drew all of these flush against the border.
    expect(cellMarginTwips('720', 'pct', true)).toBe(0);
    expect(cellMarginTwips('720', 'auto', true)).toBe(0);
    expect(cellMarginTwips('720', 'nil', true)).toBe(0);
    expect(cellMarginTwips('10%', 'dxa', true)).toBe(0);
    expect(cellMarginTwips('-720', 'dxa', true)).toBe(0);
  });

  it("drops a cell's own side given that way, so the table's shows through", () => {
    expect(cellMarginTwips('720', 'pct', false)).toBeUndefined();
    expect(cellMarginTwips('720', 'auto', false)).toBeUndefined();
    expect(cellMarginTwips('720', 'nil', false)).toBeUndefined();
    expect(cellMarginTwips('10%', 'dxa', false)).toBeUndefined();
  });

  it('reads an absolute side in either layer', () => {
    expect(cellMarginTwips('720', 'dxa', true)).toBe(720);
    expect(cellMarginTwips('720', undefined, false)).toBe(720);
    expect(cellMarginTwips('0.5in', 'auto', true)).toBe(720);
    expect(cellMarginTwips('0.5in', 'auto', false)).toBe(720);
  });
});

function tableOf(
  tblPrInner: string,
  tcPrInner = '',
): {
  table: TableProperties;
  cell: CellProperties;
} {
  const docx = buildDocxFromBody(
    `<w:tbl><w:tblPr>${tblPrInner}</w:tblPr>` +
      '<w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid>' +
      `<w:tr><w:tc><w:tcPr>${tcPrInner}</w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>` +
      '<w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
  );
  const el = parseDocument(OpcPackage.open(docx).getMainDocument().data).find(
    (b) => b.kind === 'table',
  );
  if (el?.kind !== 'table') throw new Error('expected a table');
  return { table: el.table.properties, cell: el.table.rows[0]!.cells[0]!.properties };
}

describe('a table width written as a percent string (#145)', () => {
  it("reads the issue's table: tblW 100% and tcW 50%", () => {
    const { table, cell } = tableOf(
      '<w:tblW w:type="pct" w:w="100%"/><w:tblLayout w:type="fixed"/>',
      '<w:tcW w:type="pct" w:w="50%"/>',
    );
    expect(table.widthFraction).toBe(1);
    expect(table.widthType).toBe('pct');
    expect(cell.widthFraction).toBe(0.5);
    expect(cell.widthType).toBe('pct');
  });

  it('takes the type from the percent sign or the unit', () => {
    // Word: "preferred width percent 100" for type dxa.
    expect(tableOf('<w:tblW w:type="dxa" w:w="100%"/>').table).toMatchObject({
      widthFraction: 1,
      widthType: 'pct',
    });
    expect(tableOf('<w:tblW w:type="auto" w:w="50%"/>').table).toMatchObject({
      widthFraction: 0.5,
      widthType: 'pct',
    });
    // Word: "preferred width points 144" for type pct.
    const inches = tableOf('<w:tblW w:type="pct" w:w="2in"/>').table;
    expect(inches.widthPt).toBe(144);
    expect(inches.widthType).toBe('dxa');
    expect(inches.widthFraction).toBeUndefined();
    const cell = tableOf('', '<w:tcW w:type="pct" w:w="1in"/>').cell;
    expect(cell.width).toBe(72);
    expect(cell.widthType).toBe('dxa');
    expect(tableOf('', '<w:tcW w:type="dxa" w:w="25%"/>').cell).toMatchObject({
      widthFraction: 0.25,
      widthType: 'pct',
    });
  });

  it('keeps bare numbers as their type says', () => {
    expect(tableOf('<w:tblW w:type="pct" w:w="2500"/>').table).toMatchObject({
      widthFraction: 0.5,
      widthType: 'pct',
    });
    const auto = tableOf('<w:tblW w:type="auto" w:w="0"/>').table;
    expect(auto.widthType).toBe('auto');
    expect(auto.widthFraction).toBeUndefined();
    // No type at all: the width is kept, the type stays unsaid.
    const untyped = tableOf('<w:tblW w:w="1440"/>').table;
    expect(untyped.widthPt).toBe(72);
    expect(untyped.widthType).toBeUndefined();
  });

  it('indents only by an absolute measure', () => {
    expect(tableOf('<w:tblInd w:w="1in" w:type="dxa"/>').table.indentPt).toBe(72);
    // Word drew both of these flush to the margin.
    expect(tableOf('<w:tblInd w:w="10%" w:type="dxa"/>').table.indentPt).toBeUndefined();
    expect(tableOf('<w:tblInd w:w="500" w:type="pct"/>').table.indentPt).toBeUndefined();
  });

  it('reads cell margins by layer', () => {
    const { table, cell } = tableOf(
      '<w:tblCellMar><w:left w:w="10%" w:type="pct"/><w:right w:w="720" w:type="dxa"/></w:tblCellMar>',
      '<w:tcMar><w:left w:w="10%" w:type="pct"/><w:right w:w="0.5in" w:type="auto"/></w:tcMar>',
    );
    expect(table.defaultCellMargins).toEqual({ left: 0, right: twipsToPt(720) });
    expect(cell.margins).toEqual({ right: twipsToPt(720) });
  });
});

describe("a table style's indent and cell margins", () => {
  it('reads them as a table reads its own', () => {
    const sheet = parseStyles(
      new TextEncoder().encode(
        '<?xml version="1.0" encoding="UTF-8"?>\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
          '<w:style w:type="table" w:styleId="T"><w:tblPr>' +
          '<w:tblInd w:w="0.25in" w:type="dxa"/>' +
          '<w:tblCellMar><w:top w:w="10%" w:type="dxa"/><w:left w:w="720" w:type="auto"/>' +
          '<w:bottom w:w="-100" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar>' +
          '</w:tblPr></w:style>' +
          '<w:style w:type="table" w:styleId="P"><w:tblPr><w:tblInd w:w="10%" w:type="dxa"/></w:tblPr></w:style>' +
          '</w:styles>',
      ),
    );
    const layer = sheet.styles.get('T')?.tableLayer;
    expect(layer?.indentPt).toBe(18);
    expect(layer?.cellMargins).toEqual({ top: 0, left: 0, bottom: 0, right: twipsToPt(90) });
    expect(sheet.styles.get('P')?.tableLayer?.indentPt).toBeUndefined();
  });
});

describe('a percent-string table draws as its fiftieths twin (#145)', () => {
  const fonts = {
    regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
    bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
  };
  const page = async (tblW: string, tcW: string): Promise<string> => {
    const cell = (text: string): string =>
      `<w:tc><w:tcPr><w:tcW w:type="pct" w:w="${tcW}"/></w:tcPr>` +
      `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
    const body =
      '<w:tbl><w:tblPr>' +
      `<w:tblW w:type="pct" w:w="${tblW}"/>` +
      '<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>' +
      '<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>' +
      '<w:insideV w:val="single" w:sz="4"/></w:tblBorders>' +
      '<w:tblLayout w:type="fixed"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid>' +
      `<w:tr>${cell('left cell text')}${cell('right cell text')}</w:tr></w:tbl>`;
    const file = PdfFile.parse(await Ream.parse(buildDocxFromBody(body)).convert('pdf', { fonts }));
    return new TextDecoder('latin1').decode(file.pageContent(file.pages()[0]!));
  };

  it('fills the line instead of shrinking to its 10pt grid', async () => {
    const percent = await page('100%', '50%');
    expect(percent).toBe(await page('5000', '2500'));
    // The rules the table draws: its right edge stands far beyond the 10pt
    // the grid alone would give it.
    const xs = [...percent.matchAll(/\n([\d.]+) [\d.]+ m\n/gu)].map((m) => Number(m[1]));
    expect(Math.max(...xs) - Math.min(...xs)).toBeGreaterThan(400);
  });
});
