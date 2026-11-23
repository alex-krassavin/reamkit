// Which cell margin wins when a table, its style, the style's regions and the
// cell each name one. Every expectation is Word 16.114's inset on our own
// probes (the style-cellmar series: the probe ids are in the test names);
// none of it comes from the spec, which ranks nothing here.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { CellMargins, Table } from '@/core/document-model';
import type { BorderItem, TextLineItem } from '@/layout/page-doc';
import { Ream } from '@/core/converter/ream';
import { flowRenderOptions } from '@/core/converter/project';
import { FontRegistry } from '@/core/font';
import { twipsToPt } from '@/core/ir';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { readDocx } from '@/word/docx-reader';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const GRID =
  '<w:tblBorders>' +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`)
    .join('') +
  '</w:tblBorders>';

/** A `w:tblCellMar` / `w:tcMar` with the sides given, in twips. */
const mar = (tag: 'w:tblCellMar' | 'w:tcMar', sides: Partial<Record<Side, number>>): string =>
  `<${tag}>` +
  (['top', 'left', 'bottom', 'right'] as const)
    .filter((s) => sides[s] !== undefined)
    .map((s) => `<w:${s} w:w="${String(sides[s])}" w:type="dxa"/>`)
    .join('') +
  `</${tag}>`;
type Side = 'top' | 'left' | 'bottom' | 'right';

/** Word's own definition of its default table style, as Word writes it. */
const normalTable = (left = 108, right = 108): string =>
  '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>' +
  '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/>' +
  mar('w:tblCellMar', { top: 0, left, bottom: 0, right }) +
  '</w:tblPr></w:style>';

interface StyleSpec {
  readonly basedOn?: string;
  readonly tblCellMar?: Partial<Record<Side, number>>;
  readonly tcMar?: Partial<Record<Side, number>>;
  readonly regions?: string;
}
const tableStyle = (id: string, spec: StyleSpec): string =>
  `<w:style w:type="table" w:styleId="${id}"><w:name w:val="${id}"/>` +
  (spec.basedOn ? `<w:basedOn w:val="${spec.basedOn}"/>` : '') +
  `<w:tblPr>${GRID}${spec.tblCellMar ? mar('w:tblCellMar', spec.tblCellMar) : ''}</w:tblPr>` +
  (spec.tcMar ? `<w:tcPr>${mar('w:tcMar', spec.tcMar)}</w:tcPr>` : '') +
  (spec.regions ?? '') +
  '</w:style>';

/** A `w:tblStylePr` with a table-level and/or a cell-level margin. */
const region = (
  type: string,
  margins: { tbl?: Partial<Record<Side, number>>; tc?: Partial<Record<Side, number>> },
): string =>
  `<w:tblStylePr w:type="${type}">` +
  (margins.tbl ? `<w:tblPr>${mar('w:tblCellMar', margins.tbl)}</w:tblPr>` : '') +
  (margins.tc ? `<w:tcPr>${mar('w:tcMar', margins.tc)}</w:tcPr>` : '') +
  '</w:tblStylePr>';

interface TableSpec {
  readonly style?: string;
  readonly tblCellMar?: Partial<Record<Side, number>>;
  /** The first cell's own `w:tcMar`. */
  readonly tcMar?: Partial<Record<Side, number>>;
  readonly rows?: number;
  readonly look?: string;
}
const tableXml = (spec: TableSpec): string => {
  const cell = (text: string, tcMar?: Partial<Record<Side, number>>): string =>
    `<w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/>${tcMar ? mar('w:tcMar', tcMar) : ''}</w:tcPr>` +
    `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  let rows = '';
  for (let r = 0; r < (spec.rows ?? 1); r++) {
    rows += `<w:tr>${cell(`r${String(r)}a`, r === 0 ? spec.tcMar : undefined)}${cell(`r${String(r)}b`)}</w:tr>`;
  }
  return (
    '<w:tbl><w:tblPr>' +
    (spec.style ? `<w:tblStyle w:val="${spec.style}"/>` : '') +
    '<w:tblW w:w="5760" w:type="dxa"/>' +
    (spec.style ? '' : GRID) +
    '<w:tblLayout w:type="fixed"/>' +
    (spec.tblCellMar ? mar('w:tblCellMar', spec.tblCellMar) : '') +
    (spec.look ??
      '<w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>') +
    '</w:tblPr><w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/></w:tblGrid>' +
    rows +
    '</w:tbl>'
  );
};

const SECT =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>';

function docxOf(styles: string, table: TableSpec): Uint8Array {
  return buildDocxFromBody(tableXml(table) + SECT, { stylesXml: styles });
}

function tableOf(styles: string, spec: TableSpec): Table {
  const { doc } = readDocx(docxOf(styles, spec));
  for (const el of doc.body) if (el.kind === 'table') return el.table;
  throw new Error('no table');
}

/**
 * The margin each side of a cell ends up with, as the layout reads it: the
 * cell's own side, else the table's, else Word's 108 twips left and right.
 */
function effective(table: Table, row: number, col: number): Required<CellMargins> {
  const own = table.rows[row]!.cells[col]!.properties.margins;
  const tbl = table.properties.defaultCellMargins;
  const fallback = { top: 0, left: 5.4, bottom: 0, right: 5.4 };
  const side = (s: Side): number => own?.[s] ?? tbl?.[s] ?? fallback[s];
  return {
    top: side('top'),
    left: side('left'),
    bottom: side('bottom'),
    right: side('right'),
  } as Required<CellMargins>;
}

/** Left/right inset in points, rounded to a tenth, of every cell of a row. */
function insets(table: Table, row: number): Array<[number, number]> {
  return [0, 1].map((col) => {
    const m = effective(table, row, col);
    return [Math.round(m.left * 10) / 10, Math.round(m.right * 10) / 10];
  });
}

const NT = normalTable();
const TABLE_GRID = tableStyle('TableGrid', { basedOn: 'TableNormal' });
const M360 = tableStyle('M360', { basedOn: 'TableNormal', tblCellMar: { left: 360, right: 360 } });

describe("a table's own w:tblCellMar against its style's", () => {
  it('wins: TableGrid with a direct left of 720 insets 36pt (t02)', () => {
    const table = tableOf(NT + TABLE_GRID, { style: 'TableGrid', tblCellMar: { left: 720 } });
    expect(insets(table, 0)).toEqual([
      [36, 5.4],
      [36, 5.4],
    ]);
    // The style's margins are the table's default, not every cell's own.
    expect(table.rows[0]!.cells[0]!.properties.margins).toBeUndefined();
  });

  it('wins side by side, the style filling the sides it leaves (t03, t04, t18)', () => {
    expect(insets(tableOf(NT + M360, { style: 'M360', tblCellMar: { left: 720 } }), 0)[0]).toEqual([
      36, 18,
    ]);
    expect(insets(tableOf(NT + M360, { style: 'M360', tblCellMar: { right: 0 } }), 0)[0]).toEqual([
      18, 0,
    ]);
    const tb = tableStyle('MTB', {
      basedOn: 'TableNormal',
      tblCellMar: { top: 144, left: 360, bottom: 144, right: 360 },
    });
    const table = tableOf(NT + tb, { style: 'MTB', tblCellMar: { top: 0 } });
    expect(table.properties.defaultCellMargins).toEqual({
      top: 0,
      left: twipsToPt(360),
      bottom: twipsToPt(144),
      right: twipsToPt(360),
    });
  });

  it('is drawn where Word draws it', () => {
    const flow = Ream.parse(docxOf(NT + M360, { style: 'M360', tblCellMar: { left: 720 } })).flow;
    const page = layoutStyledDocument(flow.body, {
      registry: FontRegistry.fromBytes(FONTS),
      ...flowRenderOptions(flow),
    }).pages[0]!;
    const frameLeft = Math.min(
      ...page.commands.filter((c): c is BorderItem => c.type === 'border').map((b) => b.x),
    );
    const lineX = (text: string): number => {
      const line = page.commands.find(
        (c): c is TextLineItem =>
          c.type === 'line' && c.line.tokens.some((t) => t.kind === 'text' && t.text === text),
      );
      if (!line) throw new Error(`no line ${text}`);
      return line.originX;
    };
    // Each cell is 144pt wide; its text starts 36pt in. The style's 18pt
    // (written into each cell, as it was) put it at 18.
    expect(lineX('r0a') - frameLeft).toBeCloseTo(36, 1);
    expect(lineX('r0b') - frameLeft).toBeCloseTo(144 + 36, 1);
  });
});

describe('the bottom of the chain', () => {
  it("takes Word's own Normal Table, whatever the document defines (scm2, scm7)", () => {
    const wide = normalTable(720, 720);
    expect(insets(tableOf(wide + TABLE_GRID, { style: 'TableGrid' }), 0)[0]).toEqual([5.4, 5.4]);
    // The same definition under another name is an ordinary style (scm6).
    const renamed = wide.replace('Normal Table', 'Plain Default');
    expect(insets(tableOf(renamed + TABLE_GRID, { style: 'TableGrid' }), 0)[0]).toEqual([36, 36]);
  });

  it('is no margin at all for a style based on nothing (t06, t07)', () => {
    expect(insets(tableOf(NT + tableStyle('Plain', {}), { style: 'Plain' }), 0)[0]).toEqual([0, 0]);
    const plainM = tableStyle('PlainM', { tblCellMar: { left: 360 } });
    expect(insets(tableOf(NT + plainM, { style: 'PlainM' }), 0)[0]).toEqual([18, 0]);
  });

  it('folds side by side through basedOn (t05, t19)', () => {
    const leftOnly = tableStyle('M360L', { basedOn: 'TableNormal', tblCellMar: { left: 360 } });
    expect(insets(tableOf(NT + leftOnly, { style: 'M360L' }), 0)[0]).toEqual([18, 5.4]);
    const kid = tableStyle('Kid', { basedOn: 'M360', tblCellMar: { left: 0 } });
    expect(insets(tableOf(NT + M360 + kid, { style: 'Kid' }), 0)[0]).toEqual([0, 18]);
  });
});

describe('cell-level margins', () => {
  it("a cell's own w:tcMar beats everything, side by side (t09, t10)", () => {
    const table = tableOf(NT + M360, { style: 'M360', tcMar: { left: 1080 } });
    expect(insets(table, 0)).toEqual([
      [54, 18],
      [18, 18],
    ]);
    const mixed = tableOf(NT + M360, {
      style: 'M360',
      tblCellMar: { left: 720 },
      tcMar: { right: 0 },
    });
    expect(insets(mixed, 0)).toEqual([
      [36, 0],
      [36, 18],
    ]);
  });

  it("a style's w:tcMar beats the table's own w:tblCellMar (t14, t20)", () => {
    const sty = tableStyle('StyTc', { basedOn: 'TableNormal', tcMar: { left: 1080, right: 1080 } });
    expect(insets(tableOf(NT + sty, { style: 'StyTc' }), 0)[1]).toEqual([54, 54]);
    const table = tableOf(NT + sty, { style: 'StyTc', tblCellMar: { left: 720 } });
    expect(insets(table, 0)[1]).toEqual([54, 54]);
    expect(table.properties.defaultCellMargins?.left).toBe(36);
  });

  it("folds a style's w:tcMar side by side through basedOn (k02)", () => {
    const parent = tableStyle('P', { basedOn: 'TableNormal', tcMar: { left: 1080, right: 1080 } });
    const kid = tableStyle('K', { basedOn: 'P', tcMar: { left: 0 } });
    expect(insets(tableOf(NT + parent + kid, { style: 'K' }), 0)[0]).toEqual([0, 54]);
  });
});

describe("a style's regions", () => {
  const FIRST_ROW =
    '<w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>';

  it("a region's w:tblCellMar beats the style's but not the table's (t11, t15, x02, x08)", () => {
    const condT = tableStyle('CondT', {
      basedOn: 'TableNormal',
      regions: region('firstRow', { tbl: { left: 1080, right: 1080 } }),
    });
    const alone = tableOf(NT + condT, { style: 'CondT', rows: 2, look: FIRST_ROW });
    expect(insets(alone, 0)[0]).toEqual([54, 54]);
    expect(insets(alone, 1)[0]).toEqual([5.4, 5.4]);
    const direct = tableOf(NT + condT, {
      style: 'CondT',
      rows: 2,
      look: FIRST_ROW,
      tblCellMar: { left: 720 },
    });
    expect(insets(direct, 0)[0]).toEqual([36, 54]);
    expect(insets(direct, 1)[0]).toEqual([36, 5.4]);
  });

  it("a region's w:tcMar beats the table's own w:tblCellMar, a cell's beats it (t16, t17)", () => {
    const condC = tableStyle('CondC', {
      basedOn: 'TableNormal',
      regions: region('firstRow', { tc: { left: 1080, right: 1080 } }),
    });
    const direct = tableOf(NT + condC, {
      style: 'CondC',
      rows: 2,
      look: FIRST_ROW,
      tblCellMar: { left: 720 },
    });
    expect(insets(direct, 0)[0]).toEqual([54, 54]);
    expect(insets(direct, 1)[0]).toEqual([36, 5.4]);
    const own = tableOf(NT + condC, {
      style: 'CondC',
      rows: 2,
      look: FIRST_ROW,
      tcMar: { left: 0 },
    });
    expect(insets(own, 0)).toEqual([
      [0, 54],
      [54, 54],
    ]);
  });

  it('a column region insets the whole row, its shading stays put (c01, x05)', () => {
    const firstCol = tableStyle('FirstC', {
      basedOn: 'TableNormal',
      regions:
        '<w:tblStylePr w:type="firstCol"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"/>' +
        mar('w:tcMar', { left: 1080 }) +
        '</w:tcPr></w:tblStylePr>',
    });
    const table = tableOf(NT + firstCol, {
      style: 'FirstC',
      rows: 2,
      look: '<w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>',
    });
    for (const row of [0, 1])
      expect(insets(table, row)).toEqual([
        [54, 5.4],
        [54, 5.4],
      ]);
    expect(table.rows[0]!.cells[0]!.properties.shading?.colorHex).toBe('DDDDDD');
    expect(table.rows[0]!.cells[1]!.properties.shading).toBeUndefined();
  });

  it('lays a row down cell by cell, so the last cell has the last word (q04, q06)', () => {
    const both =
      '<w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="1" w:noHBand="1" w:noVBand="1"/>';
    // firstCol 1080 in the first cell, lastCol 0 in the second: 0.
    const cols = tableStyle('Cols', {
      basedOn: 'TableNormal',
      regions: region('firstCol', { tc: { left: 1080 } }) + region('lastCol', { tc: { left: 0 } }),
    });
    expect(insets(tableOf(NT + cols, { style: 'Cols', rows: 2, look: both }), 1)[0]).toEqual([
      0, 5.4,
    ]);
    // nwCell 0 outranks firstRow 1080 in the first cell, but the second
    // cell's firstRow comes after it: 54.
    const rowNw = tableStyle('RowNw', {
      basedOn: 'TableNormal',
      regions: region('firstRow', { tc: { left: 1080 } }) + region('nwCell', { tc: { left: 0 } }),
    });
    const table = tableOf(NT + rowNw, { style: 'RowNw', rows: 2, look: both });
    expect(insets(table, 0)).toEqual([
      [54, 5.4],
      [54, 5.4],
    ]);
  });

  it("a derived style's region replaces its parent's margins, unless it states no tblPr or tcPr (k03, k04, q01, q02)", () => {
    const parent = tableStyle('P', {
      basedOn: 'TableNormal',
      regions: region('firstRow', { tc: { left: 1080, right: 1080 } }),
    });
    const firstRow = (kid: string): Array<[number, number]> =>
      insets(tableOf(NT + parent + kid, { style: 'K', rows: 2, look: FIRST_ROW }), 0);
    // A firstRow of its own with a tcMar left of 0: the parent's right is gone.
    expect(
      firstRow(
        tableStyle('K', { basedOn: 'P', regions: region('firstRow', { tc: { left: 0 } }) }),
      )[0],
    ).toEqual([0, 5.4]);
    // …and with nothing but a shading in its tcPr, both sides are.
    const shadeOnly =
      '<w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"/></w:tcPr></w:tblStylePr>';
    expect(firstRow(tableStyle('K', { basedOn: 'P', regions: shadeOnly }))[0]).toEqual([5.4, 5.4]);
    // A firstRow with only an rPr, or none at all, keeps the parent's.
    const boldOnly = '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr>';
    expect(firstRow(tableStyle('K', { basedOn: 'P', regions: boldOnly }))[0]).toEqual([54, 54]);
    expect(firstRow(tableStyle('K', { basedOn: 'P' }))[0]).toEqual([54, 54]);
  });
});

describe('the HTML view', () => {
  it("pads a cell side by side: its own left, the table's right", async () => {
    const docx = docxOf(NT + M360, {
      style: 'M360',
      tblCellMar: { left: 720 },
      tcMar: { left: 0 },
    });
    const html = new TextDecoder().decode(await Ream.parse(docx).convert('html'));
    const pads = [...html.matchAll(/padding:([\d.]+pt [\d.]+pt [\d.]+pt [\d.]+pt)/gu)].map(
      (m) => m[1],
    );
    // top right bottom left: the first cell's own left of 0, the style's
    // right of 18 (not Word's default 5.4); the second cell the table's 36.
    expect(pads).toEqual(['0pt 18pt 0pt 0pt', '0pt 18pt 0pt 36pt']);
  });
});

describe('band regions', () => {
  it('inset nothing unless the style states a band size (t21, x03)', () => {
    const banded = (size: string): Array<[number, number]> => {
      const style =
        '<w:style w:type="table" w:styleId="Band"><w:name w:val="Band"/><w:basedOn w:val="TableNormal"/>' +
        `<w:tblPr>${size}${GRID}</w:tblPr>` +
        region('band1Horz', { tc: { left: 1080 } }) +
        '</w:style>';
      const table = tableOf(NT + style, {
        style: 'Band',
        rows: 3,
        look: '<w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>',
      });
      return [0, 1, 2].map((r) => insets(table, r)[0]!);
    };
    expect(banded('')).toEqual([
      [5.4, 5.4],
      [5.4, 5.4],
      [5.4, 5.4],
    ]);
    expect(banded('<w:tblStyleRowBandSize w:val="1"/>')).toEqual([
      [54, 5.4],
      [5.4, 5.4],
      [54, 5.4],
    ]);
  });
});
