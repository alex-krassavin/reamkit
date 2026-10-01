// A sheet that reads from the right (§18.3.1.87 `rightToLeft`), and text that
// does. Measured against Excel's own PDF of a probe workbook (2026-10-01):
// Excel turns the whole grid round — column A at the right edge, each cell's
// left and right borders crossed over, the drawings mirrored but not turned —
// and sets each cell's text as it would anyway: a General cell against the
// side its words start from (Hebrew to the right, on ANY sheet), numbers to
// the right, an explicit side where it names one. Text runs on past its cell
// the way its alignment points, and on paper the sheet meets the right margin.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import type { BodyElement, ShapeBlock, Table } from '@/core/document-model';
import { FontRegistry } from '@/core/font';
import { Ream } from '@/core/converter/ream';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { writeXlsx } from '@/excel/xlsx-writer';
import { projectSheetDoc } from '@/excel/sheet-to-flow';

// Arimo carries Hebrew, so a Hebrew cell lays out at its real width.
const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Arimo-Regular.ttf')),
};

// cellXfs: 0 General, 1 left, 2 right, 3 a red left rule and a blue right
// one, 4 a yellow fill.
const STYLES = `
  <fonts count="1"><font><sz val="11"/><name val="Arial"/></font></fonts>
  <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills>
  <borders count="2"><border/><border><left style="thick"><color rgb="FFFF0000"/></left><right style="thick"><color rgb="FF0000FF"/></right></border></borders>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment horizontal="left"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>
  </cellXfs>`;

interface PlacedText {
  readonly text: string;
  readonly x: number;
  readonly width: number;
}

/** Every text line the layout placed: where it starts and how wide it is. */
function placed(xlsx: Uint8Array): Array<PlacedText> {
  const flow = Ream.parse(xlsx).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  const out: Array<PlacedText> = [];
  for (const command of laid.pages[0]!.commands) {
    if (command.type !== 'line') continue;
    const tokens = command.line.tokens.filter((t) => t.kind === 'text');
    const text = tokens
      .map((t) => t.text)
      .join('')
      .trim();
    const width = tokens.reduce((sum, t) => sum + t.widthPt, 0);
    if (text.length > 0) out.push({ text, x: command.originX, width });
  }
  return out;
}

const at = (items: ReadonlyArray<PlacedText>, has: string): PlacedText => {
  const hit = items.find((i) => i.text.includes(has));
  if (!hit) throw new Error(`no placed text "${has}" among ${items.map((i) => i.text).join('|')}`);
  return hit;
};

const firstTable = (body: ReadonlyArray<BodyElement>): Table => {
  const el = body.find((b) => b.kind === 'table');
  if (el?.kind !== 'table') throw new Error('no table');
  return el.table;
};

describe('a sheet that reads from the right (§18.3.1.87)', () => {
  it("reads the view's direction, and writes it back", () => {
    const doc = readXlsxToSheetDoc(
      buildXlsx({
        sheets: [
          { name: 'R', rows: [['a']], rightToLeft: true },
          { name: 'L', rows: [['b']] },
        ],
      }),
    );
    expect(doc.sheets.map((s) => s.grid.rightToLeft)).toEqual([true, undefined]);
    const again = readXlsxToSheetDoc(writeXlsx(doc).bytes);
    expect(again.sheets.map((s) => s.grid.rightToLeft)).toEqual([true, undefined]);
  });

  it('puts column A at the right edge and runs the columns leftwards', () => {
    const items = placed(
      buildXlsx({ sheets: [{ name: 'R', rows: [['A1', 'B1', 'C1']], rightToLeft: true }] }),
    );
    expect(at(items, 'A1').x).toBeGreaterThan(at(items, 'B1').x);
    expect(at(items, 'B1').x).toBeGreaterThan(at(items, 'C1').x);
  });

  it('meets the right margin on paper, and starts at the left of a window', () => {
    const doc = readXlsxToSheetDoc(
      buildXlsx({ sheets: [{ name: 'R', rows: [['x']], rightToLeft: true }] }),
    );
    expect(firstTable(projectSheetDoc(doc).body).properties.alignment).toBe('right');
    expect(firstTable(projectSheetDoc(doc, { screen: true }).body).properties.alignment).toBe(
      undefined,
    );
  });

  it("crosses each cell's left and right borders over with it", () => {
    const doc = readXlsxToSheetDoc(
      buildXlsx({
        sheets: [{ name: 'R', rows: [['', { value: 'ruled', styleIndex: 3 }]], rightToLeft: true }],
        stylesXml: STYLES,
      }),
    );
    const row = firstTable(projectSheetDoc(doc, { screen: true }).body).rows[0]!;
    // B1 comes first now; its left rule (red) is on its right, toward A.
    const ruled = row.cells[0]!;
    expect(ruled.properties.borders?.right?.colorHex).toBe('FF0000');
    expect(ruled.properties.borders?.left?.colorHex).toBe('0000FF');
  });

  it('keeps the side a cell names, and sets a General one by its words', () => {
    const rows = [
      [{ value: 'named left', styleIndex: 1 }],
      ['general'],
      ['שלום עולם'],
      [{ value: 'named right', styleIndex: 2 }],
      [1234],
    ];
    // On either sheet: "left" and Latin text from the left edge, Hebrew and
    // numbers and "right" to the right one. The sheet's direction does not
    // enter into it.
    for (const rightToLeft of [false, true]) {
      const items = placed(
        buildXlsx({
          sheets: [{ name: 'S', rows, rightToLeft, columns: [{ min: 1, max: 1, widthChars: 30 }] }],
          stylesXml: STYLES,
        }),
      );
      const left = at(items, 'named left');
      const right = at(items, 'named right');
      const rightEdge = right.x + right.width;
      expect(at(items, 'general').x).toBeCloseTo(left.x, 1);
      expect(right.x).toBeGreaterThan(left.x + 50);
      const hebrew = at(items, 'ש');
      expect(hebrew.x + hebrew.width).toBeCloseTo(rightEdge, 0);
      const number = at(items, '1234');
      expect(number.x + number.width).toBeCloseTo(rightEdge, 0);
    }
  });

  it('runs right-aligned text over the empty cells on its left, its paint staying home', () => {
    const doc = readXlsxToSheetDoc(
      buildXlsx({
        sheets: [
          {
            name: 'R',
            rows: [[{ value: 'טקסט ארוך מספיק כדי לגלוש מעבר לתא שלו', styleIndex: 4 }]],
            rightToLeft: true,
            columns: [{ min: 1, max: 4, widthChars: 8 }],
          },
        ],
        stylesXml: STYLES,
      }),
    );
    const flow = projectSheetDoc(doc, { screen: true });
    const table = firstTable(flow.body);
    // A1 spans the free columns to its left; the fill is A1's alone, at the
    // span's far end.
    const cell = table.rows[0]!.cells.find((c) => c.content.length > 0)!;
    expect(cell.properties.colSpan).toBeGreaterThan(1);
    expect(cell.properties.paintColumns).toBe(1);
    expect(cell.properties.paintAtEnd).toBe(true);

    const laid = layoutStyledDocument(flow.body, {
      registry: FontRegistry.fromBytes(FONTS),
      ...flowRenderOptions(flow),
    });
    const fill = laid.pages[0]!.commands.find((c) => c.type === 'fill')!;
    const marginLeft = flow.sections[0]!.properties.margins!.left;
    const gridRight = table.grid.reduce((sum, w) => sum + w, 0);
    const lastColumn = table.grid[table.grid.length - 1]!;
    expect(fill.x - marginLeft + fill.width).toBeCloseTo(gridRight, 1);
    expect(fill.width).toBeCloseTo(lastColumn, 1);
  });

  it('mirrors a drawing across the grid, without turning it round', () => {
    const book = (rightToLeft: boolean): Array<ShapeBlock> => {
      const doc = readXlsxToSheetDoc(
        buildXlsx({
          sheets: [{ name: 'S', rows: [['a', 'b', 'c', 'd', 'e']], rightToLeft }],
          sheetShape: { preset: 'rightArrow', anchor: { from: [1, 1], to: [3, 3] } },
        }),
      );
      return projectSheetDoc(doc, { screen: true }).body.flatMap((el) =>
        el.kind === 'shape' ? [el.shape] : [],
      );
    };
    const [ltr] = book(false);
    const [rtl] = book(true);
    const doc = readXlsxToSheetDoc(
      buildXlsx({
        sheets: [{ name: 'S', rows: [['a', 'b', 'c', 'd', 'e']], rightToLeft: true }],
        sheetShape: { preset: 'rightArrow', anchor: { from: [1, 1], to: [3, 3] } },
      }),
    );
    const width = firstTable(projectSheetDoc(doc, { screen: true }).body).grid.reduce(
      (sum, w) => sum + w,
      0,
    );
    const x = ltr!.float!.posH!.offsetPt!;
    expect(rtl!.float!.posH!.offsetPt).toBeCloseTo(width - x - ltr!.width);
    expect(rtl!.float!.posV!.offsetPt).toBeCloseTo(ltr!.float!.posV!.offsetPt!);
    // The arrow still points right, as Excel draws it.
    expect(rtl!.transform?.flipH).toBeUndefined();
  });
});
