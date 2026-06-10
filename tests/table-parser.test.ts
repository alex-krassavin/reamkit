import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { eighthPtToPt, emuToPt, halfPtToPt, twipsToPt } from '@/ir';

import { OpcPackage } from '@/opc';
import { parseDocument } from '@/ooxml/wordproc';

function parse(bodyInnerXml: string) {
  const docx = buildDocxFromBody(bodyInnerXml);
  const pkg = OpcPackage.open(docx);
  return parseDocument(pkg.getMainDocument().data);
}

function textOf(cellContent: ReturnType<typeof parse>): string {
  const para = cellContent.find((b) => b.kind === 'paragraph');
  if (!para) return '';
  return para.paragraph.runs.map((r) => r.text).join('');
}

describe('parseTable + body ordering', () => {
  it('parses a 2x2 grid and exposes per-cell paragraphs', () => {
    const body = parse(`
      <w:tbl>
        <w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
        <w:tblGrid>
          <w:gridCol w:w="2500"/>
          <w:gridCol w:w="2500"/>
        </w:tblGrid>
        <w:tr>
          <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>`);

    expect(body).toHaveLength(1);
    expect(body[0]!.kind).toBe('table');
    if (body[0]!.kind !== 'table') throw new Error('unreachable');
    const tbl = body[0]!.table;
    expect(tbl.grid).toEqual([twipsToPt(2500), twipsToPt(2500)]);
    expect(tbl.properties.widthFraction).toBe(1); // tblW 5000 pct = 100%
    expect(tbl.properties.widthType).toBe('pct');
    expect(tbl.rows).toHaveLength(2);
    expect(tbl.rows[0]!.cells).toHaveLength(2);
    expect(textOf(tbl.rows[0]!.cells[0]!.content)).toBe('A1');
    expect(textOf(tbl.rows[0]!.cells[1]!.content)).toBe('B1');
    expect(textOf(tbl.rows[1]!.cells[0]!.content)).toBe('A2');
    expect(textOf(tbl.rows[1]!.cells[1]!.content)).toBe('B2');
  });

  it('parses cell shading (w:shd @w:fill → CellProperties.shading)', () => {
    const body = parse(`
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr>
          <w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="00FFFF"/></w:tcPr><w:p><w:r><w:t>Cyan</w:t></w:r></w:p></w:tc>
          <w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:tcPr><w:p><w:r><w:t>None</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>`);
    if (body[0]!.kind !== 'table') throw new Error('unreachable');
    const cells = body[0]!.table.rows[0]!.cells;
    expect(cells[0]!.properties.shading?.colorHex).toBe('00FFFF'); // explicit fill
    expect(cells[1]!.properties.shading).toBeUndefined(); // fill="auto" → unshaded
  });

  it('preserves declaration order of paragraphs and tables in the body', () => {
    const body = parse(`
      <w:p><w:r><w:t>Before</w:t></w:r></w:p>
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr><w:tc><w:p><w:r><w:t>InTable</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
      <w:p><w:r><w:t>After</w:t></w:r></w:p>`);

    expect(body.map((b) => b.kind)).toEqual(['paragraph', 'table', 'paragraph']);
  });

  it('parses borders and cell properties', () => {
    const body = parse(`
      <w:tbl>
        <w:tblPr>
          <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:color="000000"/>
            <w:bottom w:val="single" w:sz="8" w:color="ff0000"/>
          </w:tblBorders>
          <w:tblCellMar>
            <w:top w:w="80" w:type="dxa"/>
            <w:bottom w:w="80" w:type="dxa"/>
            <w:left w:w="100" w:type="dxa"/>
            <w:right w:w="100" w:type="dxa"/>
          </w:tblCellMar>
        </w:tblPr>
        <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr>
          <w:tc>
            <w:tcPr>
              <w:tcW w:w="2000" w:type="dxa"/>
              <w:gridSpan w:val="2"/>
            </w:tcPr>
            <w:p><w:r><w:t>X</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
      </w:tbl>`);

    if (body[0]!.kind !== 'table') throw new Error('expected table');
    const tbl = body[0]!.table;
    expect(tbl.properties.borders).toEqual({
      top: { style: 'single', width: eighthPtToPt(4), colorHex: '000000' },
      bottom: { style: 'single', width: eighthPtToPt(8), colorHex: 'FF0000' },
    });
    expect(tbl.properties.defaultCellMargins).toEqual({
      top: twipsToPt(80),
      bottom: twipsToPt(80),
      left: twipsToPt(100),
      right: twipsToPt(100),
    });
    const cell = tbl.rows[0]!.cells[0]!;
    expect(cell.properties.width).toBe(twipsToPt(2000));
    expect(cell.properties.gridSpan).toBe(2);
  });

  it('parses vertical merge restart and continue', () => {
    const body = parse(`
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr>
          <w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>X</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
        </w:tr>
      </w:tbl>`);

    if (body[0]!.kind !== 'table') throw new Error('expected table');
    expect(body[0]!.table.rows[0]!.cells[0]!.properties.vMerge).toBe('restart');
    expect(body[0]!.table.rows[1]!.cells[0]!.properties.vMerge).toBe('continue');
  });

  it('supports nested tables (table inside cell)', () => {
    const body = parse(`
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
        <w:tr><w:tc>
          <w:tbl>
            <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
            <w:tr><w:tc><w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:tc></w:tr>
          </w:tbl>
        </w:tc></w:tr>
      </w:tbl>`);

    if (body[0]!.kind !== 'table') throw new Error('expected outer table');
    const cell = body[0]!.table.rows[0]!.cells[0]!;
    expect(cell.content[0]!.kind).toBe('table');
    if (cell.content[0]!.kind !== 'table') throw new Error('expected inner');
    expect(cell.content[0]!.table.grid).toEqual([twipsToPt(2000)]);
  });
});
