import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { readDocx } from '@/word/docx-reader';

// A 3×3 table referencing a style, with explicit tblLook flags.
const tbl = (look: string, extraRow1Cell = '') =>
  '<w:tbl><w:tblPr><w:tblStyle w:val="Fancy"/>' +
  look +
  '</w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  ['r0', 'r1', 'r2']
    .map(
      (r, i) =>
        '<w:tr>' +
        ['c0', 'c1', 'c2']
          .map(
            (c, j) =>
              `<w:tc>${i === 0 && j === 0 ? extraRow1Cell : ''}<w:p><w:r><w:t>${r}${c}</w:t></w:r></w:p></w:tc>`,
          )
          .join('') +
        '</w:tr>',
    )
    .join('') +
  '</w:tbl>';

const FANCY_STYLE =
  '<w:style w:type="table" w:styleId="Fancy">' +
  '<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="8"/><w:insideH w:val="single" w:sz="4"/></w:tblBorders>' +
  '<w:tblCellMar><w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
  '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr><w:tcPr><w:shd w:val="clear" w:fill="4472C4"/></w:tcPr></w:tblStylePr>' +
  '<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr>' +
  '</w:style>';

function firstTable(docx: Uint8Array) {
  const { doc } = readDocx(docx);
  for (const el of doc.body) {
    if (el.kind === 'table') return el.table;
  }
  throw new Error('no table');
}

describe('table styles (§17.7.6)', () => {
  it('applies the base layer to table properties and conditional layers per region', () => {
    const docx = buildDocxFromBody(tbl('<w:tblLook w:firstRow="1" w:noHBand="0" w:noVBand="1"/>'), {
      stylesXml: FANCY_STYLE,
    });
    const table = firstTable(docx);

    // Base layer → table chrome.
    expect(table.properties.borders?.top?.width).toBeCloseTo(1); // sz=8 eighths
    expect(table.properties.defaultCellMargins?.left).toBeCloseTo(6); // 120 twips
    // firstRow: shading + bold runs on row 0.
    const r0c0 = table.rows[0]!.cells[0]!;
    expect(r0c0.properties.shading?.colorHex).toBe('4472C4');
    const r0run = r0c0.content[0]!;
    expect(r0run.kind === 'paragraph' && r0run.paragraph.runs[0]!.properties.bold).toBe(true);
    // Banding starts after the first row: row 1 = band1 (shaded), row 2 = band2.
    expect(table.rows[1]!.cells[1]!.properties.shading?.colorHex).toBe('D9E2F3');
    expect(table.rows[2]!.cells[1]!.properties.shading).toBeUndefined();
  });

  it('tblLook gates the conditional formats', () => {
    const docx = buildDocxFromBody(tbl('<w:tblLook w:firstRow="0" w:noHBand="1" w:noVBand="1"/>'), {
      stylesXml: FANCY_STYLE,
    });
    const table = firstTable(docx);
    expect(table.rows[0]!.cells[0]!.properties.shading).toBeUndefined(); // firstRow off
    expect(table.rows[1]!.cells[0]!.properties.shading).toBeUndefined(); // bands off
    // Banding off shifts nothing else; base borders still apply.
    expect(table.properties.borders?.insideH?.style).toBe('single');
  });

  it('parses the legacy hex bitmask form of tblLook', () => {
    // 0x0220 = firstRow (0020) + noHBand (0200).
    const docx = buildDocxFromBody(tbl('<w:tblLook w:val="0220"/>'), { stylesXml: FANCY_STYLE });
    const table = firstTable(docx);
    expect(table.properties.look?.firstRow).toBe(true);
    expect(table.properties.look?.noHBand).toBe(true);
    expect(table.rows[0]!.cells[0]!.properties.shading?.colorHex).toBe('4472C4');
    expect(table.rows[1]!.cells[0]!.properties.shading).toBeUndefined();
  });

  it('direct cell formatting wins over the style layer', () => {
    const docx = buildDocxFromBody(
      tbl(
        '<w:tblLook w:firstRow="1" w:noHBand="1" w:noVBand="1"/>',
        '<w:tcPr><w:shd w:val="clear" w:fill="FF0000"/></w:tcPr>',
      ),
      { stylesXml: FANCY_STYLE },
    );
    const table = firstTable(docx);
    expect(table.rows[0]!.cells[0]!.properties.shading?.colorHex).toBe('FF0000'); // direct
    expect(table.rows[0]!.cells[1]!.properties.shading?.colorHex).toBe('4472C4'); // style
  });

  it('folds the basedOn chain root-first', () => {
    const styles =
      '<w:style w:type="table" w:styleId="Base">' +
      '<w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:val="clear" w:fill="111111"/></w:tcPr></w:tblStylePr>' +
      '</w:style>' +
      '<w:style w:type="table" w:styleId="Fancy"><w:basedOn w:val="Base"/>' +
      '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr>' +
      '</w:style>';
    const docx = buildDocxFromBody(tbl('<w:tblLook w:firstRow="1" w:noHBand="1"/>'), {
      stylesXml: styles,
    });
    const table = firstTable(docx);
    const cell = table.rows[0]!.cells[0]!;
    expect(cell.properties.shading?.colorHex).toBe('111111'); // inherited from Base
    const p = cell.content[0]!;
    expect(p.kind === 'paragraph' && p.paragraph.runs[0]!.properties.bold).toBe(true); // own layer
  });

  it('renders through to HTML (smoke)', async () => {
    const docx = buildDocxFromBody(tbl('<w:tblLook w:firstRow="1" w:noHBand="0"/>'), {
      stylesXml: FANCY_STYLE,
    });
    const html = new TextDecoder().decode(await Ream.parse(docx).convert('html'));
    expect(html).toContain('background-color:#4472C4'); // header row from the style
    expect(html).toContain('background-color:#D9E2F3'); // band
    expect(html).toContain('font-weight:700');
  });

  it('tables without tblStyle are untouched', () => {
    const plain =
      '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const table = firstTable(buildDocxFromBody(plain, { stylesXml: FANCY_STYLE }));
    expect(table.properties.borders).toBeUndefined();
    expect(table.rows[0]!.cells[0]!.properties.shading).toBeUndefined();
  });
});
