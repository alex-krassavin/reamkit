// A hard line feed inside a cell (§18.4.12 `<t>`, written `&#10;`). It divides
// the cell into lines when the cell wraps; when it does not, Excel draws the
// text on one line and the feed — a control character with no width — simply
// goes. preserve_space.xlsx holds a bold "123", a feed and a plain "456", and
// we printed "123 456" where the reference prints "123456".

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { Ream } from '@/core/converter/ream';

const WRAP = `
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1">
      <alignment wrapText="1"/>
    </xf>
  </cellXfs>`;

/** Every line of the first cell, as the page will draw them. */
function lines(styleIndex: number): Array<string> {
  const { flow } = Ream.parse(
    buildXlsx({
      rows: [[{ value: 'one\ntwo', styleIndex }]],
      stylesXml: WRAP,
    }),
  );
  const table = flow.body.find((el) => el.kind === 'table');
  const cell = table?.kind === 'table' ? table.table.rows[0]?.cells[0] : undefined;
  return (cell?.content ?? []).flatMap((block) =>
    block.kind === 'paragraph' ? [block.paragraph.runs.map((r) => r.text).join('')] : [],
  );
}

describe('a hard line feed in a cell', () => {
  it('divides a wrapping cell into its lines', () => {
    expect(lines(1)).toEqual(['one', 'two']);
  });

  it('leaves nothing behind in a cell that does not wrap', () => {
    expect(lines(0)).toEqual(['onetwo']);
  });
});
