// ISO/IEC 29500-3 — `<mc:AlternateContent>` offers the same content twice, and
// a reader that does not implement the namespace its `<mc:Choice>` requires
// takes the `<mc:Fallback>`. Style collections are indexed BY POSITION
// (§18.8.23), so a wrapper the reader skips does not merely lose one entry — it
// shifts every entry after it. style-alternate-content.xlsx declares 29 cell
// formats of which 22 are unwrapped, and its form lost its title's size, its
// table's borders and its alignment together.

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { resolveAlternateContent } from '@/core/opc/alternate-content';
import { Ream } from '@/core/converter/ream';

describe('markup compatibility (ISO/IEC 29500-3)', () => {
  it('leaves a part without a compatibility block as it was read', () => {
    const xml = '<styleSheet><fonts><font><b/></font></fonts></styleSheet>';
    expect(resolveAlternateContent(xml)).toBe(xml);
  });

  it('puts the fallback where the block stood', () => {
    expect(
      resolveAlternateContent(
        '<fonts><font sz="1"/>' +
          '<mc:AlternateContent><mc:Choice Requires="hs"><font sz="2" hs:x="1"/></mc:Choice>' +
          '<mc:Fallback><font sz="2"/></mc:Fallback></mc:AlternateContent>' +
          '<font sz="3"/></fonts>',
      ),
    ).toBe('<fonts><font sz="1"/><font sz="2"/><font sz="3"/></fonts>');
  });

  it('drops a block that offers no fallback', () => {
    expect(
      resolveAlternateContent(
        '<a><mc:AlternateContent><mc:Choice Requires="hs"><b/></mc:Choice></mc:AlternateContent></a>',
      ),
    ).toBe('<a></a>');
  });

  it('unwraps a block nested inside another', () => {
    expect(
      resolveAlternateContent(
        '<a><mc:AlternateContent><mc:Fallback>' +
          '<mc:AlternateContent><mc:Fallback><x/></mc:Fallback></mc:AlternateContent>' +
          '</mc:Fallback></mc:AlternateContent></a>',
      ),
    ).toBe('<a><x/></a>');
  });

  it('keeps a cellXf index counting past a wrapped one', () => {
    // Three formats, the middle one wrapped. The cell asks for index 2 — bold —
    // and reading only the unwrapped pair would hand it the italic at index 1.
    const { flow } = Ream.parse(
      buildXlsx({
        rows: [[{ value: 'x', styleIndex: 2 }]],
        stylesXml: `
          <fonts count="3">
            <font><sz val="11"/></font>
            <font><i/><sz val="11"/></font>
            <font><b/><sz val="11"/></font>
          </fonts>
          <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
          <borders count="1"><border/></borders>
          <cellXfs count="3">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
            <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
              <mc:Choice xmlns:hs="urn:hs" Requires="hs">
                <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1" hs:extension="1"/>
              </mc:Choice>
              <mc:Fallback>
                <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
              </mc:Fallback>
            </mc:AlternateContent>
            <xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="1"/>
          </cellXfs>`,
      }),
    );
    const table = flow.body.find((el) => el.kind === 'table');
    const cell = table?.kind === 'table' ? table.table.rows[0]?.cells[0] : undefined;
    const block = cell?.content[0];
    const run = block?.kind === 'paragraph' ? block.paragraph.runs[0] : undefined;
    expect(run?.properties.bold).toBe(true);
    expect(run?.properties.italic).toBe(false);
  });
});
