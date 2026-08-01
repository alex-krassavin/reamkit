// The 2009 extension's conditional formats — `<extLst><ext><x14:cfRule>`.
// ISO/IEC 29500 could not express everything Excel 2010 wanted, so the rest
// went here, and nineteen corpus workbooks put rules in it that we read none
// of: tdf122102.xlsx paints four cells yellow, green, red and grey and we
// printed four plain ones.

import { describe, expect, it } from 'vitest';

import { parseWorksheet } from '@/excel/worksheet-parser';

const NS =
  'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
  'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ' +
  'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"';

function sheet(rules: string): Uint8Array {
  return new TextEncoder().encode(
    `<?xml version="1.0"?><worksheet ${NS}><sheetData/>` +
      `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">` +
      `<x14:conditionalFormattings>${rules}</x14:conditionalFormattings>` +
      `</ext></extLst></worksheet>`,
  );
}

const RED = '<x14:dxf><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill></x14:dxf>';

describe('x14 conditional formats', () => {
  it('reads a rule whose range and format are written inside it', () => {
    const ws = parseWorksheet(
      sheet(
        '<x14:conditionalFormatting><x14:cfRule type="expression" priority="1">' +
          `<xm:f>A1&gt;0</xm:f>${RED}</x14:cfRule><xm:sqref>A1:B2</xm:sqref>` +
          '</x14:conditionalFormatting>',
      ),
    );
    const cf = ws.conditionalFormats?.[0];
    expect(cf?.ranges).toEqual([{ startColumn: 0, startRow: 0, endColumn: 1, endRow: 1 }]);
    const rule = cf?.rules[0];
    expect(rule?.type).toBe('expression');
    // The format rides on the rule; `dxfId` names nothing in this schema.
    expect(rule && 'dxf' in rule ? rule.dxf.fill?.bgColorHex : undefined).toBe('FF0000');
  });

  it('reads a text rule as the test it carries, not as the reference’s spelling', () => {
    // The extension exists BECAUSE the needle is a reference: comparing the
    // cell against the literal "$B$1" matches nothing.
    const ws = parseWorksheet(
      sheet(
        '<x14:conditionalFormatting><x14:cfRule type="containsText" operator="containsText" priority="1">' +
          `<xm:f>NOT(ISERROR(SEARCH($B$1,A1)))</xm:f><xm:f>$B$1</xm:f>${RED}` +
          '</x14:cfRule><xm:sqref>A1</xm:sqref></x14:conditionalFormatting>',
      ),
    );
    const rule = ws.conditionalFormats?.[0]?.rules[0];
    expect(rule?.type).toBe('expression');
    expect(rule && 'formula' in rule ? rule.formula : undefined).toBe(
      'NOT(ISERROR(SEARCH($B$1,A1)))',
    );
  });

  it('keeps the rules the base schema declares beside them', () => {
    const ws = parseWorksheet(
      new TextEncoder().encode(
        `<?xml version="1.0"?><worksheet ${NS}><sheetData/>` +
          '<conditionalFormatting sqref="C1"><cfRule type="expression" priority="2" dxfId="3">' +
          '<formula>C1&gt;1</formula></cfRule></conditionalFormatting>' +
          '<extLst><ext><x14:conditionalFormattings><x14:conditionalFormatting>' +
          `<x14:cfRule type="expression" priority="1"><xm:f>A1&gt;0</xm:f>${RED}</x14:cfRule>` +
          '<xm:sqref>A1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings>' +
          '</ext></extLst></worksheet>',
      ),
    );
    expect(ws.conditionalFormats).toHaveLength(2);
    const base = ws.conditionalFormats?.[0]?.rules[0];
    expect(base && 'dxfId' in base ? base.dxfId : undefined).toBe(3);
  });

  it('ignores a block that names no range', () => {
    const ws = parseWorksheet(
      sheet(
        '<x14:conditionalFormatting><x14:cfRule type="expression" priority="1">' +
          `<xm:f>A1&gt;0</xm:f>${RED}</x14:cfRule></x14:conditionalFormatting>`,
      ),
    );
    expect(ws.conditionalFormats).toBeUndefined();
  });
});
