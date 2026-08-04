// E-FONT F2 — which substitute families a document needs, read off the parsed
// model rather than one format's XML. Before this the answer came from a regex
// over `w:rFonts`, so every non-docx input was rendered in ONE family: a deck
// whose theme is Times and whose body is Calibri came out entirely in the
// default sans, and no per-run resolution could help because only one registry
// was ever built.

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { Ream } from '@/core/converter/ream';
import { familiesInFlow } from '@/core/fonts/families';

const families = (bytes: Uint8Array): Array<string> =>
  [...familiesInFlow(Ream.parse(bytes).flow)].sort();

/** A workbook whose Normal style names `face`. */
const workbook = (face: string): Uint8Array =>
  buildXlsx({
    rows: [['hello']],
    stylesXml:
      `<fonts count="1"><font><sz val="11"/><name val="${face}"/></font></fonts>` +
      `<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
      `<borders count="1"><border/></borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>`,
  });

describe('the families a document names (E-FONT F2)', () => {
  it('reads a workbook’s own typeface', () => {
    // §18.8.29 `<name>` — carried for every other property of the font and not
    // for this one, so a Calibri workbook asked for no family at all.
    expect(families(workbook('Calibri'))).toEqual(['arimo', 'carlito']);
    expect(families(workbook('Times New Roman'))).toEqual(['arimo', 'tinos']);
  });

  it('always keeps the default sans as the fallback', () => {
    // Unstyled runs, math and chart labels resolve through it.
    expect(families(workbook('Arial'))).toEqual(['arimo']);
  });
});
