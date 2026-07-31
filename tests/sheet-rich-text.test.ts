// E-SHEET W6 — in-cell rich text. A shared string built from multiple <r> runs,
// each with its own <rPr> (bold / italic / colour / size / vertAlign), projects
// to one document-model run per <r> so a single cell can mix formatting. The
// flattened text still drives value resolution and round-trip; the rich runs are
// render-only (the writer flattens them back to plain text).

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import type { Run } from '@/core/document-model';
import { Ream } from '@/core/converter/ream';
import { convertXlsxToPdfSync } from '@/core/converter';

// One string cell at A1 (interns shared-string index 0), with the shared string
// overridden by a rich <si>: "Total: " (plain) + "42" (bold, red, 14pt).
const RICH_SI =
  '<si>' +
  '<r><t xml:space="preserve">Total: </t></r>' +
  '<r><rPr><b/><color rgb="FFFF0000"/><sz val="14"/></rPr><t>42</t></r>' +
  '</si>';

function cellRuns(xlsx: Uint8Array, row = 0, col = 0): ReadonlyArray<Run> {
  const flow = Ream.parse(xlsx).flow;
  const table = flow.body.find((el) => el.kind === 'table');
  if (table?.kind !== 'table') throw new Error('expected a grid table');
  const cell = table.table.rows[row]?.cells[col];
  const para = cell?.content[0];
  return para?.kind === 'paragraph' ? para.paragraph.runs : [];
}

describe('in-cell rich text (E-SHEET W6)', () => {
  it('splits a rich shared string into one run per <r>', () => {
    const runs = cellRuns(buildXlsx({ rows: [['x']], sharedStringsXml: RICH_SI }));
    expect(runs.map((r) => r.text)).toEqual(['Total: ', '42']);
  });

  it('carries the bold / colour / size run apart from the plain one (props resolved)', () => {
    const runs = cellRuns(buildXlsx({ rows: [['x']], sharedStringsXml: RICH_SI }));
    // "Total: " stays the cell default; "42" takes the <rPr> bold / red / 14pt.
    expect(runs[0]?.properties.bold).toBe(false);
    expect(runs[1]?.properties).toMatchObject({ bold: true, colorHex: 'FF0000' });
    expect(runs[1]?.properties.fontSizePt).toBeCloseTo(14, 5);
    expect(runs[0]?.properties.fontSizePt).not.toBeCloseTo(14, 5);
  });

  it('maps <vertAlign> to super/subscript and <i>/<u> to italic/underline', () => {
    const si =
      '<si>' +
      '<r><rPr><i/><u/></rPr><t>x</t></r>' +
      '<r><rPr><vertAlign val="superscript"/></rPr><t>2</t></r>' +
      '</si>';
    const runs = cellRuns(buildXlsx({ rows: [['p']], sharedStringsXml: si }));
    expect(runs[0]?.properties).toMatchObject({ italic: true, underline: 'single' });
    expect(runs[1]?.properties).toMatchObject({ verticalAlign: 'superscript' });
  });

  it('strikes a run through where §18.8.37 says so', () => {
    // 58315.xlsx reads "320-338 350" with the middle run struck out — the whole
    // point of the cell, and the one boolean toggle this reader skipped.
    const si =
      '<si>' +
      '<r><t>320</t></r>' +
      '<r><rPr><strike/><color rgb="FF0070C0"/></rPr><t>-338</t></r>' +
      '<r><t xml:space="preserve"> 350</t></r>' +
      '</si>';
    const runs = cellRuns(buildXlsx({ rows: [['p']], sharedStringsXml: si }));
    expect(runs.map((r) => r.properties.strike)).toEqual([false, true, false]);
    expect(runs[1]?.properties.colorHex).toBe('0070C0');
  });

  it('lets a run turn bold OFF inside a bold cell', () => {
    // An <rPr> states the run's WHOLE font, so a run that omits <b> is not bold
    // even when the cell is. Read as "bold if set, inherit otherwise",
    // tdf171828.xlsx printed "ohne Sondertilgung" entirely bold where the file
    // bolds "ohne" alone.
    const si =
      '<si>' +
      '<r><rPr><b/></rPr><t xml:space="preserve">ohne</t></r>' +
      '<r><rPr><sz val="11"/></rPr><t xml:space="preserve"> Sondertilgung</t></r>' +
      '</si>';
    const stylesXml =
      `<fonts count="2"><font/><font><b/></font></fonts>` +
      `<fills count="1"><fill/></fills><borders count="1"><border/></borders>` +
      `<cellXfs count="2"><xf/><xf fontId="1" applyFont="1"/></cellXfs>`;
    const runs = cellRuns(
      buildXlsx({ rows: [[{ value: 'x', styleIndex: 1 }]], sharedStringsXml: si, stylesXml }),
    );
    expect(runs.map((r) => r.text)).toEqual(['ohne', ' Sondertilgung']);
    expect(runs[0]?.properties.bold).toBe(true);
    expect(runs[1]?.properties.bold).toBe(false);
  });

  it('leaves a plain (single <t>) shared string as one run, unchanged', () => {
    const runs = cellRuns(buildXlsx({ rows: [['Plain text']] }));
    expect(runs.map((r) => r.text)).toEqual(['Plain text']);
  });

  it('resolves the flattened text for value lookups (round-trip safe)', () => {
    // The cell's value is the concatenation of the runs — drives overflow,
    // number-format, search, etc. exactly as a plain string would.
    const runs = cellRuns(buildXlsx({ rows: [['x']], sharedStringsXml: RICH_SI }));
    expect(runs.map((r) => r.text).join('')).toBe('Total: 42');
  });

  it('renders a rich-text sheet to a valid PDF', () => {
    const pdf = convertXlsxToPdfSync(buildXlsx({ rows: [['x']], sharedStringsXml: RICH_SI }), {
      fonts: {
        regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
        bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
      },
    });
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
    expect(pdf.length).toBeGreaterThan(1000);
  });
});
