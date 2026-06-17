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
