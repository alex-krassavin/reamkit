// E-SHEET W3 — cell hyperlinks. A worksheet <hyperlinks><hyperlink> with an
// external r:id resolves to a URL through the worksheet relationships; the
// projection stamps run.href on every covered cell, so the existing href path
// draws a PDF /Link annotation and an HTML <a>. In-workbook (location-only)
// links carry no URL and are not rendered. Render-only (not written back).

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { Ream } from '@/core/converter/ream';
import { convertXlsxToPdfSync } from '@/core/converter';

const URL = 'https://www.anthropic.com';

const runHrefAt = (
  flow: ReturnType<typeof Ream.parse>['flow'],
  row: number,
  col: number,
): string | undefined => {
  const table = flow.body.find((el) => el.kind === 'table');
  if (table?.kind !== 'table') throw new Error('expected a table');
  const para = table.table.rows[row]?.cells[col]?.content[0];
  if (para?.kind !== 'paragraph') return undefined;
  return para.paragraph.runs[0]?.href;
};

describe('cell hyperlinks — resolve (E-SHEET W3)', () => {
  it('resolves an external cell hyperlink to its URL', () => {
    const sheet = readXlsxToSheetDoc(
      buildXlsx({ rows: [['Anthropic']], hyperlinks: [{ ref: 'A1', url: URL }] }),
    ).sheets[0]!;
    expect(sheet.hyperlinks).toHaveLength(1);
    expect(sheet.hyperlinks![0]).toMatchObject({ url: URL });
    expect(sheet.hyperlinks![0]!.ref).toMatchObject({ startColumn: 0, startRow: 0 });
  });

  it('drops an in-workbook (location-only) link — no URL to resolve', () => {
    const sheet = readXlsxToSheetDoc(
      buildXlsx({ rows: [['x']], hyperlinks: [{ ref: 'A1', location: 'Sheet1!B2' }] }),
    ).sheets[0]!;
    expect(sheet.hyperlinks).toBeUndefined();
  });

  it('leaves a sheet with no hyperlinks untouched', () => {
    const sheet = readXlsxToSheetDoc(buildXlsx({ rows: [[1]] })).sheets[0]!;
    expect(sheet.hyperlinks).toBeUndefined();
  });
});

describe('cell hyperlinks — projection (E-SHEET W3)', () => {
  it('stamps the URL as the covered cell run href', () => {
    const flow = Ream.parse(
      buildXlsx({ rows: [['Anthropic']], hyperlinks: [{ ref: 'A1', url: URL }] }),
    ).flow;
    expect(runHrefAt(flow, 0, 0)).toBe(URL);
  });

  it('covers every cell of a range hyperlink', () => {
    const flow = Ream.parse(
      buildXlsx({
        rows: [
          ['a', 'b'],
          ['c', 'd'],
        ],
        hyperlinks: [{ ref: 'A1:B2', url: URL }],
      }),
    ).flow;
    expect(runHrefAt(flow, 0, 0)).toBe(URL);
    expect(runHrefAt(flow, 1, 1)).toBe(URL);
  });
});

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
  italic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Italic.ttf')),
  boldItalic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-BoldItalic.ttf')),
};

describe('cell hyperlinks — render (E-SHEET W3)', () => {
  it('renders the cell as an <a> in HTML', async () => {
    const xlsx = buildXlsx({ rows: [['Anthropic']], hyperlinks: [{ ref: 'A1', url: URL }] });
    const html = new TextDecoder().decode(await Ream.parse(xlsx).convert('html'));
    expect(html).toContain('<a ');
    expect(html).toContain(URL);
  });

  it('renders a sheet with a cell hyperlink to a valid PDF', () => {
    const xlsx = buildXlsx({ rows: [['Anthropic']], hyperlinks: [{ ref: 'A1', url: URL }] });
    const pdf = convertXlsxToPdfSync(xlsx, { fonts: FONTS });
    expect(pdf.length).toBeGreaterThan(1000);
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });
});
