// E-SHEET W4 — sheet header/footer text. Excel's &-code mini-language expands
// into one aligned paragraph per region (left / centre / right), with &P/&N as
// dynamic PAGE/NUMPAGES field runs the renderer resolves per page and &A as the
// sheet name. The content rides on FlowDoc.headersFooters and the section's
// header/footer references, so the existing HF band layout paints it.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import type { BodyElement, Run } from '@/core/document-model';
import { buildHeaderFooterContent } from '@/excel/header-footer';
import { flowRenderOptions } from '@/core/converter/project';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { convertXlsxToPdfSync } from '@/core/converter';

const paraRuns = (
  content: ReadonlyArray<BodyElement>,
  alignment: 'left' | 'center' | 'right',
): ReadonlyArray<Run> => {
  const el = content.find(
    (e) => e.kind === 'paragraph' && e.paragraph.properties.alignment === alignment,
  );
  return el?.kind === 'paragraph' ? el.paragraph.runs : [];
};

describe('header/footer &-codes (E-SHEET W4)', () => {
  it('splits &L/&C/&R into three aligned regions', () => {
    const content = buildHeaderFooterContent('&LLeft&CCenter&RRight', 'Sheet1');
    expect(paraRuns(content, 'left')[0]?.text).toBe('Left');
    expect(paraRuns(content, 'center')[0]?.text).toBe('Center');
    expect(paraRuns(content, 'right')[0]?.text).toBe('Right');
  });

  it('treats the default region (no code) as centre', () => {
    const content = buildHeaderFooterContent('My Title', 'Sheet1');
    expect(paraRuns(content, 'center')[0]?.text).toBe('My Title');
    expect(paraRuns(content, 'left')).toEqual([]);
  });

  it('maps &P/&N to PAGE/NUMPAGES field runs and &A to the sheet name', () => {
    const runs = paraRuns(buildHeaderFooterContent('&CPage &P of &N — &A', 'Budget'), 'center');
    expect(runs.map((r) => r.text)).toEqual(['Page ', '1', ' of ', '1', ' — Budget']);
    expect(runs.map((r) => r.field)).toEqual([undefined, 'PAGE', undefined, 'NUMPAGES', undefined]);
  });

  it('applies &B bold and unescapes &&', () => {
    const runs = paraRuns(buildHeaderFooterContent('&CA &B&& B', 'S'), 'center');
    expect(runs[0]).toMatchObject({ text: 'A ', properties: {} });
    expect(runs[runs.length - 1]).toMatchObject({ properties: { bold: true } });
    expect(runs.map((r) => r.text).join('')).toBe('A & B');
  });

  it('drops non-deterministic / styling codes (&D &F &"font" &14)', () => {
    const runs = paraRuns(buildHeaderFooterContent('&C&"Arial,Bold"&14Title &D&F', 'S'), 'center');
    expect(runs.map((r) => r.text).join('')).toBe('Title ');
  });

  it('returns no content for an empty string', () => {
    expect(buildHeaderFooterContent('', 'S')).toEqual([]);
  });
});

describe('sheet header/footer — projection (E-SHEET W4)', () => {
  it('attaches header/footer band content to the section', () => {
    const flow = Ream.parse(
      buildXlsx({
        rows: [['data']],
        headerFooter: { oddHeader: '&CReport', oddFooter: '&RPage &P' },
      }),
    ).flow;
    // The id carries the sheet's index: the map is document-level, so a fixed
    // id let a second sheet overwrite the first's bands and left every sheet
    // after the first with no header at all.
    expect(flow.section?.headers).toEqual([
      { type: 'default', relationshipId: '_xlsxHeaderDefault0' },
    ]);
    expect(flow.section?.footers).toEqual([
      { type: 'default', relationshipId: '_xlsxFooterDefault0' },
    ]);
    expect(flow.headersFooters?.has('_xlsxHeaderDefault0')).toBe(true);
    expect(flow.headersFooters?.has('_xlsxFooterDefault0')).toBe(true);
  });

  it('leaves a sheet with no header/footer without bands', () => {
    const flow = Ream.parse(buildXlsx({ rows: [[1]] })).flow;
    expect(flow.headersFooters).toBeUndefined();
  });
});

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
  italic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Italic.ttf')),
  boldItalic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-BoldItalic.ttf')),
};

function pageText(commands: ReadonlyArray<{ type: string }>): string {
  let out = '';
  for (const c of commands) {
    if (c.type !== 'line') continue;
    const line = (
      c as unknown as { line: { tokens: ReadonlyArray<{ kind: string; text?: string }> } }
    ).line;
    for (const t of line.tokens) if (t.kind === 'text') out += t.text ?? '';
    out += '\n';
  }
  return out;
}

describe('sheet header/footer — render (E-SHEET W4)', () => {
  it('resolves &P/&N per page in the footer band', () => {
    const xlsx = buildXlsx({
      rows: [['a'], ['b']],
      rowBreaks: [1], // row 1 starts page 2 → two pages
      headerFooter: { oddFooter: '&CPage &P of &N' },
    });
    const flow = Ream.parse(xlsx).flow;
    const laid = layoutStyledDocument(flow.body, {
      registry: FontRegistry.fromBytes(FONTS),
      ...flowRenderOptions(flow),
    });
    expect(laid.pages.length).toBe(2);
    expect(pageText(laid.pages[0]!.commands)).toContain('Page 1 of 2');
    expect(pageText(laid.pages[1]!.commands)).toContain('Page 2 of 2');
  });

  it('applies &U, &nn and the style half of &"family,style"', () => {
    // tdf171828.xlsx writes its header as &"BernhardFashion BT,Bold"&16&U… —
    // three codes we dropped, so a 16pt bold underlined title printed as plain
    // body text. The family itself stays dropped (one font set is available).
    const xlsx = buildXlsx({
      rows: [['x']],
      headerFooter: { oddHeader: '&L&"Some Face,Bold Italic"&16&UTitle&RPlain' },
    });
    const flow = Ream.parse(xlsx).flow;
    const bands = [...(flow.headersFooters?.values() ?? [])].flat();
    const runs = bands.flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs : []));
    const title = runs.find((r) => r.text === 'Title');
    expect(title?.properties).toMatchObject({
      bold: true,
      italic: true,
      underline: 'single',
      fontSizePt: 16,
    });
    // Each region starts from the sheet's own font — the right region here is
    // plain, exactly as LibreOffice prints tdf171828's "Seite &P".
    expect(runs.find((r) => r.text === 'Plain')?.properties).toMatchObject({
      bold: false,
      italic: false,
      underline: 'none',
    });
  });

  it('renders a sheet with a header to a valid PDF', () => {
    const xlsx = buildXlsx({ rows: [['x']], headerFooter: { oddHeader: '&CMy Report' } });
    const pdf = convertXlsxToPdfSync(xlsx, { fonts: FONTS });
    expect(pdf.length).toBeGreaterThan(1000);
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });
});

describe('&D and &T print the caller\u2019s reference date', () => {
  it('substitutes the date and the time when one is supplied', () => {
    // customIndexedColors.xlsx heads every page with `&F - &A - &P`, then
    // `&D - &A - &T`. LibreOffice prints "07/31/2026 - - 17:41:36" there.
    const when = new Date(Date.UTC(2026, 6, 31, 17, 41, 36));
    const content = buildHeaderFooterContent(
      '&C&D - &T',
      'Sheet1',
      1,
      11,
      undefined,
      undefined,
      when,
    );
    expect(
      paraRuns(content, 'center')
        .map((r) => r.text)
        .join(''),
    ).toBe('07/31/2026 - 17:41:36');
  });

  it('drops both without one, so the output never reads the wall clock', () => {
    const content = buildHeaderFooterContent('&C&D - &T', 'Sheet1');
    expect(
      paraRuns(content, 'center')
        .map((r) => r.text)
        .join(''),
    ).toBe(' - ');
  });
});
