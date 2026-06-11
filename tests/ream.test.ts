import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildXlsx } from './fixtures/build-xlsx';
import { Ream, convertDocxToPdfSync, convertXlsxToPdfSync } from '@/index';
import { ConversionLossError } from '@/core/ir';
import { remoteFontProvider } from '@/core/fonts/provider';

const fonts = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const DOCX = buildDocxFromBody(
  '<w:p><w:r><w:t>interlayer</w:t></w:r></w:p>' +
    '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
);

describe('Ream (parse once → convert many)', () => {
  it('parse exposes the interlayer and the format', () => {
    const doc = Ream.parse(DOCX);
    expect(doc.format).toBe('docx');
    expect(doc.flow.kind).toBe('flow');
    expect(doc.flow.body.length).toBeGreaterThan(0);
    expect(doc.losses).toEqual([]);
  });

  it('convert("pdf") is byte-identical to the direct docx converter', async () => {
    const viaClass = await Ream.parse(DOCX).convert('pdf', { fonts });
    const direct = convertDocxToPdfSync(DOCX, { fonts });
    expect(Buffer.from(viaClass).equals(Buffer.from(direct))).toBe(true);
  });

  it('convert("pdf") is byte-identical for xlsx too', async () => {
    const xlsx = buildXlsx([
      ['A1', 'B1'],
      ['A2', 'B2'],
    ]);
    const viaClass = await Ream.parse(xlsx).convert('pdf', { fonts });
    const direct = convertXlsxToPdfSync(xlsx, { fonts });
    expect(Buffer.from(viaClass).equals(Buffer.from(direct))).toBe(true);
  });

  it('one parse serves multiple conversions (pdf + svg)', async () => {
    const doc = Ream.parse(DOCX);
    const pdf = await doc.convert('pdf', { fonts });
    const svg = new TextDecoder().decode(await doc.convert('svg', { fonts }));
    expect(pdf.slice(0, 5)).toEqual(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d])); // %PDF-
    expect(svg).toContain('<svg');
    expect(svg).toContain('data-page="1"');
    expect(svg).toContain('>interlayer<');
  });

  it('unknown bytes throw with the reader list', () => {
    expect(() => Ream.parse(new Uint8Array([1, 2, 3, 4]))).toThrow(/docx, xlsx/);
  });

  it('font chain records a substitution loss; strict throws', async () => {
    const fakeFetch = (url: string) => {
      const bytes = url.includes('Bold') ? fonts.bold : fonts.regular;
      return Promise.resolve({
        ok: true,
        status: 200,
        arrayBuffer: () => Promise.resolve(bytes.buffer.slice(0)),
      });
    };
    const providers = [remoteFontProvider({ fetch: fakeFetch })];
    const doc = Ream.parse(DOCX);
    const report = await doc.convertWithReport('pdf', { fontProviders: providers });
    expect(report.losses).toHaveLength(1);
    expect(report.losses[0]!.severity).toBe('substituted');
    await expect(doc.convert('pdf', { fontProviders: providers, strict: true })).rejects.toThrow(
      ConversionLossError,
    );
  });
});
