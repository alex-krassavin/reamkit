// XLS-3 — the legacy `.xls` reader wired into the converter. A synthesized .xls
// flows through the public APIs: Ream.parse exposes its SheetDoc and the
// styling-deferred losses, and converts to PDF / HTML / xlsx; the facade detects
// it and renders PDF. A CFB that is not a workbook (a .doc-like container) is not
// misdetected.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXls, labelSstRec, numberRec } from './fixtures/build-xls';
import { buildCfb } from './fixtures/build-cfb';
import { Ream } from '@/core/converter/ream';
import { createConverter } from '@/core/converter/facade';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';

const fonts = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const sampleXls = (): Uint8Array =>
  buildXls({
    sst: ['Hello'],
    sheets: [{ name: 'Sheet1', records: [labelSstRec(0, 0, 0), numberRec(0, 1, 42)] }],
  });

describe('xls wired into the converter (XLS-3)', () => {
  it('Ream.parse exposes the SheetDoc, format id and the secondary-features loss', () => {
    const doc = Ream.parse(sampleXls());
    expect(doc.format).toBe('xls');
    expect(doc.sheet?.kind).toBe('sheet');
    expect(doc.sheet?.sheets[0]?.name).toBe('Sheet1');
    expect(doc.losses.some((l) => l.feature === 'cellFormatting')).toBe(true);
  });

  it('converts an .xls to a valid PDF', async () => {
    const pdf = await Ream.parse(sampleXls()).convert('pdf', { fonts });
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
    expect(pdf.length).toBeGreaterThan(1000);
  });

  it('converts an .xls to HTML carrying the cell text', async () => {
    const html = new TextDecoder().decode(await Ream.parse(sampleXls()).convert('html'));
    expect(html).toContain('Hello');
  });

  it('re-writes an .xls to a valid .xlsx (legacy → modern)', async () => {
    const xlsx = await Ream.parse(sampleXls()).convert('xlsx');
    expect(xlsx[0]).toBe(0x50); // ZIP "PK"
    expect(xlsx[1]).toBe(0x4b);
    const reparsed = readXlsxToSheetDoc(xlsx);
    expect(reparsed.sheets[0]?.name).toBe('Sheet1');
    expect(reparsed.sharedStrings).toContain('Hello');
  });

  it('the facade detects .xls and renders it to PDF', async () => {
    const converter = createConverter();
    expect(converter.detect(sampleXls())?.id).toBe('xls');
    const { bytes } = await converter.convert(sampleXls(), { to: 'pdf', fonts });
    expect(new TextDecoder().decode(bytes.subarray(0, 5))).toBe('%PDF-');
  });

  it('does not misdetect a non-workbook compound file as .xls', () => {
    // A WordDocument stream (no Workbook) is a legacy .doc — the .doc reader
    // claims it, the .xls reader must not; and a CFB with neither matches nothing.
    const docLike = buildCfb([{ name: 'WordDocument', data: new Uint8Array(2000) }]);
    expect(createConverter().detect(docLike)?.id).toBe('doc');
    const neither = buildCfb([{ name: 'RandomStream', data: new Uint8Array(2000) }]);
    expect(createConverter().detect(neither)).toBeUndefined();
  });
});
