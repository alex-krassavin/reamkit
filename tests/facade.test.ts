import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildXlsx } from './fixtures/build-xlsx';
import { convertDocxToPdfSync } from '@/converter';
import { createConverter } from '@/converter/facade';
import { docxReader, readDocx } from '@/readers/docx-reader';
import { xlsxReader } from '@/readers/xlsx-reader';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const DOCX = buildDocxFromBody('<w:p><w:r><w:t>facade</w:t></w:r></w:p>');

describe('reader sniffing', () => {
  it('docxReader sniffs docx and rejects xlsx', () => {
    expect(docxReader.sniff(DOCX)).toBe(true);
    const xlsx = buildXlsx([['a']]);
    expect(docxReader.sniff(xlsx)).toBe(false);
    expect(xlsxReader.sniff(xlsx)).toBe(true);
    expect(xlsxReader.sniff(DOCX)).toBe(false);
  });

  it('rejects non-zip bytes', () => {
    expect(docxReader.sniff(new Uint8Array([1, 2, 3, 4]))).toBe(false);
  });
});

describe('readDocx → FlowDoc', () => {
  it('produces a flow tree with body and sections', () => {
    const { doc, losses } = readDocx(DOCX);
    expect(doc.kind).toBe('flow');
    expect(doc.body.length).toBeGreaterThan(0);
    expect(doc.sections.length).toBeGreaterThan(0);
    expect(losses).toEqual([]);
  });
});

describe('createConverter facade', () => {
  it('detects the format and produces the exact bytes of the direct converter', async () => {
    const ream = createConverter();
    expect(ream.detect(DOCX)?.id).toBe('docx');

    const viaFacade = await ream.convert(DOCX, { fonts: FONTS });
    const direct = convertDocxToPdfSync(DOCX, { fonts: FONTS });
    expect(viaFacade.losses).toEqual([]);
    expect(Buffer.from(viaFacade.bytes).equals(Buffer.from(direct))).toBe(true);
  });

  it('throws on unrecognized input', async () => {
    const ream = createConverter();
    await expect(ream.convert(new Uint8Array([9, 9, 9]), { fonts: FONTS })).rejects.toThrow(
      /Unrecognized input format/,
    );
  });
});
