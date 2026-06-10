import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildXlsx } from './fixtures/build-xlsx';
import { convertDocxToPdfSync } from '@/core/converter';
import { createConverter } from '@/core/converter/facade';
import { docxReader, readDocx } from '@/word/docx-reader';
import { xlsxReader } from '@/excel/xlsx-reader';

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

describe('svg writer (stage-6 crash test)', () => {
  it('converts docx → SVG through FlowDoc → layout → svgWriter', async () => {
    const docx = buildDocxFromBody(`
      <w:p><w:r><w:t>svg smoke</w:t></w:r></w:p>
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr><w:tc>
          <w:tcPr><w:shd w:val="clear" w:fill="FFCC00"/></w:tcPr>
          <w:p><w:r><w:t>cell</w:t></w:r></w:p>
        </w:tc></w:tr>
      </w:tbl>`);
    const ream = createConverter();
    const r = await ream.convert(docx, { to: 'svg', fonts: FONTS });
    const svg = new TextDecoder().decode(r.bytes);
    expect(svg.startsWith('<svg ')).toBe(true);
    expect(svg).toContain('>smoke</text>'); // tokens emit per word
    expect(svg).toContain('cell');
    expect(svg).toContain('fill="#FFCC00"'); // table shading → <rect>
    expect(svg).toContain('data-page="1"');
    expect(svg).toContain('</svg>');
  });
});
