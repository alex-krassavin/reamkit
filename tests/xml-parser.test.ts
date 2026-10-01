// The XML parser takes a tag of any length. fast-xml-parser before 5.7.2 built
// each tag's text by spreading its characters into String.fromCharCode, and a
// tag past about 125 000 characters overflowed the stack (its #817) — a
// RangeError out of the whole conversion. Word writes such tags itself: it
// keeps a DrawingML copy of every VML shape in `o:gfxdata`, base64 of a zip,
// and the corpus has them at 57 000 characters already.

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildXlsx } from './fixtures/build-xlsx';
import { readDocx } from '@/word/docx-reader';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';

const p = (text: string): string => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

describe('a tag of any length', () => {
  it('reads a document whose VML shape carries a long o:gfxdata', () => {
    const shape =
      '<w:p><w:r><w:pict>' +
      '<v:rect xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"' +
      ` style="width:10pt;height:10pt" o:gfxdata="${'UEsDBBQ'.repeat(30_000)}"/>` +
      '</w:pict></w:r></w:p>';
    const { doc } = readDocx(buildDocxFromBody(p('before') + shape + p('after')));
    const text = JSON.stringify(doc.body);
    expect(text).toContain('before');
    expect(text).toContain('after');
  });

  it('reads a workbook whose slicer names a style of 165 000 characters', () => {
    const xlsx = buildXlsx({
      rows: [['Region'], ['North'], ['South']],
      tables: [{ ref: 'A1:A3', name: 'Data' }],
      slicers: [
        {
          name: 'Region',
          caption: 'Region',
          cacheName: 'Slicer_Region',
          styleName: 'SlicerStyle'.repeat(15_000),
          cache: { sourceName: 'Region', tableId: 1, column: 1 },
        },
      ],
    });
    const slicer = readXlsxToSheetDoc(xlsx).sheets[0]!.slicers?.[0];
    expect(slicer?.items.map((i) => i.label)).toEqual(['North', 'South']);
  });
});
