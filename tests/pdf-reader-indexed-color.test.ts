// §8.6.6.3 — a fill set by `sc` in an Indexed space is the table's colour at
// that index, snapped into the table.

import { describe, expect, it } from 'vitest';

import { PdfFile } from '@/pdf-reader/document';
import { collectPageVectors } from '@/pdf-reader/vector';

function swatches(content: string): Uint8Array {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Contents 4 0 R ' +
      '/Resources << /ColorSpace << /Cs1 [/Indexed /DeviceRGB 2 <FF0000 00FF00 0000FF>] >> >> >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
  ];
  let pdf = '%PDF-1.7\n';
  const offsets: Array<number> = [];
  objects.forEach((body, i) => {
    offsets.push(pdf.length);
    pdf += `${String(i + 1)} 0 obj\n${body}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
  pdf += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(xref)}\n%%EOF\n`;
  return new TextEncoder().encode(pdf);
}

describe('a colour in an Indexed space (§8.6.6.3)', () => {
  it('fills with the table entry the index names, snapped into the table', () => {
    // IndexedCS_negative_and_high.pdf: left unread, every swatch kept the
    // colour set before it.
    const file = PdfFile.parse(
      swatches(
        [
          '0 0 0 rg /Cs1 cs',
          '1 sc 0 0 10 10 re f',
          '-17 sc 20 0 10 10 re f',
          '1.6 sc 40 0 10 10 re f',
          '17 sc 60 0 10 10 re f',
        ].join('\n'),
      ),
    );
    const fills = collectPageVectors(file, file.pages()[0]!, []).vectors.map((v) => v.fillHex);
    expect(fills).toEqual(['00FF00', 'FF0000', '0000FF', '0000FF']);
  });
});
