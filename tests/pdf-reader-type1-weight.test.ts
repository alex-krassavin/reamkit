// §9.9 — the weight a Type 1 program names in its FontInfo. pdfTeX states no
// /FontWeight in the descriptor and names URW's bold Times
// "NimbusRomNo9L-Medi", which no rule of names reads as bold: comments.pdf's
// headings and captions came back in the weight of its body.

import { describe, expect, it } from 'vitest';

import type { PdfValue } from '@/pdf/objects';
import { PdfFile } from '@/pdf-reader/document';
import { extractPageText } from '@/pdf-reader/text';
import { dict, name, stream } from '@/pdf/objects';
import { PdfDocument } from '@/pdf/writer';

/** Whether a run set in a Type 1 face `base`, whose program names `weight`, reads as bold. */
function boldOf(base: string, weight: string): boolean | undefined {
  const doc = new PdfDocument();
  // The cleartext a Type 1 program opens with; nothing past it is read for this.
  const program = doc.add(
    stream(
      { Length1: 200, Length2: 0, Length3: 0 },
      new TextEncoder().encode(
        `%!PS-AdobeFont-1.0: ${base} 1.05\n/FontInfo 10 dict dup begin\n` +
          `/Weight (${weight}) readonly def\n/ItalicAngle 0 def\nend readonly def\n` +
          `/FontName /${base} def\ncurrentfile eexec\n`,
      ),
    ),
  );
  const descriptor = doc.add(
    dict({
      Type: name('FontDescriptor'),
      FontName: name(base),
      Flags: 4,
      ItalicAngle: 0,
      FontFile: program,
    }),
  );
  const font = doc.add(
    dict({
      Type: name('Font'),
      Subtype: name('Type1'),
      BaseFont: name(`TACTGM+${base}`),
      FontDescriptor: descriptor,
    }),
  );
  const pagesMap = dict({ Type: name('Pages'), Kids: [], Count: 1 });
  const pagesRef = doc.add(pagesMap);
  const content = doc.add(
    stream({}, new TextEncoder().encode('BT /F1 9 Tf 72 700 Td (Figure 13.) Tj ET')),
  );
  const leaf = doc.add(
    dict({
      Type: name('Page'),
      Parent: pagesRef,
      MediaBox: [0, 0, 612, 792],
      Resources: dict({ Font: dict({ F1: font }) }),
      Contents: content,
    }),
  );
  (pagesMap.get('Kids') as Array<PdfValue>).push(leaf);
  const file = PdfFile.parse(doc.build(doc.add(dict({ Type: name('Catalog'), Pages: pagesRef }))));
  return extractPageText(file, file.pages()[0]!)[0]?.bold;
}

describe('the weight a Type 1 program names (§9.9)', () => {
  it('makes a face bold that its name does not', () => {
    expect(boldOf('NimbusRomNo9L-Medi', 'Bold')).toBe(true);
  });

  it('leaves a book weight as it is, whatever it is called', () => {
    // Computer Modern names its book weight "Medium".
    expect(boldOf('CMR9', 'Medium')).toBeUndefined();
    expect(boldOf('NimbusRomNo9L-Regu', 'Regular')).toBeUndefined();
  });
});
