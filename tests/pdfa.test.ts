import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildTinyPng } from './fixtures/build-png';
import { convertDocxToPdfSync } from '@/core/converter';
import { buildSrgbIccProfile } from '@/pdf/icc-profile';
import { buildXmpPacket } from '@/pdf/xmp';

const here = dirname(fileURLToPath(import.meta.url));
const FONTS = {
  regular: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf'))),
  bold: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Bold.ttf'))),
};
const latin1 = new TextDecoder('latin1');
const asLatin1 = (b: Uint8Array): string => latin1.decode(b);

function tag4(bytes: Uint8Array, offset: number): string {
  return String.fromCharCode(
    bytes[offset]!,
    bytes[offset + 1]!,
    bytes[offset + 2]!,
    bytes[offset + 3]!,
  );
}
function u32(bytes: Uint8Array, offset: number): number {
  return (
    bytes[offset]! * 0x1000000 +
    ((bytes[offset + 1]! << 16) | (bytes[offset + 2]! << 8) | bytes[offset + 3]!)
  );
}

describe('ICC profile generator', () => {
  it('produces a structurally valid ICC v2 RGB display profile', () => {
    const icc = buildSrgbIccProfile();
    // Header: size matches, signature 'acsp', class 'mntr', space 'RGB ', PCS 'XYZ '.
    expect(u32(icc, 0)).toBe(icc.length);
    expect(tag4(icc, 36)).toBe('acsp');
    expect(tag4(icc, 12)).toBe('mntr');
    expect(tag4(icc, 16)).toBe('RGB ');
    expect(tag4(icc, 20)).toBe('XYZ ');
    // Version 2.x.
    expect(icc[8]).toBe(0x02);
  });

  it('contains all required tags for a matrix/TRC RGB profile', () => {
    const icc = buildSrgbIccProfile();
    const count = u32(icc, 128);
    const tags = new Set<string>();
    for (let i = 0; i < count; i++) {
      const p = 132 + i * 12;
      tags.add(tag4(icc, p));
      // Each tag's data must lie within the profile.
      const off = u32(icc, p + 4);
      const size = u32(icc, p + 8);
      expect(off + size).toBeLessThanOrEqual(icc.length);
    }
    for (const required of [
      'desc',
      'wtpt',
      'rXYZ',
      'gXYZ',
      'bXYZ',
      'rTRC',
      'gTRC',
      'bTRC',
      'cprt',
    ]) {
      expect(tags.has(required)).toBe(true);
    }
  });
});

describe('XMP packet', () => {
  it('carries the PDF/A identifier and core metadata', () => {
    const xmp = asLatin1(
      buildXmpPacket({
        pdfaPart: '1',
        pdfaConformance: 'B',
        title: 'Doc',
        author: 'Alice',
        producer: 'Ream',
      }),
    );
    expect(xmp).toContain('<?xpacket begin=');
    expect(xmp).toContain('pdfaid:part>1<');
    expect(xmp).toContain('pdfaid:conformance>B<');
    expect(xmp).toContain('http://www.aiim.org/pdfa/ns/id/');
    expect(xmp).toContain('<dc:title>');
    expect(xmp).toContain('Alice');
    expect(xmp).toContain('<?xpacket end="w"?>');
  });

  it('escapes XML metacharacters in metadata values', () => {
    const xmp = asLatin1(buildXmpPacket({ pdfaPart: '1', title: 'A & B < C' }));
    expect(xmp).toContain('A &amp; B &lt; C');
  });
});

describe('PDF/A-1b end-to-end', () => {
  const docx = () => buildDocxFromBody('<w:p><w:r><w:t>PDF/A document body</w:t></w:r></w:p>');

  it('emits a PDF 1.4 header in PDF/A mode', () => {
    const normal = convertDocxToPdfSync(docx(), { fonts: FONTS });
    const pdfa = convertDocxToPdfSync(docx(), { fonts: FONTS, pdfA: 'PDF/A-1b' });
    expect(asLatin1(normal).startsWith('%PDF-1.7')).toBe(true);
    expect(asLatin1(pdfa).startsWith('%PDF-1.4')).toBe(true);
  });

  it('emits a /ID array in the trailer', () => {
    const pdfa = convertDocxToPdfSync(docx(), { fonts: FONTS, pdfA: 'PDF/A-1b' });
    expect(asLatin1(pdfa)).toMatch(/\/ID \[<[0-9A-F]{32}> <[0-9A-F]{32}>\]/);
  });

  it('embeds an OutputIntent with GTS_PDFA1 and an ICC DestOutputProfile', () => {
    const text = asLatin1(convertDocxToPdfSync(docx(), { fonts: FONTS, pdfA: 'PDF/A-1b' }));
    expect(text).toContain('/OutputIntents');
    expect(text).toContain('/S /GTS_PDFA1');
    expect(text).toContain('/DestOutputProfile');
    expect(text).toMatch(/\/N 3\b/); // ICC stream component count
  });

  it('emits document XMP /Metadata with the pdfaid identifier', () => {
    const text = asLatin1(
      convertDocxToPdfSync(docx(), {
        fonts: FONTS,
        pdfA: 'PDF/A-1b',
        info: { title: 'My Title', author: 'Bob' },
      }),
    );
    expect(text).toContain('/Type /Metadata');
    expect(text).toContain('/Subtype /XML');
    expect(text).toContain('pdfaid:part>1<');
    // /Info and XMP both present and consistent (Title in both).
    expect(text).toMatch(/\/Title <FEFF/);
  });

  it('does not emit PDF/A structures in normal mode', () => {
    const text = asLatin1(convertDocxToPdfSync(docx(), { fonts: FONTS }));
    expect(text).not.toContain('/OutputIntents');
    expect(text).not.toContain('GTS_PDFA1');
    expect(text).not.toContain('/Type /Metadata');
    expect(text).not.toMatch(/\/ID \[</);
  });

  it('flattens PNG alpha (no /SMask) in PDF/A mode', () => {
    const png = buildTinyPng(4, 4, [200, 50, 50, 128]); // semi-transparent
    const drawing = `<w:r><w:drawing>
      <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="304800" cy="304800"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:blipFill><a:blip r:embed="rId20"/></pic:blipFill>
              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="304800" cy="304800"/></a:xfrm></pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline></w:drawing></w:r>`;
    const body = `<w:p>${drawing}</w:p>`;
    const opts = {
      images: {
        rId20: { contentType: 'image/png' as const, bytes: png, extension: 'png' as const },
      },
    };

    const normal = asLatin1(convertDocxToPdfSync(buildDocxFromBody(body, opts), { fonts: FONTS }));
    const pdfa = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(body, opts), { fonts: FONTS, pdfA: 'PDF/A-1b' }),
    );
    // Normal mode keeps the alpha as a soft mask; PDF/A mode flattens it away.
    expect(normal).toContain('/SMask');
    expect(pdfa).not.toContain('/SMask');
    // The image XObject is still drawn.
    expect(pdfa).toMatch(/\/Im\d+ Do/);
  });

  it('subset fonts use a tag prefix (required by PDF/A)', () => {
    const text = asLatin1(convertDocxToPdfSync(docx(), { fonts: FONTS, pdfA: 'PDF/A-1b' }));
    expect(text).toMatch(/\/BaseFont \/[A-Z]{6}\+Roboto/);
    expect(text).toMatch(/\/FontName \/[A-Z]{6}\+Roboto/);
  });

  it('emits a /CIDSet in every subset font descriptor (PDF/A §6.3.5)', () => {
    const text = asLatin1(convertDocxToPdfSync(docx(), { fonts: FONTS, pdfA: 'PDF/A-1b' }));
    // The font descriptor references a CIDSet stream right after its FontFile2.
    expect(text).toMatch(/\/FontFile2 \d+ 0 R \/CIDSet \d+ 0 R/);
  });

  it('produces byte-identical output for identical input (deterministic /ID)', () => {
    const a = convertDocxToPdfSync(docx(), { fonts: FONTS, pdfA: 'PDF/A-1b' });
    const b = convertDocxToPdfSync(docx(), { fonts: FONTS, pdfA: 'PDF/A-1b' });
    expect(asLatin1(a)).toBe(asLatin1(b));
  });
});

describe('PDF/A-2 (ISO 19005-2)', () => {
  const docx = () => buildDocxFromBody('<w:p><w:r><w:t>PDF/A-2 body.</w:t></w:r></w:p>');
  const out = (pdfA: 'PDF/A-2b' | 'PDF/A-2u' | 'PDF/A-2a') =>
    asLatin1(convertDocxToPdfSync(docx(), { fonts: FONTS, pdfA }));

  it('uses a PDF 1.7 header (part 2 is PDF 1.7-based)', () => {
    expect(out('PDF/A-2b').startsWith('%PDF-1.7')).toBe(true);
  });

  it('writes XMP part 2 and the requested conformance level', () => {
    expect(out('PDF/A-2u')).toContain('pdfaid:part>2<');
    expect(out('PDF/A-2u')).toContain('pdfaid:conformance>U<');
    expect(out('PDF/A-2a')).toContain('pdfaid:conformance>A<');
    expect(out('PDF/A-2b')).toContain('pdfaid:conformance>B<');
  });

  it('gives pages a transparency group and omits the optional /CIDSet', () => {
    const text = out('PDF/A-2b');
    // Device-independent (ICCBased sRGB) blend space, required when keeping the
    // image soft mask (PDF/A-2 §6.2.4.3).
    expect(text).toContain('/Group <</S /Transparency /CS [/ICCBased');
    expect(text).not.toContain('/CIDSet');
  });

  it('PDF/A-2a implies a tagged structure tree; -2b/-2u do not', () => {
    expect(out('PDF/A-2a')).toContain('/StructTreeRoot');
    expect(out('PDF/A-2b')).not.toContain('/StructTreeRoot');
  });

  it('keeps the image soft mask (transparency allowed, unlike PDF/A-1)', () => {
    const png = buildTinyPng(4, 4, [200, 50, 50, 128]);
    const drawing = `<w:r><w:drawing>
      <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="304800" cy="304800"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:blipFill><a:blip r:embed="rId20"/></pic:blipFill>
              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="304800" cy="304800"/></a:xfrm></pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline></w:drawing></w:r>`;
    const opts = {
      images: {
        rId20: { contentType: 'image/png' as const, bytes: png, extension: 'png' as const },
      },
    };
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(`<w:p>${drawing}</w:p>`, opts), {
        fonts: FONTS,
        pdfA: 'PDF/A-2b',
      }),
    );
    expect(text).toContain('/SMask');
  });
});

// Byte-subsequence search (for confirming an embedded file's bytes round-trip).
function indexOfBytes(hay: Uint8Array, needle: Uint8Array): number {
  outer: for (let i = 0; i + needle.length <= hay.length; i++) {
    for (let j = 0; j < needle.length; j++) {
      if (hay[i + j] !== needle[j]) continue outer;
    }
    return i;
  }
  return -1;
}

describe('PDF/A-3 associated files (ISO 19005-3)', () => {
  const sourceDocx = () => buildDocxFromBody('<w:p><w:r><w:t>PDF/A-3 body.</w:t></w:r></w:p>');

  it('embeds the source document as an associated /Source file', () => {
    const docx = sourceDocx();
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS, pdfA: 'PDF/A-3b', embedSource: true });
    const text = asLatin1(pdf);
    expect(text).toContain('/Type /EmbeddedFile');
    expect(text).toContain('/Type /Filespec');
    expect(text).toContain('/AFRelationship /Source');
    expect(text).toContain('/AF [');
    expect(text).toContain('/EmbeddedFiles');
    // The MIME type is stored as a (hex-escaped) Name.
    expect(text).toContain('/application#2Fvnd.openxmlformats');
    // The exact source bytes are embedded (uncompressed stream → present verbatim).
    expect(indexOfBytes(pdf, docx)).toBeGreaterThan(0);
  });

  it('does not embed source files for PDF/A-2 (forbidden there)', () => {
    const text = asLatin1(
      convertDocxToPdfSync(sourceDocx(), { fonts: FONTS, pdfA: 'PDF/A-2b', embedSource: true }),
    );
    expect(text).not.toContain('/Type /EmbeddedFile');
    expect(text).not.toContain('/AFRelationship');
  });

  it('embeds a caller-supplied attachment with its relationship', () => {
    const data = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]);
    const text = asLatin1(
      convertDocxToPdfSync(sourceDocx(), {
        fonts: FONTS,
        pdfA: 'PDF/A-3b',
        attachments: [
          {
            name: 'data.bin',
            bytes: data,
            mimeType: 'application/octet-stream',
            relationship: 'Data',
          },
        ],
      }),
    );
    expect(text).toContain('/AFRelationship /Data');
    expect(text).toContain('/Type /EmbeddedFile');
  });

  it('PDF/A-3a still builds the tagged structure tree', () => {
    const text = asLatin1(
      convertDocxToPdfSync(sourceDocx(), {
        fonts: FONTS,
        pdfA: 'PDF/A-3a',
        embedSource: true,
        info: { title: 'Doc' },
      }),
    );
    expect(text).toContain('/StructTreeRoot');
    expect(text).toContain('pdfaid:part>3<');
    expect(text).toContain('/Type /EmbeddedFile');
  });
});
