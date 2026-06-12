import { describe, expect, it } from 'vitest';

import { OpcPackage, buildOpcPackage, relsPathFor } from '@/core/opc';

const enc = (s: string) => new TextEncoder().encode(s);
const dec = (b: Uint8Array | undefined) => (b ? new TextDecoder().decode(b) : undefined);

const DOC_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const OFFICE_DOC =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

describe('OPC writer (ECMA-376 Part 2)', () => {
  it('round-trips parts and relationships through OpcPackage.open', () => {
    const bytes = buildOpcPackage({
      parts: [
        { path: 'word/document.xml', data: enc('<document/>'), contentType: DOC_TYPE },
        {
          path: 'word/media/image1.png',
          data: new Uint8Array([1, 2, 3]),
          contentType: 'image/png',
        },
      ],
      rootRelationships: [
        { id: 'rId1', type: OFFICE_DOC, target: 'word/document.xml', targetMode: 'Internal' },
      ],
      partRelationships: [
        {
          sourcePart: 'word/document.xml',
          relationships: [
            { id: 'rId1', type: IMAGE_REL, target: 'media/image1.png', targetMode: 'Internal' },
          ],
        },
      ],
      defaultsByExtension: { png: 'image/png' },
    });

    const pkg = OpcPackage.open(bytes);
    // Main document resolves through the root relationship.
    expect(dec(pkg.getMainDocument().data)).toBe('<document/>');
    // The binary part survives verbatim.
    expect([...pkg.getPart('word/media/image1.png')!]).toEqual([1, 2, 3]);
    // Per-part relationships read back identically.
    const docRels = pkg.getPartRelationships('word/document.xml');
    expect(docRels).toHaveLength(1);
    expect(docRels[0]).toMatchObject({ id: 'rId1', type: IMAGE_REL, target: 'media/image1.png' });
  });

  it('is deterministic — identical input yields identical bytes', () => {
    const make = () =>
      buildOpcPackage({
        parts: [{ path: 'word/document.xml', data: enc('<x/>'), contentType: DOC_TYPE }],
        rootRelationships: [
          { id: 'rId1', type: OFFICE_DOC, target: 'word/document.xml', targetMode: 'Internal' },
        ],
      });
    expect([...make()]).toEqual([...make()]);
  });

  it('Content_Types: binary extensions use Default, XML parts use Override', () => {
    const bytes = buildOpcPackage({
      parts: [
        { path: 'word/document.xml', data: enc('<x/>'), contentType: DOC_TYPE },
        { path: 'word/media/image1.png', data: new Uint8Array([0]), contentType: 'image/png' },
      ],
      rootRelationships: [],
      defaultsByExtension: { png: 'image/png' },
    });
    // [Content_Types].xml is the first stored part — decode it from the package.
    const pkg = OpcPackage.open(bytes);
    const ct = dec(pkg.getPart('[Content_Types].xml'));
    expect(ct).toContain('<Default Extension="png" ContentType="image/png"/>');
    expect(ct).toContain(`<Override PartName="/word/document.xml"`);
    // The PNG, covered by its Default, gets no Override.
    expect(ct).not.toContain('PartName="/word/media/image1.png"');
  });

  it('relsPathFor places the rels part beside its owner', () => {
    expect(relsPathFor('word/document.xml')).toBe('word/_rels/document.xml.rels');
    expect(relsPathFor('word/header1.xml')).toBe('word/_rels/header1.xml.rels');
  });

  it('escapes XML-significant characters in relationship targets', () => {
    const bytes = buildOpcPackage({
      parts: [{ path: 'word/document.xml', data: enc('<x/>'), contentType: DOC_TYPE }],
      rootRelationships: [],
      partRelationships: [
        {
          sourcePart: 'word/document.xml',
          relationships: [
            {
              id: 'rId9',
              type: 'http://x/hyperlink',
              target: 'https://e.com/?a=1&b=2',
              targetMode: 'External',
            },
          ],
        },
      ],
    });
    const pkg = OpcPackage.open(bytes);
    const rels = pkg.getPartRelationships('word/document.xml');
    expect(rels[0]!.target).toBe('https://e.com/?a=1&b=2');
    expect(rels[0]!.targetMode).toBe('External');
  });
});
