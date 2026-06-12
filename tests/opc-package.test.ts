import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { OpcPackage } from '@/core/opc';
import { parseRelationships } from '@/core/opc/relationships';

const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

describe('parseRelationships — namespace prefix (corpus: 58760.xlsx)', () => {
  it('parses relationships whose elements carry a namespace prefix', () => {
    const xml =
      `<ns0:Relationships xmlns:ns0="${REL_NS}">` +
      '<ns0:Relationship Id="rId1" Type="http://x/worksheet" Target="sheet1.xml"/>' +
      '<ns0:Relationship Id="rId2" Type="http://x/styles" Target="styles.xml" TargetMode="Internal"/>' +
      '</ns0:Relationships>';
    const rels = parseRelationships(new TextEncoder().encode(xml));
    expect(rels.map((r) => r.id)).toEqual(['rId1', 'rId2']);
    expect(rels[0]).toEqual({
      id: 'rId1',
      type: 'http://x/worksheet',
      target: 'sheet1.xml',
      targetMode: 'Internal',
    });
  });

  it('still parses the default (unprefixed) form', () => {
    const xml = `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="http://x/t" Target="a.xml"/></Relationships>`;
    expect(parseRelationships(new TextEncoder().encode(xml))).toHaveLength(1);
  });
});

describe('OpcPackage.open — zip-bomb hardening', () => {
  it('opens a normal package with the default limits', () => {
    const pkg = OpcPackage.open(buildDocxFromBody('<w:p><w:r><w:t>hi</w:t></w:r></w:p>'));
    expect(pkg.getMainDocumentPath()).toContain('document.xml');
  });

  it('rejects an archive larger than maxArchiveBytes before unzipping', () => {
    expect(() => OpcPackage.open(new Uint8Array(200), { maxArchiveBytes: 100 })).toThrow(
      /too large/,
    );
  });

  it('rejects a single entry over the per-entry uncompressed cap', () => {
    // 2 MiB of zeros compresses to a few bytes — a classic bomb shape.
    const bomb = zipSync({ 'big.bin': new Uint8Array(2 * 1024 * 1024) });
    expect(() => OpcPackage.open(bomb, { maxEntryBytes: 1024 })).toThrow(/zip-bomb guard/);
  });

  it('rejects when total uncompressed size exceeds the cap', () => {
    const bomb = zipSync({ a: new Uint8Array(1024 * 1024), b: new Uint8Array(1024 * 1024) });
    expect(() => OpcPackage.open(bomb, { maxTotalBytes: 4096 })).toThrow(/zip-bomb guard/);
  });

  it('rejects when the entry count exceeds the cap', () => {
    const bomb = zipSync({ a: new Uint8Array(1), b: new Uint8Array(1), c: new Uint8Array(1) });
    expect(() => OpcPackage.open(bomb, { maxEntries: 1 })).toThrow(/zip-bomb guard/);
  });
});
