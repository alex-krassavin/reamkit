import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { bytesIncludePartName, packageHasPart } from '@/core/bytes';
import { OpcPackage } from '@/core/opc';
import { parseRelationships } from '@/core/opc/relationships';

const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

describe('packageHasPart — what the archive holds, not what its bytes contain', () => {
  const enc = (t: string): Uint8Array => new TextEncoder().encode(t);

  it('finds a part by name and refuses one that is only text inside another', () => {
    // The same trap the corpus decks fell into, in miniature: an entry whose
    // CONTENT names a part of some other package.
    const zip = zipSync({
      'ppt/presentation.xml': enc('<p:presentation/>'),
      // STORED, as PowerPoint writes an embedded workbook: its bytes go in
      // verbatim, so what they say is readable from the outside.
      'ppt/embeddings/book.xlsx': [enc('PK...xl/workbook.xml...'), { level: 0 }],
    });
    expect(packageHasPart(zip, 'ppt/presentation.xml')).toBe(true);
    expect(bytesIncludePartName(zip, 'xl/workbook.xml')).toBe(true); // the old probe
    expect(packageHasPart(zip, 'xl/workbook.xml')).toBe(false);
  });

  it('matches a directory prefix, the backslash spelling, and any case', () => {
    const zip = zipSync({ 'word/_rels/document.xml.rels': enc('<Relationships/>') });
    expect(packageHasPart(zip, 'word/_rels/')).toBe(true);
    expect(packageHasPart(zip, 'word/')).toBe(true);
    expect(packageHasPart(zip, 'xl/')).toBe(false);
    // OPC compares part names without case (§9.1.1.1), as does the reader.
    const odd = zipSync({ 'Word\\Document.xml': enc('<w:document/>') });
    expect(packageHasPart(odd, 'word/document.xml')).toBe(true);
  });

  it('falls back to the byte probe when there is no directory to read', () => {
    // Truncated to the first entry: no end record, so nothing to walk. The old
    // answer stands and the reader behind the sniff reports the real problem.
    const zip = zipSync({ 'xl/workbook.xml': enc('<workbook/>') });
    // Enough for the first local header — which carries the entry's name.
    const truncated = zip.subarray(0, 60);
    expect(packageHasPart(truncated, 'xl/workbook.xml')).toBe(true);
    expect(packageHasPart(new Uint8Array(8), 'xl/workbook.xml')).toBe(false);
  });

  it('reads the directory of a real package', () => {
    const docx = buildDocxFromBody('<w:p/>');
    expect(packageHasPart(docx, 'word/document.xml')).toBe(true);
    expect(packageHasPart(docx, '[Content_Types].xml')).toBe(true);
    expect(packageHasPart(docx, 'ppt/presentation.xml')).toBe(false);
  });
});

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

  it('refuses a corrupt entry by name instead of leaking the inflater', () => {
    // forcepoint107.xlsx truncates a deflate stream. fflate threw `RangeError:
    // offset is out of bounds` straight through us — a message that says
    // nothing about the document, and an unhandled crash where every other
    // malformed archive here gets a refusal.
    const good = zipSync({ 'a.bin': new Uint8Array(4096) });
    // Chop the payload but keep the directory, so the entry inflates short.
    const truncated = good.slice(0, Math.floor(good.length / 2));
    expect(() => OpcPackage.open(truncated)).toThrow(/invalid zip data|zip/i);
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

describe('OpcPackage.open — Windows-style backslash entry names', () => {
  // APPNOTE §4.4.17.1 requires forward slashes, but Windows producers write
  // "_rels\.rels" and "xl\workbook.xml" anyway. Excel, POI and LibreOffice all
  // normalize; we used to reject the package outright as missing its root
  // relationships. Corpus: tdf131575.xlsx, tdf76115.xlsx, 49609.xlsx.
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const enc = (s: string): Uint8Array => new TextEncoder().encode(s);

  const backslashPackage = (): Uint8Array =>
    zipSync({
      '[Content_Types].xml': enc(
        `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
      ),
      '_rels\\.rels': enc(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      ),
      'xl\\workbook.xml': enc(`<workbook/>`),
    });

  it('opens a package whose entries use backslashes', () => {
    const pkg = OpcPackage.open(backslashPackage());
    expect(pkg.listParts()).toContain('xl/workbook.xml');
    expect(pkg.getPart('xl/workbook.xml')).toBeDefined();
  });

  it('resolves the main document part through the normalized root rels', () => {
    const pkg = OpcPackage.open(backslashPackage());
    expect(pkg.getMainDocumentPath()).toBe('xl/workbook.xml');
  });
});
describe('part names are case-insensitive (ISO/IEC 29500-2 §9.1.1.1)', () => {
  it('finds a part and its relationships whatever case they were written in', () => {
    // Producers mix them: 123233_charts.xlsx writes `xl/worksheets/Sheet1.xml`
    // and its relationships as `_rels/sheet1.xml.rels`. An exact lookup found no
    // relationships for that sheet at all, which cost it four charts silently —
    // a sheet with no drawing rel is indistinguishable from one with no drawing.
    const enc = new TextEncoder();
    const pkg = OpcPackage.open(
      zipSync({
        '_rels/.rels': enc.encode(
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
        ),
        'xl/worksheets/Sheet1.xml': enc.encode('<worksheet/>'),
        'xl/worksheets/_rels/sheet1.xml.rels': enc.encode(
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://x/drawing" Target="../drawings/drawing1.xml"/>' +
            '</Relationships>',
        ),
      }),
    );
    // The exact name still resolves…
    expect(pkg.getPart('xl/worksheets/Sheet1.xml')).toBeDefined();
    // …and so does one that differs only in case.
    expect(pkg.getPart('xl/worksheets/sheet1.xml')).toBeDefined();
    const rels = pkg.getPartRelationships('xl/worksheets/Sheet1.xml');
    expect(rels.map((r) => r.id)).toEqual(['rId1']);
  });
});
