// ECMA-376 Part 1 §17.8.1 — the faces a document is set in, embedded in its
// package (see `src/word/font-embed.ts`): each family's slots written as
// obfuscated TrueType parts, where the face's licence lets it travel.

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { GlyphSeg } from '@/core/font';
import type { FaceGlyph, FaceOutlines, FlowDoc } from '@/core/ir/flow';
import { parseTtf } from '@/core/font';
import { OpcPackage } from '@/core/opc';
import { writeDocx } from '@/word/docx-writer';
import { readDocx } from '@/word/docx-reader';
import { deobfuscateEmbeddedFont } from '@/word/font-table';

const decode = (b: Uint8Array): string => new TextDecoder().decode(b);

const BAR: ReadonlyArray<GlyphSeg> = [
  { op: 'move', x: 0.1, y: 0 },
  { op: 'line', x: 0.1, y: 0.7 },
  { op: 'line', x: 0.2, y: 0.7 },
  { op: 'line', x: 0.2, y: 0 },
  { op: 'close' },
];

/** A face drawing each of `chars` as a bar, `advance` wide. */
function face(chars: string, options: Partial<FaceOutlines> = {}): FaceOutlines {
  const glyphs = new Map<string, FaceGlyph>(
    [...chars].map((c) => [c, { outline: c === ' ' ? [] : BAR, advance: 500 }]),
  );
  return {
    glyphs,
    bold: false,
    italic: false,
    postScriptName: 'Inter-Regular',
    ascent: 969,
    descent: -241,
    italicAngle: 0,
    fixedPitch: false,
    ...options,
  };
}

/** A one-paragraph document whose runs name the given faces, `bold` where asked. */
function documentIn(
  runs: ReadonlyArray<{ face: string; text: string; bold?: boolean }>,
  outlines: ReadonlyMap<string, FaceOutlines>,
): FlowDoc {
  const { doc } = readDocx(
    buildDocxFromBody(
      `<w:p>${runs
        .map(
          (r) =>
            `<w:r><w:rPr><w:rFonts w:ascii="${r.face}" w:hAnsi="${r.face}"/>${r.bold === true ? '<w:b/>' : ''}</w:rPr>` +
            `<w:t xml:space="preserve">${r.text}</w:t></w:r>`,
        )
        .join('')}</w:p>`,
    ),
  );
  return {
    ...doc,
    faceFamilies: new Map(
      runs.map((r) => [r.face, { family: 'Inter', generic: 'swiss' as const }] as const),
    ),
    faceOutlines: outlines,
  };
}

/** The embedded font a `w:embed*` element points to, de-obfuscated. */
function embeddedFont(bytes: Uint8Array, slot: string): Uint8Array {
  const pkg = OpcPackage.open(bytes);
  const table = decode(pkg.getPart('word/fontTable.xml')!);
  const m = new RegExp(`<w:embed${slot} r:id="([^"]+)" w:fontKey="([^"]+)"`, 'u').exec(table);
  if (!m) throw new Error(`no ${slot} face`);
  const rels = decode(pkg.getPart('word/_rels/fontTable.xml.rels')!);
  const target = new RegExp(`Id="${m[1]!}"[^>]*Target="([^"]+)"`, 'u').exec(rels)?.[1];
  return deobfuscateEmbeddedFont(pkg.getPart(`word/${target!}`)!, m[2]!);
}

describe('a document that embeds the faces it is set in (§17.8.1)', () => {
  it('writes each face as an obfuscated TrueType part the font table points to', () => {
    const flow = documentIn(
      [{ face: 'inter-regular', text: 'Hi' }],
      new Map([['inter-regular', face('Hi ')]]),
    );
    const { bytes, losses } = writeDocx(flow);
    expect(losses).toEqual([]);
    const pkg = OpcPackage.open(bytes);
    const table = decode(pkg.getPart('word/fontTable.xml')!);
    expect(table).toMatch(
      /<w:font w:name="Inter"><w:family w:val="swiss"\/><w:pitch w:val="variable"\/><w:embedRegular r:id="rId1" w:fontKey="\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}" w:subsetted="1"\/><\/w:font>/u,
    );
    expect(decode(pkg.getPart('word/_rels/fontTable.xml.rels')!)).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font1.odttf"',
    );
    expect(decode(pkg.getPart('[Content_Types].xml')!)).toContain(
      '<Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>',
    );
    // §17.15.1.42, .74 — kept embedded, as subsets, when the document is saved again.
    expect(decode(pkg.getPart('word/settings.xml')!)).toContain(
      '<w:embedTrueTypeFonts/><w:saveSubsetFonts/>',
    );
    // Obfuscated, the part is no font; with its key it is the face.
    expect(pkg.getPart('word/fonts/font1.odttf')!.subarray(0, 4)).not.toEqual(
      Uint8Array.from([0, 1, 0, 0]),
    );
    const font = parseTtf(embeddedFont(bytes, 'Regular'));
    expect(font.glyphForCodepoint(0x48)).toBeGreaterThan(0);
    expect(font.glyphForCodepoint(0x69)).toBeGreaterThan(0);
  });

  it('is read back as the family its runs name', () => {
    const flow = documentIn(
      [{ face: 'inter-regular', text: 'Hi' }],
      new Map([['inter-regular', face('Hi ')]]),
    );
    const { doc } = readDocx(writeDocx(flow).bytes);
    expect(doc.embeddedFonts?.has('inter')).toBe(true);
  });

  it('puts a bold face in the bold slot, and two faces of one slot into one font', () => {
    // A SemiBold and a Bold are both written bold: one font, the fuller face
    // first and the other's characters it lacks after.
    const flow = documentIn(
      [
        { face: 'inter-regular', text: 'a' },
        { face: 'inter-semibold', text: 'bc', bold: true },
        { face: 'inter-bold', text: 'd', bold: true },
      ],
      new Map([
        ['inter-regular', face('a ')],
        ['inter-semibold', face('bc ', { bold: true, postScriptName: 'Inter-SemiBold' })],
        ['inter-bold', face('d ', { bold: true, postScriptName: 'Inter-Bold' })],
      ]),
    );
    const bytes = writeDocx(flow).bytes;
    const table = decode(OpcPackage.open(bytes).getPart('word/fontTable.xml')!);
    expect(table.match(/<w:embed/gu)).toHaveLength(2);
    const bold = parseTtf(embeddedFont(bytes, 'Bold'));
    for (const c of 'bcd') expect(bold.glyphForCodepoint(c.codePointAt(0)!)).toBeGreaterThan(0);
    expect(bold.postScriptName).toBe('Inter-SemiBold');
  });

  it('leaves out a face its licence restricts, and says so', () => {
    // OS/2 fsType 2 — Restricted License embedding: the face may not travel.
    const flow = documentIn(
      [{ face: 'inter-regular', text: 'Hi' }],
      new Map([['inter-regular', face('Hi ', { fsType: 2 })]]),
    );
    const { bytes, losses } = writeDocx(flow);
    const pkg = OpcPackage.open(bytes);
    expect(decode(pkg.getPart('word/fontTable.xml')!)).not.toContain('w:embed');
    expect(pkg.getPart('word/fonts/font1.odttf')).toBeUndefined();
    expect(pkg.getPart('word/settings.xml')).toBeUndefined();
    expect(losses).toEqual([
      expect.objectContaining({ feature: 'fonts.embedding', severity: 'degraded' }),
    ]);
  });

  it('leaves out a face allowed only in documents opened read-only', () => {
    const flow = documentIn(
      [{ face: 'inter-regular', text: 'Hi' }],
      new Map([['inter-regular', face('Hi ', { fsType: 4 })]]),
    );
    expect(
      decode(OpcPackage.open(writeDocx(flow).bytes).getPart('word/fontTable.xml')!),
    ).not.toContain('w:embed');
  });

  it('states Editable embedding for a face whose program stated nothing', () => {
    const flow = documentIn(
      [{ face: 'inter-regular', text: 'Hi' }],
      new Map([['inter-regular', face('Hi ')]]),
    );
    const font = embeddedFont(writeDocx(flow).bytes, 'Regular');
    const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
    let os2 = -1;
    for (let i = 0; i < view.getUint16(4); i++) {
      const at = 12 + i * 16;
      if (String.fromCharCode(...font.subarray(at, at + 4)) === 'OS/2')
        os2 = view.getUint32(at + 8);
    }
    expect(view.getUint16(os2 + 8)).toBe(8);
  });

  it('carries the pairs the source kerned the face by, in a kern table', () => {
    const flow = documentIn(
      [{ face: 'inter-regular', text: 'AV' }],
      new Map([['inter-regular', face('AV ', { kerning: new Map([['AV', -80]]) })]]),
    );
    const font = embeddedFont(writeDocx(flow).bytes, 'Regular');
    const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
    let kern = -1;
    for (let i = 0; i < view.getUint16(4); i++) {
      const at = 12 + i * 16;
      if (String.fromCharCode(...font.subarray(at, at + 4)) === 'kern')
        kern = view.getUint32(at + 8);
    }
    expect(kern).toBeGreaterThan(0);
    const parsed = parseTtf(font);
    const [a, v] = ['A', 'V'].map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!));
    expect(view.getUint16(kern + 10)).toBe(1); // one pair
    expect([
      view.getUint16(kern + 18),
      view.getUint16(kern + 20),
      view.getInt16(kern + 22),
    ]).toEqual([a, v, -80]);
  });

  it('writes a run that asks to be kerned with w:kern, where the schema puts it (§17.3.2.19)', () => {
    const { doc } = readDocx(
      buildDocxFromBody(
        '<w:p><w:r><w:rPr><w:color w:val="FF0000"/><w:kern w:val="2"/><w:sz w:val="24"/></w:rPr>' +
          '<w:t>AV</w:t></w:r></w:p>',
      ),
    );
    const bytes = writeDocx(doc).bytes;
    const body = decode(OpcPackage.open(bytes).getMainDocument().data);
    expect(body).toContain('<w:color w:val="FF0000"/><w:kern w:val="2"/><w:sz w:val="24"/>');
    const again = readDocx(bytes).doc.body[0];
    if (again?.kind !== 'paragraph') throw new Error('a paragraph');
    expect(again.paragraph.runs[0]?.properties.kerningMinPt).toBe(1);
  });

  it("carries the ligatures the source drew, as its GSUB's standard ligatures", () => {
    const joined: FaceGlyph = { outline: BAR, advance: 520 };
    const flow = documentIn(
      [{ face: 'inter-regular', text: 'fit' }],
      new Map([['inter-regular', face('fit ', { ligatures: new Map([['fi', joined]]) })]]),
    );
    const parsed = parseTtf(embeddedFont(writeDocx(flow).bytes, 'Regular'));
    const [f, i] = ['f', 'i'].map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!));
    const glyph = parsed.ligatures.get(`${String(f)},${String(i)}`);
    expect(glyph).toBeDefined();
    expect(parsed.advanceWidths[glyph!]).toBe(520);
  });

  it('writes a run set with its ligatures as w14:ligatures, in a part a reader may pass it over in', () => {
    // [MS-DOCX] — Word 2010's own element, after the base schema's, in a part
    // that declares its namespace ignorable (ECMA-376 Part 3 `mc:Ignorable`).
    const { doc } = readDocx(
      buildDocxFromBody(
        '<w:p><w:r><w:rPr><w:lang w:val="en-US"/><w14:ligatures w14:val="standard"/></w:rPr>' +
          '<w:t>fit</w:t></w:r></w:p>',
      ),
    );
    const bytes = writeDocx(doc).bytes;
    const body = decode(OpcPackage.open(bytes).getMainDocument().data);
    expect(body).toContain('<w:lang w:val="en-US"/><w14:ligatures w14:val="standard"/></w:rPr>');
    expect(body).toMatch(/<w:document [^>]*xmlns:w14="[^"]+"[^>]* mc:Ignorable="w14"/u);
    const again = readDocx(bytes).doc.body[0];
    if (again?.kind !== 'paragraph') throw new Error('a paragraph');
    expect(again.paragraph.runs[0]?.properties.ligatures).toBe('standard');
  });

  it('declares no w14 in a part that uses none', () => {
    const { doc } = readDocx(buildDocxFromBody('<w:p><w:r><w:t>fit</w:t></w:r></w:p>'));
    const body = decode(OpcPackage.open(writeDocx(doc).bytes).getMainDocument().data);
    expect(body).not.toContain('xmlns:w14');
    expect(body).not.toContain('mc:Ignorable');
  });

  it('writes the same package from the same document', () => {
    const flow = documentIn(
      [{ face: 'inter-regular', text: 'Hi' }],
      new Map([['inter-regular', face('Hi ')]]),
    );
    expect(writeDocx(flow).bytes).toEqual(writeDocx(flow).bytes);
  });
});
