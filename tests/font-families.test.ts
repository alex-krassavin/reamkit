// E-FONT F2 — which substitute families a document needs, read off the parsed
// model rather than one format's XML. Before this the answer came from a regex
// over `w:rFonts`, so every non-docx input was rendered in ONE family: a deck
// whose theme is Times and whose body is Calibri came out entirely in the
// default sans, and no per-run resolution could help because only one registry
// was ever built.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { Ream } from '@/core/converter/ream';
import { familiesInFlow } from '@/core/fonts/families';
import { resolveFamilyStyle } from '@/core/fonts';

const families = (bytes: Uint8Array): Array<string> =>
  [...familiesInFlow(Ream.parse(bytes).flow)].sort();

/** A workbook whose Normal style names `face`. */
const workbook = (face: string): Uint8Array =>
  buildXlsx({
    rows: [['hello']],
    stylesXml:
      `<fonts count="1"><font><sz val="11"/><name val="${face}"/></font></fonts>` +
      `<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
      `<borders count="1"><border/></borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>`,
  });

describe('the families a document names (E-FONT F2)', () => {
  it('reads a workbook’s own typeface', () => {
    // §18.8.29 `<name>` — carried for every other property of the font and not
    // for this one, so a Calibri workbook asked for no family at all.
    expect(families(workbook('Calibri'))).toEqual(['arimo', 'carlito']);
    expect(families(workbook('Times New Roman'))).toEqual(['arimo', 'tinos']);
  });

  it('always keeps the default sans as the fallback', () => {
    // Unstyled runs, math and chart labels resolve through it.
    expect(families(workbook('Arial'))).toEqual(['arimo']);
  });
});

describe('a family name that also names the face (E-FONT F3)', () => {
  it('reads the weight off the end of the name', () => {
    // `Times New Roman Bold` matched no twin whole, so a Times heading fell
    // through the serif test to the generic sans.
    expect(resolveFamilyStyle('Times New Roman Bold')).toEqual({ key: 'tinos', bold: true });
    expect(resolveFamilyStyle('Arial Black')).toEqual({ key: 'arimo', bold: true });
    expect(resolveFamilyStyle('CenturySchoolbook-Bold')).toMatchObject({ bold: true });
    expect(resolveFamilyStyle('Arial,Bold')).toMatchObject({ bold: true });
    expect(resolveFamilyStyle('Frutiger 45 Light')).toEqual({ key: 'arimo' });
    // No curated family has a condensed cut, so the substitute is squeezed to
    // the width the named one sets at.
    expect(resolveFamilyStyle('Arial Narrow')).toEqual({ key: 'arimo', widthScale: 0.82 });
  });

  it('keeps a family NAMED for a face word whole', () => {
    // Only a trailing word is the face: Book Antiqua is a family.
    expect(resolveFamilyStyle('Book Antiqua')).toEqual({ key: 'tinos' });
    expect(resolveFamilyStyle('Bold')).toEqual({ key: 'arimo' });
  });

  it('finds the family behind a foundry prefix', () => {
    expect(resolveFamilyStyle('Adobe Garamond Pro Bold')).toEqual({ key: 'tinos', bold: true });
  });
});

describe('a face the font set does not carry (E-FONT F3)', () => {
  it('draws the weight and the slant the file has no cut for', async () => {
    // `FontBytesByVariant` requires only `regular`: a caller who supplies one
    // file still asks for bold headings and italic quotes, and every one of
    // them was drawn plain.
    const styled = buildXlsx({
      rows: [
        [
          { value: 'plain', styleIndex: 0 },
          { value: 'bold', styleIndex: 1 },
          { value: 'italic', styleIndex: 2 },
        ],
      ],
      stylesXml:
        `<fonts count="3"><font><sz val="11"/><name val="Arial"/></font>` +
        `<font><b/><sz val="11"/><name val="Arial"/></font>` +
        `<font><i/><sz val="11"/><name val="Arial"/></font></fonts>` +
        `<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
        `<borders count="1"><border/></borders>` +
        `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
        `<cellXfs count="3">` +
        `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
        `<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>` +
        `<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>`,
    });
    const regular = new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf'));
    const pdf = await Ream.parse(styled).convert('pdf', { fonts: { regular } });
    const text = Buffer.from(pdf).toString('latin1');
    // Rendering mode 2 strokes the glyphs as well as filling them…
    expect(text).toMatch(/2 Tr [\d.]+ w/u);
    // …and is cleared again for the runs that carry their own face.
    expect(text).toContain('0 Tr');
    // …while the italic leans through the text matrix's shear term.
    expect(text).toMatch(/1 0 0\.21\d* 1 [\d.]+ [\d.]+ Tm/u);
  });

  it('squeezes a condensed face the substitute has no cut for', async () => {
    const narrow = buildXlsx({
      rows: [
        [
          { value: 'narrow', styleIndex: 0 },
          { value: 'normal', styleIndex: 1 },
        ],
      ],
      stylesXml:
        `<fonts count="2"><font><sz val="11"/><name val="Arial Narrow"/></font>` +
        `<font><sz val="11"/><name val="Arial"/></font></fonts>` +
        `<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
        `<borders count="1"><border/></borders>` +
        `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
        `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
        `<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>`,
    });
    const regular = new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf'));
    const pdf = await Ream.parse(narrow).convert('pdf', { fonts: { regular } });
    const text = Buffer.from(pdf).toString('latin1');
    expect(text).toContain('82 Tz');
    // …and the next run is set at its own width again.
    expect(text).toContain('100 Tz');
  });

  it('asks for nothing when the set has the face', async () => {
    const plain = buildXlsx([['hello']]);
    const four = {
      regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
      bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
      italic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Italic.ttf')),
      boldItalic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-BoldItalic.ttf')),
    };
    const pdf = await Ream.parse(plain).convert('pdf', { fonts: four });
    expect(Buffer.from(pdf).toString('latin1')).not.toContain(' Tr');
  });
});
