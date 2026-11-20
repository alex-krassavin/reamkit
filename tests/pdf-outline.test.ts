// #138 — the PDF document outline (ISO 32000-1 §12.3.3) built from headings.
// Every document here is synthetic; the expectations come from LibreOffice's
// export and Word's outline levels, measured on a probe (2026-11-20).

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { ReamConvertOptions } from '@/core/converter/ream';
import type { PdfDict, PdfValue } from '@/pdf/objects';
import { Ream } from '@/core/converter/ream';
import { PdfHexString, PdfName, PdfRef } from '@/pdf/objects';
import { PdfFile } from '@/pdf-reader/document';
import { pdfTextString } from '@/pdf-reader/encodings';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const style = (id: string, name: string, lvl?: number): string =>
  `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/>` +
  `<w:basedOn w:val="Normal"/>` +
  (lvl !== undefined ? `<w:pPr><w:outlineLvl w:val="${String(lvl)}"/></w:pPr>` : '') +
  `<w:rPr><w:b/></w:rPr></w:style>`;

// The fixture wraps these in `w:styles` itself.
const STYLES =
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  style('Heading1', 'heading 1', 0) +
  style('Heading2', 'heading 2', 1) +
  style('Heading3', 'heading 3', 2) +
  // Word's own Heading 4 as some writers leave it: named, with no level.
  style('Heading4', 'heading 4') +
  style('Title', 'Title');

const esc = (t: string): string => t.replace(/&/g, '&amp;').replace(/</g, '&lt;');
const para = (styleId: string | undefined, text: string, extraPPr = ''): string =>
  `<w:p><w:pPr>${styleId ? `<w:pStyle w:val="${styleId}"/>` : ''}${extraPPr}</w:pPr>` +
  (text === '' ? '' : `<w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`) +
  '</w:p>';
const body = (n = 1): string => para(undefined, 'Body text. '.repeat(20 * n));

interface Item {
  readonly title: string;
  readonly depth: number;
  readonly count: number | undefined;
  readonly page: number;
  readonly dest: ReadonlyArray<PdfValue>;
  readonly rawTitle: PdfValue;
}

/** Read the outline back the way a viewer walks it: First → Next, depth first. */
function readOutline(
  bytes: Uint8Array,
  password = '',
): { file: PdfFile; root: PdfDict | undefined; items: Array<Item> } {
  const file = PdfFile.parse(bytes, password);
  const root = file.get(file.catalog, 'Outlines');
  if (!(root instanceof Map)) return { file, root: undefined, items: [] };
  const pages = file.pages().map((p) => p.dict);
  const items: Array<Item> = [];
  const walk = (first: PdfValue | undefined, depth: number, parentRef: PdfValue): void => {
    let ref: PdfValue | undefined = first;
    let prev: PdfValue | undefined;
    while (ref instanceof PdfRef) {
      const node = file.resolve(ref);
      if (!(node instanceof Map)) throw new Error('outline item is no dictionary');
      // §12.3.3 Table 153 — the links a viewer relies on.
      expect((node.get('Parent') as PdfRef).id).toBe((parentRef as PdfRef).id);
      if (prev === undefined) expect(node.has('Prev')).toBe(false);
      else expect((node.get('Prev') as PdfRef).id).toBe((prev as PdfRef).id);
      const raw = node.get('Title')!;
      const bytesOf =
        raw instanceof PdfHexString
          ? raw.bytes
          : Uint8Array.from(raw as string, (c) => c.charCodeAt(0) & 0xff);
      const dest = file.get(node, 'Dest') as Array<PdfValue>;
      const count = node.get('Count');
      items.push({
        title: pdfTextString(bytesOf),
        depth,
        count: typeof count === 'number' ? count : undefined,
        page: pages.indexOf(file.resolve(dest[0]!) as PdfDict),
        dest,
        rawTitle: raw,
      });
      if (node.has('First')) walk(node.get('First'), depth + 1, ref);
      prev = ref;
      ref = node.get('Next');
    }
  };
  walk(root.get('First'), 0, file.catalog.get('Outlines')!);
  return { file, root, items };
}

const convert = (bodyXml: string, options: ReamConvertOptions = {}): Promise<Uint8Array> =>
  Ream.parse(buildDocxFromBody(bodyXml, { stylesXml: STYLES })).convert('pdf', {
    fonts: FONTS,
    ...options,
  });

describe('PDF document outline from headings (#138)', () => {
  it('nests the headings as LibreOffice does, all open', async () => {
    const xml =
      para('Heading2', 'Before any chapter') +
      body() +
      para('Heading1', 'Introduction') +
      body() +
      para('Heading2', 'Background') +
      para('Heading3', 'Details') +
      para('Heading1', 'Skip') +
      // A level skipped: Heading 3 straight under Heading 1 is its child.
      para('Heading3', 'Three under one') +
      para('Heading4', 'Four by its name');
    const { root, items } = readOutline(await convert(xml));
    expect(items.map((i) => [i.depth, i.title, i.count])).toEqual([
      [0, 'Before any chapter', undefined],
      [0, 'Introduction', 2],
      [1, 'Background', 1],
      [2, 'Details', undefined],
      [0, 'Skip', 2],
      [1, 'Three under one', 1],
      [2, 'Four by its name', undefined],
    ]);
    // Table 152 — the outline's /Count is every visible item.
    expect(root!.get('Count')).toBe(7);
    expect(root!.get('Type')).toEqual(new PdfName('Outlines'));
  });

  it('jumps to the top of each heading, on its own page, keeping the zoom', async () => {
    const xml =
      para('Heading1', 'First') +
      body(2) +
      para('Heading1', 'Second', '<w:pageBreakBefore/>') +
      body();
    const { file, items } = readOutline(await convert(xml));
    expect(items.map((i) => i.page)).toEqual([0, 1]);
    for (const item of items) {
      const [, kind, left, top, zoom] = item.dest;
      expect(kind).toEqual(new PdfName('XYZ'));
      // §12.3.2.2 Table 151 — zoom 0 leaves the reader's magnification as it is.
      expect(zoom).toBe(0);
      expect(left).toBeGreaterThanOrEqual(0);
      const [, y0, , y1] = file.pages()[item.page]!.mediaBox;
      expect(top as number).toBeLessThanOrEqual(y1);
      expect(top as number).toBeGreaterThan((y0 + y1) / 2);
    }
    // Each heading opens its page, so both stand at the same top margin.
    expect(items[1]!.dest[3]).toBe(items[0]!.dest[3]);
  });

  it('writes each title as a text string — Kazakh, CJK, RTL and astral text intact', async () => {
    const titles = [
      'Plain ASCII (with parens) \\ and a backslash',
      'Қазақ тілі — ә ғ қ ң ө ұ ү һ і',
      '中文标题 日本語の見出し',
      'العنوان الرئيسي',
      'כותרת בעברית',
      'Emoji 😀 𝐁old',
      '“Quotes” – dash € euro',
    ];
    const { items } = readOutline(await convert(titles.map((t) => para('Heading1', t)).join('')));
    expect(items.map((i) => i.title)).toEqual(titles);
    // §7.9.2.2 — printable ASCII as a literal, the rest as UTF-16BE behind FE FF.
    expect(typeof items[0]!.rawTitle).toBe('string');
    for (const item of items.slice(1)) {
      expect(item.rawTitle).toBeInstanceOf(PdfHexString);
      const bytes = (item.rawTitle as PdfHexString).bytes;
      expect([bytes[0], bytes[1]]).toEqual([0xfe, 0xff]);
    }
  });

  it('titles a heading by its words on one line, its list number included', async () => {
    const xml =
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' +
      '<w:r><w:t>Tab</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>and</w:t></w:r>' +
      '<w:r><w:br/></w:r><w:r><w:t xml:space="preserve">break  soft</w:t></w:r>' +
      '<w:r><w:softHyphen/></w:r><w:r><w:t>hyphen</w:t></w:r></w:p>';
    const { items } = readOutline(await convert(xml));
    expect(items.map((i) => i.title)).toEqual(['Tab and break softhyphen']);
  });

  it('leaves out what is no heading, and headings with nothing to say', async () => {
    const xml =
      para('Title', 'The title style') +
      para(undefined, 'Body') +
      para('Heading1', '') +
      para('Heading1', '   ') +
      // §17.3.1.20 — level 9 is body text, stated; on a style that is no
      // built-in heading it is taken at its word.
      para(undefined, 'Stated body text', '<w:outlineLvl w:val="9"/>') +
      para(undefined, 'A direct level', '<w:outlineLvl w:val="0"/>') +
      // …and on a built-in heading it is not: Word reads this as level 2.
      para('Heading2', 'Heading 2 says body text', '<w:outlineLvl w:val="9"/>');
    const { items } = readOutline(await convert(xml));
    expect(items.map((i) => [i.depth, i.title])).toEqual([
      [0, 'A direct level'],
      [1, 'Heading 2 says body text'],
    ]);
  });

  it('takes a heading in a table cell, and none from the page header', async () => {
    const cell = (inner: string): string =>
      `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
    const xml =
      para('Heading1', 'Before the table') +
      '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
      `<w:tr>${cell(para('Heading2', 'In a cell'))}${cell(para(undefined, 'beside'))}</w:tr>` +
      '</w:tbl>' +
      para('Heading1', 'After the table') +
      '<w:sectPr><w:headerReference w:type="default" r:id="rId10"/></w:sectPr>';
    const bytes = await Ream.parse(
      buildDocxFromBody(xml, {
        stylesXml: STYLES,
        headerXml: para('Heading1', 'In the header'),
      }),
    ).convert('pdf', { fonts: FONTS });
    const { items } = readOutline(bytes);
    expect(items.map((i) => [i.depth, i.title])).toEqual([
      [0, 'Before the table'],
      [1, 'In a cell'],
      [0, 'After the table'],
    ]);
    // The cell's item points between the two headings around the table.
    const tops = items.map((i) => i.dest[3] as number);
    expect(tops[0]).toBeGreaterThan(tops[1]!);
    expect(tops[1]).toBeGreaterThan(tops[2]!);
  });

  it('writes no outline for a document without headings, nor when asked for none', async () => {
    expect(readOutline(await convert(body())).root).toBeUndefined();
    const xml = para('Heading1', 'A heading');
    expect(readOutline(await convert(xml, { outline: false })).root).toBeUndefined();
  });

  it('cuts at maxLevel, and opens the panel only when asked', async () => {
    const xml = para('Heading1', 'One') + para('Heading2', 'Two') + para('Heading3', 'Three');
    const cut = readOutline(await convert(xml, { outline: { maxLevel: 2 } }));
    expect(cut.items.map((i) => i.title)).toEqual(['One', 'Two']);
    // §7.7.2 Table 28 — the default leaves the viewer's own choice alone.
    expect(cut.file.catalog.has('PageMode')).toBe(false);
    const panel = readOutline(await convert(xml, { outline: { openPanel: true } }));
    expect(panel.file.catalog.get('PageMode')).toEqual(new PdfName('UseOutlines'));
    expect(panel.items).toHaveLength(3);
  });

  it('survives encryption and PDF/A', async () => {
    const xml = para('Heading1', 'Қазақ тілі') + para('Heading2', 'Plain');
    const locked = await convert(xml, { encrypt: { userPassword: 'pw' } });
    expect(readOutline(locked, 'pw').items.map((i) => i.title)).toEqual(['Қазақ тілі', 'Plain']);
    const archival = await convert(xml, { pdfA: 'PDF/A-2b' });
    expect(readOutline(archival).items.map((i) => i.title)).toEqual(['Қазақ тілі', 'Plain']);
  });
});
