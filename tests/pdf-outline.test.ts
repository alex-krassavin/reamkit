// #138 — the PDF document outline (ISO 32000-1 §12.3.3) built from headings.
// Every document here is synthetic; the expectations come from LibreOffice's
// export and Word's outline levels, measured on a probe (2026-11-20).

import { readFileSync } from 'node:fs';

import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildPptx } from './fixtures/build-pptx';
import { buildXlsx } from './fixtures/build-xlsx';
import type { ReamConvertOptions } from '@/core/converter/ream';
import type { PdfDict, PdfValue } from '@/pdf/objects';
import { Ream } from '@/core/converter/ream';
import { PDF_NULL, PdfHexString, PdfName, PdfRef, dict, name, stream } from '@/pdf/objects';
import { PdfFile } from '@/pdf-reader/document';
import { pdfTextString } from '@/pdf-reader/encodings';
import { readPdfOutline } from '@/pdf-reader/outline';
import { extractPageText } from '@/pdf-reader/text';
import { PdfDocument } from '@/pdf/writer';

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
  /** The page its /Dest jumps to; -1 for an item with none. */
  readonly page: number;
  readonly dest: ReadonlyArray<PdfValue>;
  /** The /URI of its URI action (§12.6.4.7), as written. */
  readonly uri: string | undefined;
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
      // Table 153: /Dest or /A, never both.
      expect(node.has('Dest') && node.has('A')).toBe(false);
      const dest = node.has('Dest') ? (file.get(node, 'Dest') as Array<PdfValue>) : undefined;
      const action = file.get(node, 'A');
      if (action instanceof Map) expect((file.get(action, 'S') as PdfName).value).toBe('URI');
      const count = node.get('Count');
      items.push({
        title: pdfTextString(bytesOf),
        depth,
        count: typeof count === 'number' ? count : undefined,
        page: dest ? pages.indexOf(file.resolve(dest[0]!) as PdfDict) : -1,
        dest: dest ?? [],
        uri: action instanceof Map ? (file.get(action, 'URI') as string) : undefined,
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

describe('PDF document outline for workbooks and decks (#138)', () => {
  /**
   * What every page shows — each piece of text and where it stands — to show
   * an outline adds nothing to the pages. (The first sheet's carrier paragraph
   * does add one `Tm` that sets nothing; the text is what is compared.)
   */
  const shown = (bytes: Uint8Array): Array<Array<string>> => {
    const file = PdfFile.parse(bytes);
    return file
      .pages()
      .map((p) =>
        extractPageText(file, p).map((r) => `${r.text}@${r.x.toFixed(2)},${r.y.toFixed(2)}`),
      );
  };

  it("names each sheet, as LibreOffice's export does, at its first page", async () => {
    const xlsx = buildXlsx({
      sheets: [
        { name: 'Итоги 2026', rows: [['first'], ['sheet']] },
        { name: 'Data', rows: [['second']] },
      ],
    });
    const bytes = await Ream.parse(xlsx).convert('pdf', { fonts: FONTS });
    const { items } = readOutline(bytes);
    expect(items.map((i) => [i.depth, i.title, i.page])).toEqual([
      [0, 'Итоги 2026', 0],
      [0, 'Data', 1],
    ]);
    // The entries ride paragraphs of no height: the pages are what they were.
    const plain = await Ream.parse(xlsx).convert('pdf', { fonts: FONTS, outline: false });
    expect(readOutline(plain).root).toBeUndefined();
    expect(shown(bytes)).toEqual(shown(plain));
  });

  it('titles each slide by its title placeholder, else by its number', async () => {
    const title = (type: string, paragraphs: ReadonlyArray<string>): string =>
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/>' +
      `<p:nvPr><p:ph type="${type}"/></p:nvPr></p:nvSpPr>` +
      '<p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr>' +
      '<p:txBody><a:bodyPr/>' +
      paragraphs.map((t) => `<a:p><a:r><a:rPr lang="en-US"/><a:t>${t}</a:t></a:r></a:p>`).join('') +
      '</p:txBody></p:sp>';
    const pptx = buildPptx(
      [
        title('ctrTitle', ['Қазақ тілі']),
        '',
        // A hidden slide prints no page and takes no entry.
        title('title', ['Hidden']),
        `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${title('title', ['Line one', 'Line two'])}</p:grpSp>`,
      ],
      { hiddenSlides: [false, false, true, false] },
    );
    const { items } = readOutline(await Ream.parse(pptx).convert('pdf', { fonts: FONTS }));
    expect(items.map((i) => [i.depth, i.title, i.page])).toEqual([
      [0, 'Қазақ тілі', 0],
      [0, 'Slide 2', 1],
      [0, 'Line one Line two', 2],
    ]);
  });
});

describe("a source PDF's own outline, read and carried to the PDF written from it (#138)", () => {
  const utf16 = (s: string): PdfHexString => {
    const out = [0xfe, 0xff];
    for (let i = 0; i < s.length; i++) out.push(s.charCodeAt(i) >> 8, s.charCodeAt(i) & 0xff);
    return new PdfHexString(Uint8Array.from(out));
  };

  /**
   * Three pages of Helvetica text and an outline that uses every way an item
   * can say where it jumps: an explicit /Dest, a go-to action naming a string
   * destination in the /Names tree (under a Kids node), a name looked up in the
   * catalog's /Dests, and none at all; one item closed, one title behind a
   * language escape, and a /Next that runs back to the first item.
   */
  function sourcePdf(): Uint8Array {
    const doc = new PdfDocument();
    const font = doc.add(
      dict({ Type: name('Font'), Subtype: name('Type1'), BaseFont: name('Helvetica') }),
    );
    const pagesMap = dict({ Type: name('Pages'), Kids: [], Count: 3 });
    const pagesRef = doc.add(pagesMap);
    const texts = [
      [
        [700, 'Chapter One'],
        [600, 'Section One A'],
      ],
      [
        [700, 'Chapter Two'],
        [500, 'Section Two A'],
      ],
      [[700, 'Appendix text']],
    ] as const;
    const pageRefs = texts.map((lines) => {
      const ops = ['BT /F1 12 Tf', ...lines.map(([y, t]) => `1 0 0 1 72 ${y} Tm (${t}) Tj`), 'ET'];
      const content = doc.add(stream({}, new TextEncoder().encode(ops.join('\n'))));
      const page = doc.add(
        dict({
          Type: name('Page'),
          Parent: pagesRef,
          MediaBox: [0, 0, 612, 792],
          Resources: dict({ Font: dict({ F1: font }) }),
          Contents: content,
        }),
      );
      (pagesMap.get('Kids') as Array<PdfValue>).push(page);
      return page;
    });
    const root = dict({ Type: name('Outlines') });
    const rootRef = doc.add(root);
    const item = (entries: Record<string, PdfValue>): { d: PdfDict; r: PdfRef } => {
      const d = dict({ ...entries, Parent: rootRef });
      return { d, r: doc.add(d) };
    };
    const ch1 = item({ Title: 'Chapter One', Dest: [pageRefs[0]!, name('XYZ'), 0, 700, 0] });
    const s1a = item({
      Title: 'Section One A',
      A: dict({ S: name('GoTo'), D: 'sec1a' }),
    });
    const ch2 = item({ Title: 'Chapter Two', Dest: name('Ch2') });
    const s2a = item({ Title: 'Section Two A', Dest: [pageRefs[1]!, name('XYZ'), 0, 500, 0] });
    const group = item({ Title: 'Grouping only' });
    // ESC "ru" ESC — §7.9.2.2's language escape, which no panel shows.
    const app = item({
      Title: utf16('\u001bru\u001bПриложение'),
      Dest: [pageRefs[2]!, name('Fit')],
    });
    const link = (parent: { d: PdfDict; r: PdfRef }, kids: Array<{ d: PdfDict; r: PdfRef }>) => {
      parent.d.set('First', kids[0]!.r);
      parent.d.set('Last', kids[kids.length - 1]!.r);
      kids.forEach((k, i) => {
        k.d.set('Parent', parent.r);
        if (i > 0) k.d.set('Prev', kids[i - 1]!.r);
        if (i < kids.length - 1) k.d.set('Next', kids[i + 1]!.r);
      });
    };
    link({ d: root, r: rootRef }, [ch1, ch2, group]);
    link(ch1, [s1a]);
    link(ch2, [s2a]);
    link(group, [app]);
    ch1.d.set('Count', 1);
    ch2.d.set('Count', -1); // closed
    group.d.set('Count', 1);
    root.set('Count', 5);
    // A /Next back to the first item: a cycle the reader must not follow.
    group.d.set('Next', ch1.r);
    const leaf = doc.add(
      dict({ Names: ['sec1a', dict({ D: [pageRefs[0]!, name('FitH'), 600] })] }),
    );
    const destsTree = doc.add(dict({ Kids: [leaf] }));
    const catalog = doc.add(
      dict({
        Type: name('Catalog'),
        Pages: pagesRef,
        Outlines: rootRef,
        Dests: dict({ Ch2: [pageRefs[1]!, name('XYZ'), PDF_NULL, 700, PDF_NULL] }),
        Names: dict({ Dests: destsTree }),
      }),
    );
    return doc.build(catalog);
  }

  it('reads every way an item jumps, its depth and whether it stands open', () => {
    const items = readPdfOutline(PdfFile.parse(sourcePdf()));
    expect(items).toEqual([
      { title: 'Chapter One', level: 1, open: true, pageIndex: 0, topPt: 700 },
      { title: 'Section One A', level: 2, open: false, pageIndex: 0, topPt: 600 },
      { title: 'Chapter Two', level: 1, open: false, pageIndex: 1, topPt: 700 },
      { title: 'Section Two A', level: 2, open: false, pageIndex: 1, topPt: 500 },
      // An item that jumps nowhere takes its first descendant's place.
      { title: 'Grouping only', level: 1, open: true, pageIndex: 2 },
      { title: 'Приложение', level: 2, open: false, pageIndex: 2 },
    ]);
  });

  it('writes the same outline into the PDF made from it — order, nesting, open state', async () => {
    const out = await Ream.parse(sourcePdf()).convert('pdf', { fonts: FONTS });
    const { items } = readOutline(out);
    expect(items.map((i) => [i.depth, i.title, i.page])).toEqual([
      [0, 'Chapter One', 0],
      [1, 'Section One A', 0],
      [0, 'Chapter Two', 1],
      [1, 'Section Two A', 1],
      [0, 'Grouping only', 2],
      [1, 'Приложение', 2],
    ]);
    // Table 153 — the closed chapter counts its one child negative.
    expect(items.map((i) => i.count)).toEqual([1, undefined, -1, undefined, 1, undefined]);
  });

  it('keeps the outline a document of ours had through a PDF round trip', async () => {
    const xml =
      para('Heading1', 'Қазақ тілі') +
      body() +
      para('Heading2', 'Second level') +
      para('Heading1', 'Third', '<w:pageBreakBefore/>') +
      body();
    const first = await convert(xml);
    const again = await Ream.parse(first).convert('pdf', { fonts: FONTS });
    const shape = (b: Uint8Array) => readOutline(b).items.map((i) => [i.depth, i.title, i.page]);
    expect(shape(again)).toEqual(shape(first));
  });
});

describe("a source PDF's bookmarks that open a URI, carried in their place", () => {
  /**
   * Two pages of Helvetica text and an outline whose URI actions (ISO 32000-1
   * §12.6.4.7) stand among go-to items, the way freeculture.pdf's do: one with
   * a go-to child of its own, one under a chapter whose /URI is raw UTF-8 (an
   * IRI, not the ASCII Table 206 asks for), one relative to the catalog's
   * /URI /Base (Table 207), a javascript: one, and two at the very end, after
   * the last item that jumps to a page.
   */
  function sourcePdf(): Uint8Array {
    const doc = new PdfDocument();
    const font = doc.add(
      dict({ Type: name('Font'), Subtype: name('Type1'), BaseFont: name('Helvetica') }),
    );
    const pagesMap = dict({ Type: name('Pages'), Kids: [], Count: 2 });
    const pagesRef = doc.add(pagesMap);
    const pageRefs = [
      ['Chapter One', 'Linked section'],
      ['Chapter Two', 'Body text'],
    ].map(([first, second]) => {
      const ops = `BT /F1 12 Tf 1 0 0 1 72 700 Tm (${first!}) Tj 1 0 0 1 72 600 Tm (${second!}) Tj ET`;
      const page = doc.add(
        dict({
          Type: name('Page'),
          Parent: pagesRef,
          MediaBox: [0, 0, 612, 792],
          Resources: dict({ Font: dict({ F1: font }) }),
          Contents: doc.add(stream({}, new TextEncoder().encode(ops))),
        }),
      );
      (pagesMap.get('Kids') as Array<PdfValue>).push(page);
      return page;
    });
    const root = dict({ Type: name('Outlines') });
    const rootRef = doc.add(root);
    type Node = { d: PdfDict; r: PdfRef };
    const item = (title: string, target: PdfValue | { uri: PdfValue }): Node => {
      const d = dict({ Title: title, Parent: rootRef });
      if (Array.isArray(target)) d.set('Dest', target);
      else d.set('A', dict({ S: name('URI'), URI: (target as { uri: PdfValue }).uri }));
      return { d, r: doc.add(d) };
    };
    const link = (parent: Node, kids: Array<Node>): void => {
      parent.d.set('First', kids[0]!.r);
      parent.d.set('Last', kids[kids.length - 1]!.r);
      kids.forEach((k, i) => {
        k.d.set('Parent', parent.r);
        if (i > 0) k.d.set('Prev', kids[i - 1]!.r);
        if (i < kids.length - 1) k.d.set('Next', kids[i + 1]!.r);
      });
    };
    const ch1 = item('Chapter One', [pageRefs[0]!, name('XYZ'), 0, 700, 0]);
    const site = item('Website', { uri: 'http://example.com/' });
    const under = item('Linked section', [pageRefs[0]!, name('XYZ'), 0, 600, 0]);
    const ch2 = item('Chapter Two', [pageRefs[1]!, name('XYZ'), 0, 700, 0]);
    const iri = item('Kazakh wiki', {
      uri: new PdfHexString(new TextEncoder().encode('https://пример.рф/вики?q=тілі#ә')),
    });
    const relative = item('Notes', { uri: 'notes/a b.html' });
    const script = item('Script', { uri: 'javascript:app.alert(1)' });
    const donate = item('Donate', { uri: 'https://example.com/donate' });
    link({ d: root, r: rootRef }, [ch1, site, ch2, script, donate]);
    link(site, [under]);
    link(ch2, [iri, relative]);
    site.d.set('Count', 1);
    ch2.d.set('Count', -2); // closed
    root.set('Count', 6);
    const catalog = doc.add(
      dict({
        Type: name('Catalog'),
        Pages: pagesRef,
        Outlines: rootRef,
        URI: dict({ Base: 'https://example.org/book/' }),
      }),
    );
    return doc.build(catalog);
  }

  it('reads a URI action as the URI it opens, in its place, with no page', () => {
    expect(readPdfOutline(PdfFile.parse(sourcePdf()))).toEqual([
      { title: 'Chapter One', level: 1, open: false, pageIndex: 0, topPt: 700 },
      { title: 'Website', level: 1, open: true, uri: 'http://example.com/' },
      { title: 'Linked section', level: 2, open: false, pageIndex: 0, topPt: 600 },
      { title: 'Chapter Two', level: 1, open: false, pageIndex: 1, topPt: 700 },
      // Raw UTF-8 reads as the characters it spells, as pdf.js reads it.
      { title: 'Kazakh wiki', level: 2, open: false, uri: 'https://пример.рф/вики?q=тілі#ә' },
      // Table 207: relative to the catalog's /URI /Base.
      { title: 'Notes', level: 2, open: false, uri: 'https://example.org/book/notes/a%20b.html' },
      { title: 'Script', level: 1, open: false, uri: 'javascript:app.alert(1)' },
      { title: 'Donate', level: 1, open: false, uri: 'https://example.com/donate' },
    ]);
  });

  it('writes every item in its place and level, a URI action for each URI', async () => {
    const out = await Ream.parse(sourcePdf()).convert('pdf', { fonts: FONTS });
    const { items } = readOutline(out);
    expect(items.map((i) => [i.depth, i.title, i.uri ?? i.page])).toEqual([
      [0, 'Chapter One', 0],
      [0, 'Website', 'http://example.com/'],
      [1, 'Linked section', 0],
      [0, 'Chapter Two', 1],
      // §12.6.4.7 — 7-bit ASCII: IDNA host, UTF-8 octets percent-encoded.
      [
        1,
        'Kazakh wiki',
        'https://xn--e1afmkfd.xn--p1ai/%D0%B2%D0%B8%D0%BA%D0%B8?q=%D1%82%D1%96%D0%BB%D1%96#%D3%99',
      ],
      [1, 'Notes', 'https://example.org/book/notes/a%20b.html'],
      // The link allowlist refuses the script: the item stays, doing nothing.
      [0, 'Script', -1],
      [0, 'Donate', 'https://example.com/donate'],
    ]);
    // The closed chapter still counts its two children negative.
    expect(items.map((i) => i.count)).toEqual([
      undefined,
      1,
      undefined,
      -2,
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
    for (const i of items) if (i.uri !== undefined) expect(i.uri).toMatch(/^[\x21-\x7e]+$/);
  });

  it('carries the URI items into PDF/A too — 19005 permits the URI action', async () => {
    const out = await Ream.parse(sourcePdf()).convert('pdf', { fonts: FONTS, pdfA: 'PDF/A-2b' });
    expect(
      readOutline(out)
        .items.filter((i) => i.uri !== undefined)
        .map((i) => i.title),
    ).toEqual(['Website', 'Kazakh wiki', 'Notes', 'Donate']);
  });

  it('puts no URI item into a DOCX or an HTML page, as text or as a link', async () => {
    const docx = await Ream.parse(sourcePdf()).convert('docx', { fonts: FONTS });
    const xml = new TextDecoder().decode(unzipSync(docx)['word/document.xml']);
    const html = new TextDecoder().decode(
      await Ream.parse(sourcePdf()).convert('html', { fonts: FONTS }),
    );
    for (const said of [xml, html]) {
      expect(said).toContain('Chapter Two');
      for (const absent of ['Website', 'Donate', 'Kazakh wiki', 'example.com', 'example.org']) {
        expect(said).not.toContain(absent);
      }
    }
  });
});
