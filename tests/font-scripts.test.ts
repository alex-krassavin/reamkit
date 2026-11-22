// E-FONT F4/F6 — the writing systems the curated Latin substitutes cannot draw.
// 2145 characters of the corpus are Han, Kana, Hangul, Arabic or a geometric
// symbol, and every one of them was a notdef box: the five families are Latin.
// A face for the script a document actually holds text in is fetched when it is
// needed, and a run reaches it per CHARACTER.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildPptx } from './fixtures/build-pptx';
import { buildXlsx } from './fixtures/build-xlsx';
import type { FetchLike } from '@/core/fonts';
import type { FlowDoc } from '@/core/ir/flow';
import { Ream } from '@/core/converter/ream';
import { clearFontCache } from '@/core/fonts';
import {
  forEachTextInFlow,
  hanFaceOfLang,
  langOfCharset,
  scriptForCodepoint,
  scriptsInFlow,
} from '@/core/fonts/scripts';
import { buildTrueType } from '@/core/font/ttf-build';
import { PdfFile } from '@/pdf-reader/document';
import { extractPageText } from '@/pdf-reader/text';

const cp = (s: string): number => s.codePointAt(0)!;

describe('the script a character belongs to', () => {
  it('names the face each writing system needs', () => {
    expect(scriptForCodepoint(cp('ا'))).toBe('arabic');
    expect(scriptForCodepoint(cp('א'))).toBe('hebrew');
    expect(scriptForCodepoint(cp('ก'))).toBe('thai');
    expect(scriptForCodepoint(cp('か'))).toBe('jp');
    expect(scriptForCodepoint(cp('한'))).toBe('kr');
    expect(scriptForCodepoint(cp('☐'))).toBe('symbols2');
    expect(scriptForCodepoint(cp('①'))).toBe('symbols1');
    expect(scriptForCodepoint(cp('क'))).toBe('devanagari');
    expect(scriptForCodepoint(cp('ক'))).toBe('bengali');
    expect(scriptForCodepoint(cp('ਕ'))).toBe('gurmukhi');
    expect(scriptForCodepoint(cp('ક'))).toBe('gujarati');
    expect(scriptForCodepoint(cp('କ'))).toBe('oriya');
    expect(scriptForCodepoint(cp('க'))).toBe('tamil');
    expect(scriptForCodepoint(cp('క'))).toBe('telugu');
    expect(scriptForCodepoint(cp('ಕ'))).toBe('kannada');
    expect(scriptForCodepoint(cp('ക'))).toBe('malayalam');
    expect(scriptForCodepoint(0xa8f2)).toBe('devanagari'); // Devanagari extended
    expect(scriptForCodepoint(0x1cd0)).toBe('devanagari'); // Vedic extensions
    // The dandas end a sentence in several of them, and every face has them.
    expect(scriptForCodepoint(cp('।'))).toBeUndefined();
    expect(scriptForCodepoint(cp('॥'))).toBeUndefined();
    // A Latin family draws these itself.
    expect(scriptForCodepoint(cp('A'))).toBeUndefined();
    expect(scriptForCodepoint(cp('Я'))).toBeUndefined();
  });

  it('lets the DOCUMENT decide which face draws unified Han', () => {
    // The character is the same in all three languages and drawn differently in
    // each; only its neighbours say which.
    expect(scriptForCodepoint(cp('行'))).toBe('sc');
    expect(scriptForCodepoint(cp('行'), 'jp')).toBe('jp');
    const kana = Ream.parse(sheet('行くこと')).flow;
    expect(scriptsInFlow(kana).hanFace).toBe('jp');
    const hangul = Ream.parse(sheet('행 行')).flow;
    expect(scriptsInFlow(hangul).hanFace).toBe('kr');
    const plain = Ream.parse(sheet('行政')).flow;
    expect(scriptsInFlow(plain).hanFace).toBe('sc');
    expect([...scriptsInFlow(plain).scripts]).toEqual(['sc']);
    // A Latin document asks for nothing, and so downloads nothing.
    expect([...scriptsInFlow(Ream.parse(sheet('hello')).flow).scripts]).toEqual([]);
  });
});

const sheet = (text: string): Uint8Array => buildXlsx([[text]]);

describe('a face for a script the substitutes lack', () => {
  // Every download answers with the one face the repo carries: the test is
  // about WHICH faces are asked for, and a real Noto is 200 KB at best.
  const serveRoboto = (asked: Array<string>): FetchLike => {
    const bytes = new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf'));
    return (url) => {
      asked.push(url.split('/').pop() ?? url);
      return Promise.resolve({
        ok: true,
        arrayBuffer: () =>
          Promise.resolve(
            bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
          ),
      });
    };
  };

  it('is fetched when the document holds that script', async () => {
    clearFontCache();
    const asked: Array<string> = [];
    await Ream.parse(sheet('hello مرحبا')).convert('pdf', { fontFetch: serveRoboto(asked) });
    expect(asked.some((f) => f.startsWith('NotoSansArabic'))).toBe(true);
    // …in the regular weight only: ten megabytes of Noto Sans SC four times
    // over is not a substitute, it is a download.
    expect(asked.filter((f) => f.startsWith('NotoSansArabic'))).toHaveLength(1);
  });

  it('asks for BOTH symbol faces when a document holds one symbol', async () => {
    // The repertory did not fit in one file and Noto's two barely overlap: the
    // ballot box is in the second face, the circled one in the first. Given
    // only the second, ZapfDingbats.pdf drew its stars and boxed its crosses.
    clearFontCache();
    const asked: Array<string> = [];
    await Ream.parse(sheet('☐')).convert('pdf', { fontFetch: serveRoboto(asked) });
    expect(asked.some((f) => f.startsWith('NotoSansSymbols_'))).toBe(true);
    expect(asked.some((f) => f.startsWith('NotoSansSymbols2'))).toBe(true);
  });

  it('is fetched for each script of India a document holds', async () => {
    clearFontCache();
    const asked: Array<string> = [];
    await Ream.parse(sheet('हिन्दी தமிழ் বাংলা')).convert('pdf', { fontFetch: serveRoboto(asked) });
    const noto = asked.filter((f) => f.startsWith('Noto')).sort();
    expect(noto).toEqual([
      'NotoSansBengali_400Regular.ttf',
      'NotoSansDevanagari_400Regular.ttf',
      'NotoSansTamil_400Regular.ttf',
    ]);
  });

  it('is not fetched for Devanagari when a Bengali sentence ends in a danda', async () => {
    clearFontCache();
    const asked: Array<string> = [];
    await Ream.parse(sheet('আমি বাংলায় লিখি।')).convert('pdf', { fontFetch: serveRoboto(asked) });
    expect(asked.filter((f) => f.startsWith('Noto'))).toEqual(['NotoSansBengali_400Regular.ttf']);
  });

  it('draws the letters, and they are in the text layer', async () => {
    // A face with the six letters of हिन्दी, standing in for Noto Sans
    // Devanagari; everything else is served Roboto, which has none of them.
    const BOX = [
      { op: 'move', x: 50, y: 0 },
      { op: 'line', x: 450, y: 0 },
      { op: 'line', x: 450, y: 700 },
      { op: 'line', x: 50, y: 700 },
      { op: 'close' },
    ] as const;
    const devanagari = buildTrueType({
      family: 'Test Devanagari',
      bold: false,
      italic: false,
      postScriptName: 'TestDevanagari',
      glyphs: [
        { codePoints: [0x20], outline: [], advance: 250 },
        ...[0x0939, 0x093f, 0x0928, 0x094d, 0x0926, 0x0940].map((c) => ({
          codePoints: [c],
          outline: BOX,
          advance: 500,
        })),
      ],
      ascent: 800,
      descent: -200,
      fsType: 0,
    });
    const roboto = serveRoboto([]);
    const served: FetchLike = (url) =>
      url.includes('NotoSansDevanagari')
        ? Promise.resolve({
            ok: true,
            arrayBuffer: () => Promise.resolve(devanagari.slice().buffer),
          })
        : roboto(url);
    clearFontCache();
    const { bytes, losses } = await Ream.parse(sheet('Hindi: हिन्दी')).convertWithReport('pdf', {
      fontFetch: served,
    });
    clearFontCache();
    expect(losses.filter((l) => l.feature === 'text')).toEqual([]);
    expect(new TextDecoder('latin1').decode(bytes)).toContain('+TestDevanagari');
    const text = new TextDecoder().decode(await Ream.parse(bytes).convert('md'));
    expect(text).toContain('हिन्दी');
  });

  it('…and not when it does not', async () => {
    clearFontCache();
    const asked: Array<string> = [];
    await Ream.parse(sheet('hello')).convert('pdf', { fontFetch: serveRoboto(asked) });
    expect(asked.some((f) => f.startsWith('Noto'))).toBe(false);
  });
});

// One text box holding one run, its `a:rPr` attributes as given.
const textBox = (text: string, rPr: string): string =>
  `<p:sp><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="7315200" cy="1828800"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
  `<p:txBody><a:bodyPr/><a:p><a:r><a:rPr${rPr}/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;

/** The East Asian language of each run of a document, by its text. */
function langsOf(flow: FlowDoc): Map<string, string | undefined> {
  const out = new Map<string, string | undefined>();
  forEachTextInFlow(flow, (text, run) => out.set(text, run.properties.eastAsianLang));
  return out;
}

const scriptsOf = (bytes: Uint8Array): Array<string> =>
  [...scriptsInFlow(Ream.parse(bytes).flow).scripts].sort();

describe('the face unified Han is drawn in, by the language of its run', () => {
  // A deck with a Chinese slide and a Japanese one: one kana anywhere sent
  // every ideograph to Noto Sans JP, and that face has no 这 应 该 at all —
  // the Chinese slide came out blank where its simplified characters stood.
  const MIXED = buildPptx([
    textBox('这是中文字', ' lang="zh-CN" altLang="en-US"'),
    textBox('あれは漢字', ' lang="ja-JP" altLang="en-US"'),
  ]);

  it('reads the face off a language tag', () => {
    expect(hanFaceOfLang('zh-CN')).toBe('sc');
    expect(hanFaceOfLang('zh-SG')).toBe('sc');
    // No Traditional face is fetched: SC draws 這 個 們, JP has no 这.
    expect(hanFaceOfLang('zh-TW')).toBe('sc');
    expect(hanFaceOfLang('ja-JP')).toBe('jp');
    expect(hanFaceOfLang('ko_KR')).toBe('kr');
    expect(hanFaceOfLang('KO')).toBe('kr');
    expect(hanFaceOfLang('en-US')).toBeUndefined();
    expect(hanFaceOfLang(undefined)).toBeUndefined();
  });

  it('fetches a face for each language a run of Han is in', () => {
    const { scripts, hanFace } = scriptsInFlow(Ream.parse(MIXED).flow);
    expect([...scripts].sort()).toEqual(['jp', 'sc']);
    // The face for a run that says nothing: Kana is in the document.
    expect(hanFace).toBe('jp');
  });

  it('leaves a run whose language is not East Asian to the document', () => {
    // A run IN Japanese counts as Kana does, so the Latin-tagged one goes to
    // JP with it and nothing more is fetched.
    const japanese = buildPptx([
      textBox('漢字', ' lang="ja-JP"'),
      textBox('中文', ' lang="en-US"'),
    ]);
    expect(scriptsOf(japanese)).toEqual(['jp']);
    expect(scriptsInFlow(Ream.parse(japanese).flow).hanFace).toBe('jp');
    // …and with nothing Japanese or Korean anywhere, Han is Chinese, as before.
    expect(scriptsOf(buildPptx([textBox('中文', ' lang="en-US"')]))).toEqual(['sc']);
    expect(scriptsOf(buildPptx([textBox('中文', '')]))).toEqual(['sc']);
  });

  it('takes the East Asian one of a slide run’s two languages', () => {
    const deck = buildPptx([
      textBox('甲', ' lang="en-US" altLang="ja-JP"'),
      textBox('乙', ' lang="zh-CN" altLang="en-US"'),
      textBox('丙', ' lang="en-US"'),
    ]);
    const langs = langsOf(Ream.parse(deck).flow);
    expect(langs.get('甲')).toBe('ja-JP');
    expect(langs.get('乙')).toBe('zh-CN');
    expect(langs.get('丙')).toBe('en-US');
  });

  it('reads w:lang @w:eastAsia, inherited from the document defaults or the run’s own', () => {
    const docx = buildDocxFromBody(
      '<w:p><w:r><w:t>中文</w:t></w:r></w:p>' +
        '<w:p><w:r><w:rPr><w:lang w:eastAsia="ja-JP"/></w:rPr><w:t>漢字</w:t></w:r></w:p>',
      {
        stylesXml:
          '<w:docDefaults><w:rPrDefault><w:rPr>' +
          '<w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>' +
          '</w:rPr></w:rPrDefault></w:docDefaults>',
      },
    );
    const flow = Ream.parse(docx).flow;
    const langs = langsOf(flow);
    expect(langs.get('中文')).toBe('zh-CN');
    expect(langs.get('漢字')).toBe('ja-JP');
    expect([...scriptsInFlow(flow).scripts].sort()).toEqual(['jp', 'sc']);
  });

  it('reads a spreadsheet font’s character set as the language of its cells', () => {
    const font = (charset: number): string =>
      `<font><sz val="11"/><name val="Calibri"/><charset val="${charset}"/></font>`;
    const xlsx = buildXlsx({
      rows: [[{ value: '中文', styleIndex: 0 }], [{ value: '漢字', styleIndex: 1 }]],
      stylesXml:
        `<fonts count="2">${font(134)}${font(128)}</fonts>` +
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
        '<borders count="1"><border/></borders>' +
        '<cellStyleXfs count="1"><xf/></cellStyleXfs>' +
        '<cellXfs count="2"><xf fontId="0"/><xf fontId="1"/></cellXfs>',
    });
    const flow = Ream.parse(xlsx).flow;
    const langs = langsOf(flow);
    expect(langs.get('中文')).toBe('zh-CN');
    expect(langs.get('漢字')).toBe('ja-JP');
    expect([...scriptsInFlow(flow).scripts].sort()).toEqual(['jp', 'sc']);
    expect([128, 129, 130, 134, 136, 0, 204].map(langOfCharset)).toEqual([
      'ja-JP',
      'ko-KR',
      'ko-KR',
      'zh-CN',
      'zh-TW',
      undefined,
      undefined,
    ]);
  });

  it('draws each run in its own language’s face', async () => {
    // Noto Sans SC and JP cut down to the ten characters below; JP's lacks 这,
    // as the face itself does. Everything else is served Roboto.
    const font = (name: string): Uint8Array =>
      new Uint8Array(readFileSync(`tests/fixtures/fonts/${name}.ttf`));
    const faces: ReadonlyArray<[string, Uint8Array]> = [
      ['NotoSansSC', font('NotoSansSC-subset')],
      ['NotoSansJP', font('NotoSansJP-subset')],
    ];
    const roboto = font('Roboto-Regular');
    const asked: Array<string> = [];
    const served: FetchLike = (url) => {
      asked.push(url.split('/').pop() ?? url);
      const bytes = faces.find(([name]) => url.includes(name))?.[1] ?? roboto;
      return Promise.resolve({
        ok: true,
        arrayBuffer: () => Promise.resolve(bytes.slice().buffer),
      });
    };
    clearFontCache();
    const { bytes, losses } = await Ream.parse(MIXED).convertWithReport('pdf', {
      fontFetch: served,
    });
    clearFontCache();
    expect(asked.filter((f) => f.startsWith('NotoSans')).sort()).toEqual([
      'NotoSansJP_400Regular.ttf',
      'NotoSansSC_400Regular.ttf',
    ]);
    expect(losses.filter((l) => l.feature === 'text')).toEqual([]);
    const file = PdfFile.parse(bytes);
    // Each Kana or Han character of a slide, and the face (its /BaseFont, as
    // the reader gives it) it is drawn in.
    const drawn = file
      .pages()
      .map((page) =>
        extractPageText(file, page).flatMap((run) =>
          [...run.text]
            .filter((ch) => /[぀-ヿ一-鿿]/u.test(ch))
            .map((ch) => [ch, run.fontName ?? '?'] as const),
        ),
      );
    // The Chinese slide whole in SC — 是 中 文 字 too, which JP draws as well.
    expect(drawn[0]).toEqual([...'这是中文字'].map((ch) => [ch, 'notosanssc-regular']));
    // The Japanese slide whole in JP, its 漢 and 字 among it.
    expect(drawn[1]).toEqual([...'あれは漢字'].map((ch) => [ch, 'notosansjp-regular']));
  });
});
