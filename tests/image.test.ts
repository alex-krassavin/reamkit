import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { unzlibSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildIndexedPng, buildTinyPng } from './fixtures/build-png';
import { countShown, showPattern } from './fixtures/pdf-show';
import { defaultColorResolver } from '@/core/drawingml/colors';
import { ResourceStore, eighthPtToPt, emuToPt, halfPtToPt, twipsToPt } from '@/core/ir';
import { convertDocxToPdfSync } from '@/core/converter';
import { parseTtf } from '@/core/font';
import { OpcPackage } from '@/core/opc';
import { parseDocument } from '@/word';
import { readDocx } from '@/word/docx-reader';
import { detectImageFormat, embedImage, prepareImage } from '@/pdf';
import { PdfDocument } from '@/pdf/writer';

const here = dirname(fileURLToPath(import.meta.url));
const FONTS = {
  regular: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf'))),
};

const latin1 = new TextDecoder('latin1');
const asLatin1 = (b: Uint8Array): string => latin1.decode(b);

function drawingXml(rId: string, cxEmu: number, cyEmu: number, srcRect = ''): string {
  return `<w:r><w:drawing>
    <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
      <wp:extent cx="${cxEmu}" cy="${cyEmu}"/>
      <wp:docPr id="1" name="Picture 1"/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:blipFill>
              <a:blip r:embed="${rId}"/>
              ${srcRect}
            </pic:blipFill>
            <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cxEmu}" cy="${cyEmu}"/></a:xfrm></pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing></w:r>`;
}

describe('detectImageFormat', () => {
  it('recognises PNG magic bytes', () => {
    expect(
      detectImageFormat(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])),
    ).toBe('png');
  });

  it('recognises JPEG magic bytes', () => {
    expect(detectImageFormat(new Uint8Array([0xff, 0xd8, 0xff, 0xe0]))).toBe('jpeg');
  });

  it('returns null for non-image bytes', () => {
    expect(detectImageFormat(new Uint8Array([0x00, 0x01, 0x02, 0x03]))).toBeNull();
  });

  it('recognises a JPEG 2000 JP2 box signature and a raw codestream', () => {
    const jp2Sig = new Uint8Array([
      0x00, 0x00, 0x00, 0x0c, 0x6a, 0x50, 0x20, 0x20, 0x0d, 0x0a, 0x87, 0x0a,
    ]);
    expect(detectImageFormat(jp2Sig)).toBe('jpeg2000');
    expect(detectImageFormat(new Uint8Array([0xff, 0x4f, 0xff, 0x51]))).toBe('jpeg2000');
  });
});

// Minimal JP2: signature box + jp2h{ ihdr(height=64, width=128) }. No codestream
// (we pass the bytes through to /JPXDecode and only read the dimensions).
const JP2_64x128 = new Uint8Array([
  0x00,
  0x00,
  0x00,
  0x0c,
  0x6a,
  0x50,
  0x20,
  0x20,
  0x0d,
  0x0a,
  0x87,
  0x0a, // signature
  0x00,
  0x00,
  0x00,
  0x1e,
  0x6a,
  0x70,
  0x32,
  0x68, // jp2h box (len 30)
  0x00,
  0x00,
  0x00,
  0x16,
  0x69,
  0x68,
  0x64,
  0x72, // ihdr box (len 22)
  0x00,
  0x00,
  0x00,
  0x40, // HEIGHT = 64
  0x00,
  0x00,
  0x00,
  0x80, // WIDTH = 128
  0x00,
  0x01,
  0x07,
  0x07,
  0x00,
  0x00, // NC=1, BPC, C, UnkC, IPR
]);

describe('JPEG 2000 embedding (/JPXDecode pass-through)', () => {
  it('reads ihdr dimensions and emits a JPXDecode image XObject', () => {
    const doc = new PdfDocument();
    const img = embedImage(doc, JP2_64x128);
    expect(img.widthPx).toBe(128);
    expect(img.heightPx).toBe(64);
    const pdf = new TextDecoder('latin1').decode(doc.build(img.ref));
    expect(pdf).toContain('/Filter /JPXDecode');
    expect(pdf).toContain('/Width 128');
    expect(pdf).toContain('/Height 64');
  });
});

describe('Drawing parser', () => {
  it('produces a BodyElement of kind=image when a w:drawing is present', () => {
    const body = `<w:p>${drawingXml('rId20', 914400, 685800)}</w:p>`;
    const docx = buildDocxFromBody(body);
    const pkg = OpcPackage.open(docx);
    // The parser resolves drawing relationship ids through the supplied
    // ImageResolver into content-addressed ResourceIds.
    const store = new ResourceStore();
    const expectedId = store.put(new Uint8Array([1, 2, 3]));
    const parsed = parseDocument(pkg.getMainDocument().data, {
      resolveColor: defaultColorResolver,
      resolveImage: (relId) => (relId === 'rId20' ? expectedId : undefined),
    });
    expect(parsed).toHaveLength(1);
    expect(parsed[0]!.kind).toBe('image');
    if (parsed[0]!.kind !== 'image') throw new Error('unreachable');
    expect(parsed[0]!.image.resource).toBe(expectedId);
    expect(parsed[0]!.image.width).toBe(emuToPt(914400));
    expect(parsed[0]!.image.height).toBe(emuToPt(685800));
  });

  it('produces an image from a legacy VML picture (w:pict / v:imagedata)', () => {
    // §14 VML: <v:imagedata r:id> binds the media; the @style box gives the
    // size in CSS units (pt here). The reader recovers it the same as a blip.
    const body =
      '<w:p><w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml">' +
      '<v:shape style="width:72pt;height:54pt" alt="a legacy picture">' +
      '<v:imagedata r:id="rId20"/></v:shape></w:pict></w:r></w:p>';
    const pkg = OpcPackage.open(buildDocxFromBody(body));
    const store = new ResourceStore();
    const expectedId = store.put(new Uint8Array([4, 5, 6]));
    const parsed = parseDocument(pkg.getMainDocument().data, {
      resolveColor: defaultColorResolver,
      resolveImage: (relId) => (relId === 'rId20' ? expectedId : undefined),
    });
    expect(parsed).toHaveLength(1);
    if (parsed[0]!.kind !== 'image') throw new Error('expected an image block');
    expect(parsed[0]!.image.resource).toBe(expectedId);
    expect(parsed[0]!.image.width).toBe(72); // 72pt verbatim
    expect(parsed[0]!.image.height).toBe(54);
    expect(parsed[0]!.image.altText).toBe('a legacy picture');
  });

  it('collapses a lone image to a block even when wrapped in tracked-change w:ins', () => {
    // The writer flattens tracked changes, so a re-read sees a bare <w:drawing>
    // and collapses the paragraph to an image block. To keep the round-trip
    // symmetric the FIRST read must collapse the same way even though the
    // drawing hides inside <w:ins><w:r>…</w:r></w:ins>.
    const body = `<w:p><w:ins w:id="1" w:author="A" w:date="2020-01-01T00:00:00Z">${drawingXml('rId20', 914400, 685800)}</w:ins></w:p>`;
    const pkg = OpcPackage.open(buildDocxFromBody(body));
    const store = new ResourceStore();
    const expectedId = store.put(new Uint8Array([7, 8, 9]));
    const parsed = parseDocument(pkg.getMainDocument().data, {
      resolveColor: defaultColorResolver,
      resolveImage: (relId) => (relId === 'rId20' ? expectedId : undefined),
    });
    expect(parsed).toHaveLength(1);
    if (parsed[0]!.kind !== 'image') throw new Error('expected an image block');
    expect(parsed[0]!.image.resource).toBe(expectedId);
  });

  it('does NOT collapse an image wrapped in a hyperlink (the link must survive)', () => {
    // A hyperlinked image stays inline so the writer can keep its href; the
    // paragraph must remain a paragraph, not collapse to an image block.
    const body =
      '<w:p><w:hyperlink r:id="rIdLink">' +
      `${drawingXml('rId20', 914400, 685800)}</w:hyperlink></w:p>`;
    const pkg = OpcPackage.open(buildDocxFromBody(body));
    const store = new ResourceStore();
    const expectedId = store.put(new Uint8Array([1, 1, 1]));
    const parsed = parseDocument(pkg.getMainDocument().data, {
      resolveColor: defaultColorResolver,
      resolveImage: (relId) => (relId === 'rId20' ? expectedId : undefined),
      resolveHyperlink: (relId) => (relId === 'rIdLink' ? 'https://example.com' : undefined),
    });
    expect(parsed).toHaveLength(1);
    expect(parsed[0]!.kind).toBe('paragraph');
  });

  it('skips a dangling VML picture whose media part is absent', () => {
    // A <v:imagedata r:id> pointing at media stripped from the package (some
    // corpus files do this) resolves to nothing — the reader must NOT emit a
    // phantom image; the paragraph stays plain so the round-trip is symmetric.
    const body =
      '<w:p><w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml">' +
      '<v:shape style="width:72pt;height:54pt"><v:imagedata r:id="rIdMissing"/>' +
      '</v:shape></w:pict></w:r></w:p>';
    const pkg = OpcPackage.open(buildDocxFromBody(body));
    const parsed = parseDocument(pkg.getMainDocument().data, {
      resolveColor: defaultColorResolver,
      resolveImage: () => undefined, // nothing resolves
    });
    expect(parsed.some((el) => el.kind === 'image')).toBe(false);
  });
});

describe('Image rendering end-to-end', () => {
  it('emits an XObject Image and Do operator for an embedded PNG', () => {
    const png = buildTinyPng(2, 2, [255, 0, 0, 255]); // 2×2 red
    const body = `
      <w:p><w:r><w:t>Before image</w:t></w:r></w:p>
      <w:p>${drawingXml('rId20', 914400, 914400)}</w:p>
      <w:p><w:r><w:t>After image</w:t></w:r></w:p>`;
    const docx = buildDocxFromBody(body, {
      images: { rId20: { contentType: 'image/png', bytes: png, extension: 'png' } },
    });
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    expect(text).toContain('/Type /XObject');
    expect(text).toContain('/Subtype /Image');
    expect(text).toContain('/Filter /FlateDecode');
    expect(text).toMatch(/\/Im\d+ Do/);
    expect(text).toMatch(/\d+(\.\d+)? 0 0 \d+(\.\d+)? \d+(\.\d+)? \d+(\.\d+)? cm/);
  });

  it('preserves text and renders image when a paragraph mixes both', () => {
    const png = buildTinyPng(4, 4, [0, 200, 0, 255]); // small green square
    // Paragraph with text + inline image + more text in the same w:p.
    const body = `
      <w:p>
        <w:r><w:t xml:space="preserve">Before </w:t></w:r>
        ${drawingXml('rId20', 304800, 304800)}
        <w:r><w:t xml:space="preserve"> After</w:t></w:r>
      </w:p>`;
    const docx = buildDocxFromBody(body, {
      images: { rId20: { contentType: 'image/png', bytes: png, extension: 'png' } },
    });
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    const parsed = parseTtf(FONTS.regular);

    // Both text segments must be present (not lost to image-block collapse).
    expect(text).toMatch(showPattern(parsed, 'Before'));
    expect(text).toMatch(showPattern(parsed, 'After'));
    // Image XObject is also drawn.
    expect(text).toMatch(/\/Im\d+ Do/);
    // ET / BT sequence indicates we left text mode for the inline image.
    expect(text).toMatch(/ET\nq\n[^\n]+cm\n\/Im\d+ Do\nQ\nBT/);
  });

  it('image XObject lists the right Width and Height', () => {
    const png = buildTinyPng(3, 5, [0, 128, 255, 255]);
    const body = `<w:p>${drawingXml('rId20', 914400, 1828800)}</w:p>`;
    const docx = buildDocxFromBody(body, {
      images: { rId20: { contentType: 'image/png', bytes: png, extension: 'png' } },
    });
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    expect(text).toMatch(/\/Width 3/);
    expect(text).toMatch(/\/Height 5/);
  });
});

describe('a cropped picture (§20.1.8.55 a:srcRect)', () => {
  const png = buildTinyPng(2, 2, [255, 0, 0, 255]);
  // Quarter off the left, half off the bottom: the frame shows the top-right
  // three-eighths of the picture, so the picture is drawn 4/3 as wide and twice
  // as tall as the frame and the frame clips it.
  const cropped = (srcRect: string): string =>
    asLatin1(
      convertDocxToPdfSync(
        buildDocxFromBody(`<w:p>${drawingXml('rId20', 914400, 914400, srcRect)}</w:p>`, {
          images: { rId20: { contentType: 'image/png', bytes: png, extension: 'png' } },
        }),
        { fonts: FONTS },
      ),
    );

  it('scales the picture up and clips the frame to it', () => {
    const text = cropped('<a:srcRect l="25000" b="50000"/>');
    // A clipping path the size of the frame (72pt), then the picture at 96×144.
    expect(text).toMatch(/[\d.]+ [\d.]+ 72 72 re\nW\nn\n96 0 0 144 /u);
    // Moved left by the quarter cut away and down by the half: 24pt and 72pt.
    const cm = /96 0 0 144 ([\d.]+) ([\d.]+) cm/u.exec(text)!;
    const re = /([\d.]+) ([\d.]+) 72 72 re/u.exec(text)!;
    expect(Number(re[1]) - Number(cm[1])).toBeCloseTo(24, 3);
    expect(Number(re[2]) - Number(cm[2])).toBeCloseTo(72, 3);
  });

  it('leaves a picture with no crop alone', () => {
    // Every edge zero is no crop at all — no clip, and the unit square scaled
    // straight to the frame.
    expect(cropped('<a:srcRect l="0" t="0"/>')).toMatch(/\n72 0 0 72 /u);
    expect(cropped('')).toMatch(/\n72 0 0 72 /u);
  });

  it('ignores a crop that would leave nothing of the picture', () => {
    expect(cropped('<a:srcRect l="60000" r="60000"/>')).toMatch(/\n72 0 0 72 /u);
  });
});

describe('a run of inline pictures', () => {
  it('may break a line between two of them', () => {
    // Each inline picture is a character of its own; a run of them with no
    // space between is not one long word. Without the opportunity,
    // VariousPictures.docx's five pictures were one unbreakable box 739pt wide
    // on a 468pt line and the ones we could draw ran off the paper.
    const png = buildTinyPng(2, 2, [255, 0, 0, 255]);
    // Four pictures of 144pt on a 451pt line: two per line, never five across.
    const runs = Array.from({ length: 4 }, () => drawingXml('rId20', 1828800, 914400)).join('');
    const docx = buildDocxFromBody(`<w:p>${runs}</w:p>`, {
      images: { rId20: { contentType: 'image/png', bytes: png, extension: 'png' } },
    });
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    const xs = [...text.matchAll(/\n([\d.]+) 0 0 [\d.]+ ([\d.]+) [\d.]+ cm\n\/Im\d+ Do/gu)].map(
      (m) => Number(m[2]),
    );
    expect(xs).toHaveLength(4);
    // Every picture starts within the page, not past its right edge.
    expect(Math.max(...xs)).toBeLessThan(400);
  });
});

describe('an anchored drawing beside text', () => {
  // §20.4.2.3 — an anchored drawing is not in the line: it hangs off the
  // paragraph at a position of its own and the text flows past it. Read as an
  // inline picture it split the line it sat in: anchor-position.docx put its
  // picture between the "A" and the "B" where every other reader sets "AB"
  // beside it.
  const anchored = (wrap: string): string => {
    const png = buildTinyPng(2, 2, [255, 0, 0, 255]);
    const drawing = `<w:drawing>
      <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                 distT="0" distB="0" distL="0" distR="0" behindDoc="0" locked="0"
                 layoutInCell="1" allowOverlap="1" simplePos="0" relativeHeight="1">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
        <wp:extent cx="650240" cy="650240"/>
        ${wrap}
        <wp:docPr id="1" name="Picture 1"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:blipFill><a:blip r:embed="rId20"/></pic:blipFill>
              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="650240" cy="650240"/></a:xfrm></pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:anchor>
    </w:drawing>`;
    const body = `<w:p><w:r><w:t>A</w:t></w:r><w:r>${drawing}</w:r><w:r><w:t>B</w:t></w:r></w:p>`;
    return asLatin1(
      convertDocxToPdfSync(
        buildDocxFromBody(body, {
          images: { rId20: { contentType: 'image/png', bytes: png, extension: 'png' } },
        }),
        { fonts: FONTS },
      ),
    );
  };

  it('leaves the line it hangs off unbroken', () => {
    // An inline picture leaves text mode mid-line — `ET … Do … BT`. An
    // anchored one is a block of its own, so the line is never interrupted.
    expect(anchored('<wp:wrapSquare wrapText="bothSides"/>')).not.toMatch(
      /ET\nq\n[^\n]+cm\n\/Im\d+ Do\nQ\nBT/u,
    );
  });

  it('still draws the picture', () => {
    expect(anchored('<wp:wrapNone/>')).toMatch(/\/Im\d+ Do/u);
  });
});

// A picture standing alone in a text box is a paragraph the reader collapsed to
// an image BLOCK. Skipped with the tables and nested shapes, it vanished:
// WPGbodyPr.docx sets one inside its outer circle and we drew the circle and the
// words and nothing between them.
describe('a picture inside a shape’s text box', () => {
  it('is drawn, on a line of its own', () => {
    const png = buildTinyPng(2, 2, [0, 0, 255, 255]);
    const inner =
      '<w:p><w:r><w:t>Caption</w:t></w:r></w:p>' +
      `<w:p>${drawingXml('rId20', 914400, 457200)}</w:p>`;
    const shape = `<w:r><w:drawing>
      <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="2743200" cy="1828800"/><wp:docPr id="1" name="S"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1828800"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>
              <wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx>
              <wps:bodyPr/>
            </wps:wsp>
          </a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    const docx = buildDocxFromBody(`<w:p>${shape}</w:p>`, {
      images: { rId20: { contentType: 'image/png', bytes: png, extension: 'png' } },
    });
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    expect(text).toContain('/Subtype /Image');
    expect(text).toMatch(/\/Im\d+ Do/u);
  });
});

describe('image robustness', () => {
  it('skips an unsupported/corrupt image instead of crashing the document', () => {
    // Garbage bytes labelled as PNG — embedImage throws; the document must still
    // render (the bad image is simply omitted, no dangling XObject reference).
    const body =
      `<w:p><w:r><w:t>Before image.</w:t></w:r></w:p>` +
      `<w:p>${drawingXml('rId99', 914400, 914400)}</w:p>` +
      `<w:p><w:r><w:t>After image.</w:t></w:r></w:p>`;
    const opts = {
      images: {
        rId99: {
          contentType: 'image/png' as const,
          bytes: new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]),
          extension: 'png' as const,
        },
      },
    };
    const pdf = convertDocxToPdfSync(buildDocxFromBody(body, opts), { fonts: FONTS });
    const text = asLatin1(pdf);
    // Conversion succeeded and produced a valid PDF...
    expect(text.startsWith('%PDF')).toBe(true);
    // ...with no image drawn (no `/Im… Do`, and crucially no dangling `/ Do`).
    expect(text).not.toMatch(/\/Im\d* Do/);
    expect(text).not.toMatch(/\/ Do/);
    // The surrounding text still rendered (font subset embedded).
    expect(text).toMatch(/\/BaseFont \/[A-Za-z]/);
  });
});

describe('per-part image resolution (C5)', () => {
  it("resolves a header image through the header's own rels, not the main part's", () => {
    // Same rId in both parts, pointing at DIFFERENT images: blue in the
    // header's rels, red in the main document's. The header must get blue.
    const bluePng = buildTinyPng(2, 2, [0, 0, 255, 255]);
    const redPng = buildTinyPng(2, 2, [255, 0, 0, 255]);
    const body =
      '<w:p><w:r><w:t>body</w:t></w:r></w:p>' +
      '<w:sectPr><w:headerReference w:type="default" r:id="rId10"/></w:sectPr>';
    const docx = buildDocxFromBody(body, {
      headerXml: `<w:p>${drawingXml('rId20', 190500, 190500)}</w:p>`,
      headerImages: { rId20: { contentType: 'image/png', bytes: bluePng, extension: 'png' } },
      images: { rId20: { contentType: 'image/png', bytes: redPng, extension: 'png' } },
    });
    const { doc } = readDocx(docx);
    const header = [...(doc.headersFooters?.values() ?? [])][0];
    expect(header).toBeDefined();
    const img = header!.find((el) => el.kind === 'image');
    expect(img).toBeDefined();
    if (img?.kind !== 'image') throw new Error('unreachable');
    const bytes = doc.resources.get(img.image.resource!);
    expect(bytes && Buffer.from(bytes).equals(Buffer.from(bluePng))).toBe(true);
  });
});

describe('PNG colour types beyond 8-bit RGB', () => {
  // The decoder's own output, uncompressed: prepareImage hands back Flate
  // streams, and the point of these is exactly which bytes came out.
  const decode = (
    png: Uint8Array,
  ): { rgb: Array<number>; alpha: Array<number> | undefined; space?: string } => {
    const prepared = prepareImage(png);
    return {
      rgb: [...unzlibSync(prepared.data)],
      alpha: prepared.smaskData ? [...unzlibSync(prepared.smaskData)] : undefined,
      ...(prepared.colorSpace ? { space: prepared.colorSpace } : {}),
    };
  };
  const BLACK = [0, 0, 0] as const;
  const YELLOW = [255, 255, 0] as const;

  it('expands an indexed image through its palette', () => {
    // tdf99135.docx is exactly this: a 1-bit palette of black over yellow,
    // which we drew as nothing at all.
    const png = buildIndexedPng(2, 2, 1, [BLACK, YELLOW], [0, 1, 1, 0]);
    const { rgb, space } = decode(png);
    expect(space).toBe('DeviceRGB');
    expect(rgb).toEqual([0, 0, 0, 255, 255, 0, 255, 255, 0, 0, 0, 0]);
  });

  it('reads a palette index at every legal bit depth', () => {
    for (const depth of [1, 2, 4, 8] as const) {
      const png = buildIndexedPng(2, 1, depth, [BLACK, YELLOW], [1, 0]);
      expect(decode(png).rgb).toEqual([255, 255, 0, 0, 0, 0]);
    }
  });

  it('takes a palette image\u2019s tRNS as the soft mask', () => {
    const png = buildIndexedPng(2, 1, 8, [BLACK, YELLOW], [0, 1], { alpha: [0, 255] });
    const { alpha } = decode(png);
    expect(alpha).toEqual([0, 255]);
  });

  it('de-interlaces an Adam7 image to the same pixels', () => {
    const indices = [0, 1, 1, 0];
    const plain = decode(buildIndexedPng(2, 2, 8, [BLACK, YELLOW], indices));
    const woven = decode(buildIndexedPng(2, 2, 8, [BLACK, YELLOW], indices, { interlaced: true }));
    expect(woven.rgb).toEqual(plain.rgb);
  });
});

describe('a rotated picture (§20.1.7.6 a:xfrm @rot)', () => {
  const png = buildTinyPng(2, 2, [255, 0, 0, 255]);
  const turned = (rot: string): string =>
    asLatin1(
      convertDocxToPdfSync(
        buildDocxFromBody(
          `<w:p><w:r><w:drawing>
            <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
              <wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/>
              <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                    <pic:blipFill><a:blip r:embed="rId20"/></pic:blipFill>
                    <pic:spPr><a:xfrm ${rot}><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></pic:spPr>
                  </pic:pic>
                </a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
          { images: { rId20: { contentType: 'image/png', bytes: png, extension: 'png' } } },
        ),
        { fonts: FONTS },
      ),
    );

  it('turns the picture about its centre', () => {
    // crop-pixel.docx tilts its cover by 641099/60000 = 10.685°, clockwise;
    // PDF measures the other way, so the matrix carries -10.685°.
    const text = turned('rot="641099"');
    const m = /([\d.-]+) ([\d.-]+) ([\d.-]+) ([\d.-]+) [\d.-]+ [\d.-]+ cm\n72 0 0 72 /u.exec(text);
    expect(m).not.toBeNull();
    const rad = (-10.685 * Math.PI) / 180;
    expect(Number(m![1])).toBeCloseTo(Math.cos(rad), 3);
    expect(Number(m![2])).toBeCloseTo(Math.sin(rad), 3);
  });

  it('leaves an unturned picture with no matrix of its own', () => {
    expect(turned('')).not.toMatch(
      /[\d.-]+ [\d.-]+ [\d.-]+ [\d.-]+ [\d.-]+ [\d.-]+ cm\n72 0 0 72 /u,
    );
  });
});

describe('a GIF picture', () => {
  // A 2×2 GIF89a: red, green / green, red, with the LZW stream written by hand
  // — clear, four literals, end, the last two already four bits wide because
  // the table reached eight entries — so the fixture exercises the real
  // decoder, code-width step and all.
  const gif = (transparent: boolean): Uint8Array =>
    new Uint8Array([
      0x47,
      0x49,
      0x46,
      0x38,
      0x39,
      0x61, // "GIF89a"
      0x02,
      0x00,
      0x02,
      0x00,
      0xf0,
      0x00,
      0x00, // 2×2, global table of 2
      0xff,
      0x00,
      0x00,
      0x00,
      0xff,
      0x00, // red, green
      ...(transparent
        ? [0x21, 0xf9, 0x04, 0x01, 0x00, 0x00, 0x01, 0x00] // GCE: index 1 clear
        : []),
      0x2c,
      0x00,
      0x00,
      0x00,
      0x00,
      0x02,
      0x00,
      0x02,
      0x00,
      0x00, // image descriptor
      0x02,
      0x03,
      0x44,
      0x02,
      0x05,
      0x00, // LZW min 2, one 3-byte block
      0x3b, // trailer
    ]);

  it('decodes to raw pixels, since PDF has no GIF filter', () => {
    // dml-picture-in-textframe.docx (and seventeen more) drew an empty box.
    const prepared = prepareImage(gif(false));
    expect(prepared).toMatchObject({
      format: 'gif',
      mimeType: 'image/gif',
      widthPx: 2,
      heightPx: 2,
      colorSpace: 'DeviceRGB',
      bitsPerComponent: 8,
      filter: 'FlateDecode',
    });
    expect([...unzlibSync(prepared.data)]).toEqual([255, 0, 0, 0, 255, 0, 0, 255, 0, 255, 0, 0]);
    expect(prepared.smaskData).toBeUndefined();
  });

  it('turns a transparent index into a soft mask', () => {
    const prepared = prepareImage(gif(true));
    expect([...unzlibSync(prepared.smaskData!)]).toEqual([255, 0, 0, 255]);
  });

  it('paints the see-through pixels white when PDF/A-1 forbids the mask', () => {
    const prepared = prepareImage(gif(true), { flattenAlpha: true });
    expect(prepared.smaskData).toBeUndefined();
    expect([...unzlibSync(prepared.data)].slice(3, 6)).toEqual([255, 255, 255]);
  });
});
