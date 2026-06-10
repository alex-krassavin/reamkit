// Plain-text renderer that paginates paragraphs and embeds a TrueType font.
//
// Line breaking is greedy using real font metrics from the parsed TTF.
// Replaced in M4 by Knuth-Plass paragraph breaking.

import type { ParsedTtf } from '@/core/font';
import type { EmbeddedFont } from '@/pdf/cid-font';
import type { PdfDict, PdfRef } from '@/pdf/objects';
import { parseTtf } from '@/core/font';
import { embedTtfFont } from '@/pdf/cid-font';
import { dict, name, ref, stream } from '@/pdf/objects';
import { PdfDocument } from '@/pdf/writer';

export type FontInput = ParsedTtf | { readonly bytes: Uint8Array };

export interface TextRenderOptions {
  readonly font: FontInput;
  readonly pageWidth?: number;
  readonly pageHeight?: number;
  readonly marginLeft?: number;
  readonly marginRight?: number;
  readonly marginTop?: number;
  readonly marginBottom?: number;
  readonly fontSize?: number;
  readonly lineHeight?: number;
}

const A4_WIDTH = 595;
const A4_HEIGHT = 842;

const encoder = new TextEncoder();

export function renderPlainTextPdf(
  paragraphs: ReadonlyArray<string>,
  options: TextRenderOptions,
): Uint8Array {
  const parsed = isParsedTtf(options.font) ? options.font : parseTtf(options.font.bytes);

  const pageWidth = options.pageWidth ?? A4_WIDTH;
  const pageHeight = options.pageHeight ?? A4_HEIGHT;
  const marginLeft = options.marginLeft ?? 72;
  const marginRight = options.marginRight ?? 72;
  const marginTop = options.marginTop ?? 72;
  const marginBottom = options.marginBottom ?? 72;
  const fontSize = options.fontSize ?? 12;
  const lineHeight = options.lineHeight ?? fontSize * 1.2;

  const contentWidth = pageWidth - marginLeft - marginRight;
  const contentHeight = pageHeight - marginTop - marginBottom;
  const linesPerPage = Math.max(1, Math.floor(contentHeight / lineHeight));

  const usedGids = collectUsedGids(paragraphs, parsed);

  const doc = new PdfDocument();
  const embedded = embedTtfFont(doc, parsed, { usedGids });

  const lines = paginate(paragraphs, contentWidth, fontSize, embedded);
  const pages: Array<Array<string>> = [];
  for (let i = 0; i < lines.length; i += linesPerPage) {
    pages.push(lines.slice(i, i + linesPerPage));
  }
  if (pages.length === 0) pages.push([]);

  const pagesDict: PdfDict = dict({
    Type: name('Pages'),
    Count: 0,
    Kids: [],
  });
  const pagesRef = doc.add(pagesDict);

  const pageRefs: Array<PdfRef> = [];
  for (const pageLines of pages) {
    const contentBytes = buildPageContentStream({
      lines: pageLines,
      marginLeft,
      marginTop,
      pageHeight,
      fontSize,
      lineHeight,
      embedded,
    });
    const contentsRef = doc.add(stream({}, contentBytes));
    const pageRef = doc.add(
      dict({
        Type: name('Page'),
        Parent: ref(pagesRef.id),
        MediaBox: [0, 0, pageWidth, pageHeight],
        Resources: dict({
          Font: dict({ F1: ref(embedded.fontRef.id) }),
        }),
        Contents: ref(contentsRef.id),
      }),
    );
    pageRefs.push(pageRef);
  }

  pagesDict.set('Count', pageRefs.length);
  pagesDict.set('Kids', pageRefs);

  const catalogRef = doc.add(
    dict({
      Type: name('Catalog'),
      Pages: ref(pagesRef.id),
    }),
  );

  return doc.build(catalogRef);
}

function isParsedTtf(font: FontInput): font is ParsedTtf {
  return 'unitsPerEm' in font;
}

function collectUsedGids(paragraphs: ReadonlyArray<string>, parsed: ParsedTtf): Set<number> {
  const gids = new Set<number>();
  for (const para of paragraphs) {
    for (const ch of para) {
      const cp = ch.codePointAt(0)!;
      gids.add(parsed.glyphForCodepoint(cp));
    }
  }
  return gids;
}

interface ContentStreamOptions {
  readonly lines: ReadonlyArray<string>;
  readonly marginLeft: number;
  readonly marginTop: number;
  readonly pageHeight: number;
  readonly fontSize: number;
  readonly lineHeight: number;
  readonly embedded: EmbeddedFont;
}

function buildPageContentStream(opts: ContentStreamOptions): Uint8Array {
  if (opts.lines.length === 0) return new Uint8Array(0);

  const baselineY = opts.pageHeight - opts.marginTop - opts.fontSize;
  const lines: Array<string> = [];
  lines.push('BT');
  lines.push(`/F1 ${opts.fontSize} Tf`);
  lines.push(`${opts.lineHeight} TL`);
  lines.push(`${opts.marginLeft} ${baselineY} Td`);
  for (let i = 0; i < opts.lines.length; i++) {
    if (i > 0) lines.push('T*');
    const hex = opts.embedded.encodeTextAsCidHex(opts.lines[i]!);
    lines.push(`<${hex}> Tj`);
  }
  lines.push('ET');
  return encoder.encode(lines.join('\n'));
}

function paginate(
  paragraphs: ReadonlyArray<string>,
  maxWidth: number,
  fontSize: number,
  embedded: EmbeddedFont,
): Array<string> {
  const out: Array<string> = [];
  for (const p of paragraphs) {
    if (p.length === 0) {
      out.push('');
      continue;
    }
    out.push(...wrapGreedy(p, maxWidth, fontSize, embedded));
  }
  return out;
}

function wrapGreedy(
  text: string,
  maxWidth: number,
  fontSize: number,
  embedded: EmbeddedFont,
): Array<string> {
  const words = text.split(/\s+/).filter((w) => w.length > 0);
  if (words.length === 0) return [''];
  const lines: Array<string> = [];
  let current = '';
  for (const word of words) {
    const candidate = current.length === 0 ? word : `${current} ${word}`;
    if (embedded.textWidthPt(candidate, fontSize) <= maxWidth) {
      current = candidate;
    } else {
      if (current.length > 0) lines.push(current);
      current = word;
    }
  }
  if (current.length > 0) lines.push(current);
  return lines;
}
