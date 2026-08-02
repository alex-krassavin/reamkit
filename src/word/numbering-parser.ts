// ECMA-376 Part 1 §17.9 — Numbering.xml parser.
//
// Numbering is a two-level indirection:
//   abstractNum (defines levels: numFmt, lvlText, start, indent, …)
//   num         (instance binding a numId → abstractNumId, with optional
//                per-level overrides — overrides not yet implemented).
// A paragraph references a list via <w:numPr> { numId, ilvl } in its pPr.

import { XMLParser } from 'fast-xml-parser';

import type {
  AbstractNumbering,
  Numbering,
  NumberingFormat,
  NumberingInstance,
  NumberingLevel,
  PictureBullet,
} from '@/core/document-model';
import type { ResourceId } from '@/core/ir';
import { pt } from '@/core/ir';

import { parseParagraphProperties } from '@/word/paragraph-properties';
import { parseRunProperties } from '@/word/run-properties';
import { asArray, asElement, getAttr } from '@/word/xml-helpers';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  // §4.1 of XML 1.0: a numeric character reference is not an entity — `&#10;`
  // IS a line feed and every parser must decode it. fast-xml-parser gates that
  // on `htmlEntities`, which defaults to false, so `&#10;` reached the page as
  // five literal characters (formats.xlsx writes "Hello,&#10;Calc!"). Named
  // HTML entities come along with the switch; in XML they are undefined anyway,
  // and reading `&nbsp;` as a space beats drawing it. Nested DOCTYPE entities
  // stay unexpanded either way — the parser never registers them (54764-2.xlsx).
  htmlEntities: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
});

const FORMATS = new Set<NumberingFormat>([
  'decimal',
  'decimalZero',
  'decimalFullWidth',
  'ordinal',
  'lowerLetter',
  'upperLetter',
  'lowerRoman',
  'upperRoman',
  'ideographTraditional',
  'ideographZodiac',
  'ideographLegalTraditional',
  'ideographDigital',
  'koreanDigital2',
  'chineseCounting',
  'chineseCountingThousand',
  'japaneseCounting',
  'koreanCounting',
  'taiwaneseCounting',
  'taiwaneseCountingThousand',
  'hebrew1',
  'hebrew2',
  'decimalEnclosedCircle',
  'chicago',
  'bullet',
  'none',
]);

/** The empty {@link Numbering} returned when a document has no `numbering.xml`. */
export const EMPTY_NUMBERING: Numbering = {
  abstractNums: new Map(),
  numInstances: new Map(),
};

/**
 * Parse `word/numbering.xml` (ECMA-376 Part 1 §17.9) into a {@link Numbering}:
 * the abstract numbering definitions (each level's format, start, text template
 * and indent/run properties) plus the num instances that bind a `numId` to an
 * `abstractNumId`. `basedOn` inheritance and per-level overrides are not resolved
 * here; an unparseable format defaults to `decimal`.
 *
 * @param data The raw `numbering.xml` bytes.
 * @returns The parsed numbering, or {@link EMPTY_NUMBERING} when the root is absent.
 */
export function parseNumbering(
  data: Uint8Array,
  resolveImage?: (relId: string) => ResourceId | undefined,
): Numbering {
  const xml = decoder.decode(data);
  const tree = parser.parse(xml) as Record<string, unknown>;
  const root = asElement(tree['w:numbering']);
  if (!root) return EMPTY_NUMBERING;

  // §17.9.21 — the picture bullets, by id, before the levels that name them.
  const picBullets = parsePicBullets(root, resolveImage);

  const abstractNums = new Map<string, AbstractNumbering>();
  for (const a of asArray(root['w:abstractNum'])) {
    const el = asElement(a);
    if (!el) continue;
    const id = getAttr(el, 'abstractNumId');
    if (!id) continue;
    const levels = new Map<number, NumberingLevel>();
    for (const lvlNode of asArray(el['w:lvl'])) {
      const lvl = parseLevel(lvlNode, picBullets);
      if (lvl) levels.set(lvl.ilvl, lvl);
    }
    abstractNums.set(id, { id, levels });
  }

  const numInstances = new Map<string, NumberingInstance>();
  for (const n of asArray(root['w:num'])) {
    const el = asElement(n);
    if (!el) continue;
    const numId = getAttr(el, 'numId');
    if (!numId) continue;
    const abstractNumId = getValVal(el['w:abstractNumId']);
    if (!abstractNumId) continue;
    // §17.9.27/§17.9.28 — the instance may start a level somewhere other than
    // the abstract definition does. num-override-start.docx starts its second
    // level at three, and the abstract start alone numbered its one heading
    // "1.1" where its own text reads "This should be 1.3".
    const startOverrides = new Map<number, number>();
    // …and §17.9.27 lets it redefine the level WHOLE. NumberingWOverrides.docx
    // rewrites all nine levels of one instance, and reading the abstract's
    // instead numbered its "B" and "C" items 1 and 2 at the wrong level.
    const levelOverrides = new Map<number, NumberingLevel>();
    for (const o of asArray(el['w:lvlOverride'])) {
      const ovr = asElement(o);
      if (!ovr) continue;
      const ilvl = Number(getAttr(ovr, 'ilvl'));
      if (!Number.isFinite(ilvl)) continue;
      const start = Number(getValVal(ovr['w:startOverride']));
      if (Number.isFinite(start)) startOverrides.set(ilvl, start);
      const lvl = parseLevel(ovr['w:lvl'], picBullets);
      if (lvl) levelOverrides.set(ilvl, lvl);
    }
    numInstances.set(numId, {
      numId,
      abstractNumId,
      ...(startOverrides.size > 0 ? { startOverrides } : {}),
      ...(levelOverrides.size > 0 ? { levelOverrides } : {}),
    });
  }

  return { abstractNums, numInstances };
}

// §17.9.6 `w:lvl` — one level of a list: where it starts, how it counts, the
// template it prints and the indent/run properties it lends the paragraph.
// Spelled the same inside an abstract definition and inside a `w:lvlOverride`.
function parseLevel(
  node: unknown,
  picBullets: ReadonlyMap<string, PictureBullet>,
): NumberingLevel | undefined {
  const lvlEl = asElement(node);
  if (!lvlEl) return undefined;
  const ilvlStr = getAttr(lvlEl, 'ilvl');
  if (!ilvlStr) return undefined;
  const ilvl = Number(ilvlStr);
  if (!Number.isFinite(ilvl)) return undefined;

  // §17.9.25 — a level that states no `w:start` starts at ZERO, not one.
  // FDO74105.docx omits it and LibreOffice numbers its list from 0.
  const startAttr = getValVal(lvlEl['w:start']);
  const start = startAttr !== undefined ? Number(startAttr) : 0;
  const fmtStr = getValVal(lvlEl['w:numFmt']) ?? 'decimal';
  const format: NumberingFormat = FORMATS.has(fmtStr as NumberingFormat)
    ? (fmtStr as NumberingFormat)
    : 'decimal';
  const lvlText = getValVal(lvlEl['w:lvlText']) ?? '';
  const picId = getValVal(lvlEl['w:lvlPicBulletId']);
  const picBullet = picId !== undefined ? picBullets.get(picId) : undefined;
  // §17.9.10 `w:isLgl` — legal numbering: every level of this level's marker
  // prints in decimal. listWithLgl.docx numbers its chapters in Roman and its
  // sections "Sect 1.01"; we wrote "Sect I.01".
  const isLegal = 'w:isLgl' in lvlEl && getValVal(lvlEl['w:isLgl']) !== '0';

  return {
    ilvl,
    start: Number.isFinite(start) ? start : 1,
    format,
    lvlText,
    ...(isLegal ? { isLegal: true } : {}),
    ...(picBullet ? { picBullet } : {}),
    paragraphProperties: parseParagraphProperties(lvlEl['w:pPr']),
    runProperties: parseRunProperties(lvlEl['w:rPr']),
  };
}

// §17.9.21 `w:numPicBullet` — each holds a `w:pict/v:shape` sized in CSS units
// by its `style`, wrapping the `v:imagedata` that names the image. Keyed by
// `@w:numPicBulletId`, which a level's `w:lvlPicBulletId` points at.
function parsePicBullets(
  root: Record<string, unknown>,
  resolveImage?: (relId: string) => ResourceId | undefined,
): Map<string, PictureBullet> {
  const out = new Map<string, PictureBullet>();
  for (const node of asArray(root['w:numPicBullet'])) {
    const el = asElement(node);
    if (!el) continue;
    const id = getAttr(el, 'numPicBulletId');
    if (id === undefined) continue;
    const pict = asElement(el['w:pict']);
    const shape = pict ? asElement(pict['v:shape']) : undefined;
    if (!shape) continue;
    const data = asElement(shape['v:imagedata']);
    const relId = data ? (getAttr(data, 'id') ?? getAttr(data, 'r:id')) : undefined;
    const style = getAttr(shape, 'style') ?? '';
    const widthPt = cssLengthPt(/(?:^|;)\s*width\s*:\s*([^;]+)/u.exec(style)?.[1]);
    const heightPt = cssLengthPt(/(?:^|;)\s*height\s*:\s*([^;]+)/u.exec(style)?.[1]);
    if (widthPt === undefined || heightPt === undefined) continue;
    // §17.9.21 — a picture bullet with no picture is not a bullet: Word falls
    // back to the level's own `w:lvlText` glyph. lvlPicBulletId.docx declares a
    // three-INCH `v:shape` with no `v:imagedata` at all, and reserving that box
    // for every item of its contents list stretched one page into six.
    const resource = relId !== undefined ? resolveImage?.(relId) : undefined;
    if (resource === undefined) continue;
    out.set(id, { resource, widthPt: pt(widthPt), heightPt: pt(heightPt) });
  }
  return out;
}

// A VML `style` length — `11.25pt`, `3in`, `24px`, a bare number of points.
const CSS_UNITS: ReadonlyMap<string, number> = new Map([
  ['pt', 1],
  ['in', 72],
  ['cm', 72 / 2.54],
  ['mm', 72 / 25.4],
  ['pc', 12],
  ['px', 0.75],
]);

function cssLengthPt(raw: string | undefined): number | undefined {
  const m = raw !== undefined ? /^\s*(-?[\d.]+)\s*([a-z%]*)\s*$/u.exec(raw) : null;
  if (!m) return undefined;
  const n = Number(m[1]);
  if (!Number.isFinite(n)) return undefined;
  const unit = m[2] === '' ? 'pt' : m[2]!;
  const scale = CSS_UNITS.get(unit);
  return scale === undefined ? undefined : n * scale;
}

function getValVal(node: unknown): string | undefined {
  if (!node || typeof node !== 'object') return undefined;
  const v =
    (node as Record<string, unknown>)['@_w:val'] ?? (node as Record<string, unknown>)['@_val'];
  return typeof v === 'string' ? v : undefined;
}
