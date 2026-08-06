// E-PPTX PX2 — the placeholder cascade (slide → slideLayout → slideMaster).
//
// A slide placeholder shape (p:sp with a p:ph) usually omits its own geometry
// and text formatting, inheriting them from the matching prototype in the
// slide's layout, then the master:
//   * geometry   — the layout placeholder's a:xfrm, else the master's;
//   * the anchor — the layout placeholder's a:bodyPr, else the master's;
//   * text       — four sources, nearest last: the presentation's
//     p:defaultTextStyle (§19.2.1.8, what a plain text box gets), the master's
//     p:txStyles for the placeholder's family, then the master's and the
//     layout's own prototype a:lstStyle. Each states both a paragraph's shape
//     (alignment, indent, spacing) and its runs' (size, colour, face).
//
// This module builds a resolver from the (already parsed) layout and master
// trees; the slide parser consults it whenever a slide shape lacks its own
// transform or formatting.

import type { Alignment, ParagraphProperties, RunProperties } from '@/core/document-model';
import type { ColorResolver } from '@/core/drawingml/colors';
import type { ThemeFonts } from '@/core/drawingml/theme-parser';
import type { PoNode } from '@/core/po-helpers';
import type { PlaceholderRef, ShapeBoxEmu } from '@/pptx/sp-helpers';

import type { Pt } from '@/core/ir';
import { emuToPt, pt } from '@/core/ir';
import { resolveColorNode } from '@/core/drawingml/colors';
import { poAttr, poChildren, poIntAttr, poIs } from '@/core/po-helpers';
import { fromSymbolFont } from '@/core/metafile/symbol-fonts';
import { parsePh, parseXfrmBox, rPrToRunProps } from '@/pptx/sp-helpers';

/**
 * The placeholder cascade (PX2): the resolver a slide parser consults whenever a
 * slide placeholder shape (`p:sp` with a `p:ph`) lacks its own geometry or run
 * formatting, supplying the inherited values from the slide's layout, then its
 * master.
 */
export interface PlaceholderCascade {
  /** The inherited EMU box for a slide placeholder without its own `a:xfrm`. */
  readonly geometryFor: (ph: PlaceholderRef) => ShapeBoxEmu | undefined;
  /**
   * The inherited run formatting for a paragraph at the given 0-based outline
   * level. Without a placeholder this is the presentation's default text style
   * — what an ordinary text box on a slide is written in.
   */
  readonly defaultsFor: (ph: PlaceholderRef | undefined, level: number) => RunProperties;
  /** The inherited paragraph shape at that level: alignment, indent, spacing. */
  readonly paragraphDefaultsFor: (
    ph: PlaceholderRef | undefined,
    level: number,
  ) => ParagraphProperties;
  /** The bullet that level inherits, which a paragraph states only to change. */
  readonly bulletFor: (ph: PlaceholderRef | undefined, level: number) => LevelBullet | undefined;
  /**
   * The vertical anchor the placeholder's prototype states (`a:bodyPr@anchor`),
   * the layout's before the master's, for a shape that states none itself.
   */
  readonly anchorFor: (ph: PlaceholderRef) => 't' | 'ctr' | 'b' | undefined;
  /**
   * The prototypes' own `p:spPr`, nearest first — the layout's, then the
   * master's — for the properties a slide placeholder inherits rather than
   * states: its fill, its outline, its geometry.
   *
   * A CHAIN and not one node, because the inheritance is per property. The
   * layout's prototype may state a box and nothing else, and the fill then
   * comes from the master: tdf104015's title is red in every reader and was a
   * bare outline here, because the layout's `p:spPr` existed and stopped the
   * search.
   */
  readonly shapePropsFor: (ph: PlaceholderRef) => ReadonlyArray<PoNode>;
}

/** One level of a text style: how its paragraphs sit and how their runs read. */
interface LevelStyle {
  readonly run: RunProperties;
  readonly paragraph: ParagraphProperties;
  readonly bullet?: LevelBullet;
  /**
   * §21.1.2.2.9/.10 — the space around the paragraph stated as a FRACTION of a
   * line rather than a distance. A fraction of WHAT is only known once the
   * levels have merged and the size is settled, so it travels this far raw.
   */
  readonly spacingPct?: { readonly before?: number; readonly after?: number };
}

/**
 * §21.1.2.4.4/.5/.6 — the bullet a level declares: none at all, a literal
 * character, or a number that counts. A slide paragraph states one only where
 * it differs from this: themes.pptx's second slide writes one bare line, and
 * the dot in front of it is the master's body style, nine levels up.
 */
export interface LevelBullet {
  /** What it draws: nothing at all, a character, or a number that counts. */
  readonly kind?: 'none' | 'char' | 'autoNum';
  /** `a:buChar` — the character, already read out of the face that states it. */
  readonly char?: string;
  /** `a:buAutoNum @type`/`@startAt` — the numbering and where it starts. */
  readonly type?: string;
  readonly startAt?: number;
  /** §21.1.2.4.2/.3 — its size, as a fraction of the text or in points. */
  readonly sizePct?: number;
  readonly sizePts?: number;
  /** §21.1.2.4.1 `a:buClr` — the colour it is drawn in, when it has its own. */
  readonly colorHex?: string;
}

/**
 * What an `a:pPr` (or an `a:lvlNpPr`, the same vocabulary) says about its
 * bullet — each part on its own, because a paragraph may restate the SIZE and
 * leave the character to its level.
 *
 * @param pPr    The paragraph-properties node.
 * @param colors The colour resolver, for `a:buClr`.
 * @returns What it states, or `undefined` when it says nothing about bullets.
 */
export function parseBullet(
  pPr: PoNode | undefined,
  colors?: ColorResolver,
): LevelBullet | undefined {
  if (!pPr) return undefined;
  const child = (tag: string): PoNode | undefined => poChildren(pPr).find((c) => poIs(c, tag));
  const pct = poIntAttr(child('a:buSzPct'), 'val');
  const pts = poIntAttr(child('a:buSzPts'), 'val');
  const clr = child('a:buClr');
  const colorHex =
    clr && colors
      ? poChildren(clr)
          .map((c) => resolveColorNode(c, colors))
          .find((hex) => hex !== undefined)
      : undefined;
  const size: LevelBullet = {
    ...(pct !== undefined ? { sizePct: pct / 100000 } : {}),
    ...(pts !== undefined ? { sizePts: pts / 100 } : {}),
    ...(colorHex !== undefined ? { colorHex } : {}),
  };
  const buChar = child('a:buChar');
  const buAuto = child('a:buAutoNum');
  const kind: LevelBullet | undefined = poChildren(pPr).some((c) => poIs(c, 'a:buNone'))
    ? { kind: 'none' }
    : buChar
      ? // §21.1.2.4.5 `a:buFont` — the character is stated IN THAT FACE, and the
        // symbol faces are not alphabets: Wingdings `l` is a filled circle, not
        // a letter. 45541_Header's every bullet is one, and drawn in the text's
        // own font it printed a column of `l`s down the slide.
        {
          kind: 'char',
          char: fromSymbolFont(
            poAttr(buChar, 'char') ?? '•',
            poAttr(child('a:buFont'), 'typeface'),
          ),
        }
      : buAuto
        ? {
            kind: 'autoNum',
            type: poAttr(buAuto, 'type') ?? 'arabicPeriod',
            startAt: poIntAttr(buAuto, 'startAt') ?? 1,
          }
        : undefined;
  const out = { ...size, ...kind };
  return Object.keys(out).length > 0 ? out : undefined;
}

type StyleCategory = 'title' | 'body' | 'other';

interface ParsedPlaceholder {
  readonly ref: PlaceholderRef;
  readonly box?: ShapeBoxEmu;
  /** The prototype's `p:spPr` — its fill, outline and geometry. */
  readonly spPr?: PoNode;
  /** The prototype's own `a:lstStyle`, per level. */
  readonly levels: ReadonlyArray<LevelStyle>;
  /** `a:bodyPr@anchor` — where the text sits vertically in the box. */
  readonly anchor?: 't' | 'ctr' | 'b';
}

// §19.3.1.* — p:ph @type buckets onto one of the master's three text-style
// families. Absent type defaults to 'obj' (a content placeholder → body).
function categoryOf(type: string | undefined): StyleCategory {
  if (type === 'title' || type === 'ctrTitle') return 'title';
  if (type === undefined || type === 'body' || type === 'subTitle' || type === 'obj') return 'body';
  return 'other';
}

/**
 * Build a {@link PlaceholderCascade} from a slide's already-parsed layout and
 * master trees: it resolves a placeholder's inherited geometry (the layout's
 * matching `a:xfrm`, else the master's) and per-level default run formatting (the
 * master's `p:txStyles`).
 *
 * @param layoutTree The parsed `p:sldLayout` part.
 * @param masterTree The parsed `p:sldMaster` part, when present.
 * @param colors     The deck's colour resolver, for scheme colours in the text styles.
 * @param deckDefaults The presentation's own default text style, as the floor.
 * @param themeFonts The theme's font scheme, for the `+mn-lt` tokens.
 */
export function buildPlaceholderCascade(
  layoutTree: ReadonlyArray<PoNode>,
  masterTree: ReadonlyArray<PoNode> | undefined,
  colors: ColorResolver,
  deckDefaults: ReadonlyArray<LevelStyle> = [],
  themeFonts?: ThemeFonts,
): PlaceholderCascade {
  const layoutPhs = collectPlaceholders(layoutTree, 'p:sldLayout', colors, themeFonts);
  const masterPhs = masterTree
    ? collectPlaceholders(masterTree, 'p:sldMaster', colors, themeFonts)
    : [];
  const txStyles = masterTree
    ? collectTxStyles(masterTree, colors, themeFonts)
    : { title: [], body: [], other: [] };

  // Nearest last: the deck's default, the master's family style, then the two
  // prototypes' own lists. Each layer states only what it changes.
  const styleAt = (ph: PlaceholderRef | undefined, level: number): LevelStyle => {
    const layers = [at(deckDefaults, level)];
    if (ph) {
      layers.push(
        at(txStyles[categoryOf(ph.type)], level),
        at(matchPlaceholder(masterPhs, ph)?.levels ?? [], level),
        at(matchPlaceholder(layoutPhs, ph)?.levels ?? [], level),
      );
    }
    const merged = layers.reduce(mergeLevels, EMPTY_LEVEL);
    return merged.spacingPct === undefined ? merged : withResolvedSpacing(merged);
  };

  return {
    geometryFor(ph) {
      return matchPlaceholder(layoutPhs, ph)?.box ?? matchPlaceholder(masterPhs, ph)?.box;
    },
    defaultsFor(ph, level) {
      return styleAt(ph, level).run;
    },
    paragraphDefaultsFor(ph, level) {
      return styleAt(ph, level).paragraph;
    },
    bulletFor(ph, level) {
      return styleAt(ph, level).bullet;
    },
    anchorFor(ph) {
      return matchPlaceholder(layoutPhs, ph)?.anchor ?? matchPlaceholder(masterPhs, ph)?.anchor;
    },
    shapePropsFor(ph) {
      const chain = [matchPlaceholder(layoutPhs, ph)?.spPr, matchPlaceholder(masterPhs, ph)?.spPr];
      return chain.filter((n): n is PoNode => n !== undefined);
    },
  };
}

const EMPTY_LEVEL: LevelStyle = { run: {}, paragraph: {} };

/**
 * §21.1.2.2.9/.10 — the space a paragraph states as a FRACTION of a line, in
 * points.
 *
 * A "line" here is the paragraph's own text at its natural height, which is the
 * size times the same 1.2 the layout gives an unstated line. A level that
 * states the fraction and leaves the size to another level is resolved once
 * they have merged, which is why this is not read where the fraction is.
 *
 * @param pct      The fraction, `0..n` (20% arrives as 0.2).
 * @param sizePt   The paragraph's resolved font size.
 * @returns The distance, or `undefined` when the size is not known.
 */
export function spacingFromLineFraction(
  pct: number | undefined,
  sizePt: number | undefined,
): Pt | undefined {
  if (pct === undefined || sizePt === undefined) return undefined;
  return pt(pct * sizePt * NATURAL_LINE);
}

/** What the layout gives a line of text with nothing said about its height. */
const NATURAL_LINE = 1.2;

// The merged level with its line-fraction spacing turned into distances. A
// fraction the size cannot resolve is dropped rather than guessed at.
function withResolvedSpacing(level: LevelStyle): LevelStyle {
  const size = level.run.fontSizePt;
  const before = spacingFromLineFraction(level.spacingPct?.before, size);
  const after = spacingFromLineFraction(level.spacingPct?.after, size);
  if (before === undefined && after === undefined) return level;
  return {
    ...level,
    paragraph: {
      ...level.paragraph,
      ...(before !== undefined ? { spacingBefore: before } : {}),
      ...(after !== undefined ? { spacingAfter: after } : {}),
    },
  };
}

/** The style at `level`, clamped to the list — a deeper level keeps the last. */
function at(levels: ReadonlyArray<LevelStyle>, level: number): LevelStyle {
  if (levels.length === 0) return EMPTY_LEVEL;
  return levels[Math.min(Math.max(level, 0), levels.length - 1)] ?? EMPTY_LEVEL;
}

/** `next` over `base` — a layer states only what it changes. */
function mergeLevels(base: LevelStyle, next: LevelStyle): LevelStyle {
  // Each PART of a bullet inherits on its own: a level may restate the size and
  // leave the character to the one above it.
  const bullet = base.bullet || next.bullet ? { ...base.bullet, ...next.bullet } : undefined;
  const spacingPct =
    base.spacingPct || next.spacingPct ? { ...base.spacingPct, ...next.spacingPct } : undefined;
  return {
    run: { ...base.run, ...next.run },
    paragraph: { ...base.paragraph, ...next.paragraph },
    ...(bullet ? { bullet } : {}),
    ...(spacingPct ? { spacingPct } : {}),
  };
}

/**
 * `a:lstStyle` (or a `p:txStyles` family) → its nine levels. Shared by the
 * deck's default text style, the master's families and a prototype's own list,
 * which are the same vocabulary written in different places.
 *
 * @param list   The `a:lstStyle` / `p:titleStyle` / … node.
 * @param colors The deck's colour resolver.
 * @param prefix The level element prefix: `a:lvl` here, always.
 * @returns The nine levels, or an empty list when there is no such node.
 */
export function parseLevelStyles(
  list: PoNode | undefined,
  colors: ColorResolver,
  themeFonts?: ThemeFonts,
): Array<LevelStyle> {
  if (!list) return [];
  const out: Array<LevelStyle> = [];
  for (let lvl = 1; lvl <= 9; lvl++) {
    const lvlPr = poChildren(list).find((c) => poIs(c, `a:lvl${String(lvl)}pPr`));
    const defRPr = lvlPr ? poChildren(lvlPr).find((c) => poIs(c, 'a:defRPr')) : undefined;
    const bullet = parseBullet(lvlPr, colors);
    const spacingPct = spacingFractions(lvlPr);
    out.push({
      run: rPrToRunProps(defRPr, colors, themeFonts),
      paragraph: pPrToParagraphProps(lvlPr),
      ...(bullet ? { bullet } : {}),
      ...(spacingPct ? { spacingPct } : {}),
    });
  }
  return out;
}

const ALGN: Readonly<Record<string, Alignment>> = {
  l: 'left',
  ctr: 'center',
  r: 'right',
  just: 'both',
  dist: 'distribute',
};

/**
 * A level's paragraph shape — §21.1.2.2.7 `@algn`, `@marL`/`@indent`, and the
 * space around and between its lines.
 *
 * Both spacings come in two spellings: `a:spcPts` is a distance (hundredths of
 * a point) and `a:spcPct` a FRACTION of a line (thousandths of a percent). Only
 * the distance was read, so a deck that spaces its bullets the usual way — by
 * fraction — set them solid.
 */
export function pPrToParagraphProps(lvlPr: PoNode | undefined): ParagraphProperties {
  if (!lvlPr) return {};
  const algn = poAttr(lvlPr, 'algn');
  const marL = poIntAttr(lvlPr, 'marL');
  const indent = poIntAttr(lvlPr, 'indent');
  return {
    ...(algn !== undefined && ALGN[algn] ? { alignment: ALGN[algn] } : {}),
    ...(marL !== undefined ? { indentLeft: emuToPt(marL) } : {}),
    ...(indent !== undefined ? { indentFirstLine: emuToPt(indent) } : {}),
    ...spacingPt(lvlPr, 'a:spcBef', 'spacingBefore'),
    ...spacingPt(lvlPr, 'a:spcAft', 'spacingAfter'),
    ...lineSpacing(lvlPr),
  };
}

/**
 * §21.1.2.2.5 `a:lnSpc` — the height of each line: a fraction of the natural
 * one (`a:spcPct`) or a distance (`a:spcPts`).
 *
 * The model states a multiple the way §17.3.1.33 does — 12pt under the `auto`
 * rule IS single — so a fraction needs no font size to express, which is why
 * this resolves here and the paragraph SPACING has to wait for the size.
 *
 * @param pPr The paragraph properties (a level's or a paragraph's own).
 * @returns The line spacing, or nothing when the node states none.
 */
export function lineSpacing(pPr: PoNode | undefined): ParagraphProperties {
  const holder = pPr ? poChildren(pPr).find((c) => poIs(c, 'a:lnSpc')) : undefined;
  if (!holder) return {};
  const pctVal = poIntAttr(
    poChildren(holder).find((c) => poIs(c, 'a:spcPct')),
    'val',
  );
  if (pctVal !== undefined && pctVal > 0) {
    return { spacingLine: pt((pctVal / 100000) * 12), spacingLineRule: 'auto' };
  }
  const ptsVal = poIntAttr(
    poChildren(holder).find((c) => poIs(c, 'a:spcPts')),
    'val',
  );
  if (ptsVal !== undefined && ptsVal > 0) {
    return { spacingLine: pt(ptsVal / 100), spacingLineRule: 'exact' };
  }
  return {};
}

/**
 * §21.1.2.2.9/.10 — the fraction of a line a paragraph puts before and after
 * itself, when it states one. The distance form is read beside it; this is what
 * the caller must still resolve against the size.
 *
 * @param pPr The paragraph properties.
 * @returns The two fractions (`0..n`), each present only when stated.
 */
export function spacingFractions(
  pPr: PoNode | undefined,
): { readonly before?: number; readonly after?: number } | undefined {
  const frac = (tag: string): number | undefined => {
    const holder = pPr ? poChildren(pPr).find((c) => poIs(c, tag)) : undefined;
    const pctNode = holder ? poChildren(holder).find((c) => poIs(c, 'a:spcPct')) : undefined;
    const val = pctNode ? poIntAttr(pctNode, 'val') : undefined;
    return val === undefined || val < 0 ? undefined : val / 100000;
  };
  const [before, after] = [frac('a:spcBef'), frac('a:spcAft')];
  if (before === undefined && after === undefined) return undefined;
  return { ...(before !== undefined ? { before } : {}), ...(after !== undefined ? { after } : {}) };
}

// a:spcBef/a:spcAft → the points it states, when it states points.
function spacingPt(
  lvlPr: PoNode,
  tag: string,
  key: 'spacingBefore' | 'spacingAfter',
): ParagraphProperties {
  const holder = poChildren(lvlPr).find((c) => poIs(c, tag));
  const pts = holder ? poChildren(holder).find((c) => poIs(c, 'a:spcPts')) : undefined;
  const val = pts ? poIntAttr(pts, 'val') : undefined;
  return val === undefined ? {} : { [key]: pt(val / 100) };
}

// Collect every placeholder shape (with its optional geometry) from a layout or
// master shape tree, in document order.
function collectPlaceholders(
  tree: ReadonlyArray<PoNode>,
  root: 'p:sldLayout' | 'p:sldMaster',
  colors: ColorResolver,
  themeFonts?: ThemeFonts,
): Array<ParsedPlaceholder> {
  const sld = tree.find((n) => poIs(n, root));
  const cSld = sld ? poChildren(sld).find((c) => poIs(c, 'p:cSld')) : undefined;
  const spTree = cSld ? poChildren(cSld).find((c) => poIs(c, 'p:spTree')) : undefined;
  if (!spTree) return [];
  const out: Array<ParsedPlaceholder> = [];
  for (const sp of poChildren(spTree)) {
    if (!poIs(sp, 'p:sp')) continue;
    const ref = parsePh(sp);
    if (!ref) continue;
    const spPr = poChildren(sp).find((c) => poIs(c, 'p:spPr'));
    const box = parseXfrmBox(spPr);
    const txBody = poChildren(sp).find((c) => poIs(c, 'p:txBody'));
    const lstStyle = txBody ? poChildren(txBody).find((c) => poIs(c, 'a:lstStyle')) : undefined;
    const bodyPr = txBody ? poChildren(txBody).find((c) => poIs(c, 'a:bodyPr')) : undefined;
    const a = bodyPr ? poAttr(bodyPr, 'anchor') : undefined;
    const anchor = a === 'ctr' || a === 'b' || a === 't' ? a : undefined;
    out.push({
      ref,
      ...(box ? { box } : {}),
      ...(spPr ? { spPr } : {}),
      levels: parseLevelStyles(lstStyle, colors, themeFonts),
      ...(anchor ? { anchor } : {}),
    });
  }
  return out;
}

// Match a slide placeholder to a layout/master prototype: by idx when present
// (the canonical join key), else by exact type, else by style category (so a
// title matches the master's single title prototype even if the exact type
// differs, e.g. ctrTitle ↔ title).
function matchPlaceholder(
  list: ReadonlyArray<ParsedPlaceholder>,
  ph: PlaceholderRef,
): ParsedPlaceholder | undefined {
  if (ph.idx !== undefined) {
    const byIdx = list.find((p) => p.ref.idx === ph.idx);
    if (byIdx) return byIdx;
  }
  const byType = list.find((p) => p.ref.type === ph.type);
  if (byType) return byType;
  const cat = categoryOf(ph.type);
  return list.find((p) => categoryOf(p.ref.type) === cat);
}

// The master's p:txStyles → per-level default run properties for each family.
function collectTxStyles(
  masterTree: ReadonlyArray<PoNode>,
  colors: ColorResolver,
  themeFonts?: ThemeFonts,
): Record<StyleCategory, Array<LevelStyle>> {
  const sld = masterTree.find((n) => poIs(n, 'p:sldMaster'));
  const txStyles = sld ? poChildren(sld).find((c) => poIs(c, 'p:txStyles')) : undefined;
  const family = (tag: string): Array<LevelStyle> =>
    parseLevelStyles(
      txStyles ? poChildren(txStyles).find((c) => poIs(c, tag)) : undefined,
      colors,
      themeFonts,
    );
  return {
    title: family('p:titleStyle'),
    body: family('p:bodyStyle'),
    other: family('p:otherStyle'),
  };
}
