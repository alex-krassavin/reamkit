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
import type { PoNode } from '@/core/po-helpers';
import type { PlaceholderRef, ShapeBoxEmu } from '@/pptx/sp-helpers';

import { emuToPt, pt } from '@/core/ir';
import { poAttr, poChildren, poIntAttr, poIs } from '@/core/po-helpers';
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
  /**
   * The vertical anchor the placeholder's prototype states (`a:bodyPr@anchor`),
   * the layout's before the master's, for a shape that states none itself.
   */
  readonly anchorFor: (ph: PlaceholderRef) => 't' | 'ctr' | 'b' | undefined;
  /**
   * The prototype's own `p:spPr`, for the properties a slide placeholder
   * inherits rather than states: its fill, its outline, its geometry. The
   * layout's before the master's.
   */
  readonly shapePropsFor: (ph: PlaceholderRef) => PoNode | undefined;
}

/** One level of a text style: how its paragraphs sit and how their runs read. */
interface LevelStyle {
  readonly run: RunProperties;
  readonly paragraph: ParagraphProperties;
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
 */
export function buildPlaceholderCascade(
  layoutTree: ReadonlyArray<PoNode>,
  masterTree: ReadonlyArray<PoNode> | undefined,
  colors: ColorResolver,
  deckDefaults: ReadonlyArray<LevelStyle> = [],
): PlaceholderCascade {
  const layoutPhs = collectPlaceholders(layoutTree, 'p:sldLayout', colors);
  const masterPhs = masterTree ? collectPlaceholders(masterTree, 'p:sldMaster', colors) : [];
  const txStyles = masterTree
    ? collectTxStyles(masterTree, colors)
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
    return layers.reduce(mergeLevels, EMPTY_LEVEL);
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
    anchorFor(ph) {
      return matchPlaceholder(layoutPhs, ph)?.anchor ?? matchPlaceholder(masterPhs, ph)?.anchor;
    },
    shapePropsFor(ph) {
      return matchPlaceholder(layoutPhs, ph)?.spPr ?? matchPlaceholder(masterPhs, ph)?.spPr;
    },
  };
}

const EMPTY_LEVEL: LevelStyle = { run: {}, paragraph: {} };

/** The style at `level`, clamped to the list — a deeper level keeps the last. */
function at(levels: ReadonlyArray<LevelStyle>, level: number): LevelStyle {
  if (levels.length === 0) return EMPTY_LEVEL;
  return levels[Math.min(Math.max(level, 0), levels.length - 1)] ?? EMPTY_LEVEL;
}

/** `next` over `base` — a layer states only what it changes. */
function mergeLevels(base: LevelStyle, next: LevelStyle): LevelStyle {
  return {
    run: { ...base.run, ...next.run },
    paragraph: { ...base.paragraph, ...next.paragraph },
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
): Array<LevelStyle> {
  if (!list) return [];
  const out: Array<LevelStyle> = [];
  for (let lvl = 1; lvl <= 9; lvl++) {
    const lvlPr = poChildren(list).find((c) => poIs(c, `a:lvl${String(lvl)}pPr`));
    const defRPr = lvlPr ? poChildren(lvlPr).find((c) => poIs(c, 'a:defRPr')) : undefined;
    out.push({ run: rPrToRunProps(defRPr, colors), paragraph: pPrToParagraphProps(lvlPr) });
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
 * space around it. `a:spcPct` is a fraction of the line rather than a distance,
 * which this model has no room for, so only the stated POINTS are read.
 */
function pPrToParagraphProps(lvlPr: PoNode | undefined): ParagraphProperties {
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
  };
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
      levels: parseLevelStyles(lstStyle, colors),
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
): Record<StyleCategory, Array<LevelStyle>> {
  const sld = masterTree.find((n) => poIs(n, 'p:sldMaster'));
  const txStyles = sld ? poChildren(sld).find((c) => poIs(c, 'p:txStyles')) : undefined;
  const family = (tag: string): Array<LevelStyle> =>
    parseLevelStyles(txStyles ? poChildren(txStyles).find((c) => poIs(c, tag)) : undefined, colors);
  return {
    title: family('p:titleStyle'),
    body: family('p:bodyStyle'),
    other: family('p:otherStyle'),
  };
}
