// E-PPTX PX2 — the placeholder cascade (slide → slideLayout → slideMaster).
//
// A slide placeholder shape (p:sp with a p:ph) usually omits its own geometry
// and text sizing, inheriting them from the matching prototype in the slide's
// layout, then the master:
//   * geometry  — the layout placeholder's a:xfrm, else the master's;
//   * text size/colour — the master's p:txStyles (titleStyle / bodyStyle /
//     otherStyle), per outline level.
//
// This module builds a resolver from the (already parsed) layout and master
// trees; the slide parser consults it whenever a slide shape lacks its own
// transform or run formatting.

import type { RunProperties } from '@/core/document-model';
import type { PoNode } from '@/core/po-helpers';
import type { PlaceholderRef, ShapeBoxEmu } from '@/pptx/sp-helpers';

import { poChildren, poIs } from '@/core/po-helpers';
import { parsePh, parseXfrmBox, rPrToRunProps } from '@/pptx/sp-helpers';

export interface PlaceholderCascade {
  // The inherited EMU box for a slide placeholder without its own a:xfrm.
  readonly geometryFor: (ph: PlaceholderRef) => ShapeBoxEmu | undefined;
  // The master's default run formatting for a placeholder's paragraph at the
  // given 0-based outline level (empty when no master text styles apply).
  readonly defaultsFor: (ph: PlaceholderRef, level: number) => RunProperties;
}

type StyleCategory = 'title' | 'body' | 'other';

interface ParsedPlaceholder {
  readonly ref: PlaceholderRef;
  readonly box?: ShapeBoxEmu;
}

// §19.3.1.* — p:ph @type buckets onto one of the master's three text-style
// families. Absent type defaults to 'obj' (a content placeholder → body).
function categoryOf(type: string | undefined): StyleCategory {
  if (type === 'title' || type === 'ctrTitle') return 'title';
  if (type === undefined || type === 'body' || type === 'subTitle' || type === 'obj') return 'body';
  return 'other';
}

export function buildPlaceholderCascade(
  layoutTree: ReadonlyArray<PoNode>,
  masterTree?: ReadonlyArray<PoNode>,
): PlaceholderCascade {
  const layoutPhs = collectPlaceholders(layoutTree, 'p:sldLayout');
  const masterPhs = masterTree ? collectPlaceholders(masterTree, 'p:sldMaster') : [];
  const txStyles = masterTree ? collectTxStyles(masterTree) : { title: [], body: [], other: [] };

  return {
    geometryFor(ph) {
      return matchPlaceholder(layoutPhs, ph)?.box ?? matchPlaceholder(masterPhs, ph)?.box;
    },
    defaultsFor(ph, level) {
      const levels = txStyles[categoryOf(ph.type)];
      if (levels.length === 0) return {};
      const i = Math.min(Math.max(level, 0), levels.length - 1);
      return levels[i] ?? {};
    },
  };
}

// Collect every placeholder shape (with its optional geometry) from a layout or
// master shape tree, in document order.
function collectPlaceholders(
  tree: ReadonlyArray<PoNode>,
  root: 'p:sldLayout' | 'p:sldMaster',
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
    out.push({ ref, ...(box ? { box } : {}) });
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
): Record<StyleCategory, Array<RunProperties>> {
  const sld = masterTree.find((n) => poIs(n, 'p:sldMaster'));
  const txStyles = sld ? poChildren(sld).find((c) => poIs(c, 'p:txStyles')) : undefined;
  return {
    title: levelStyles(txStyles, 'p:titleStyle'),
    body: levelStyles(txStyles, 'p:bodyStyle'),
    other: levelStyles(txStyles, 'p:otherStyle'),
  };
}

// p:titleStyle / p:bodyStyle / p:otherStyle → the a:lvl1pPr…a:lvl9pPr default
// run properties (a:defRPr), indexed by level (0-based).
function levelStyles(txStyles: PoNode | undefined, tag: string): Array<RunProperties> {
  const style = txStyles ? poChildren(txStyles).find((c) => poIs(c, tag)) : undefined;
  if (!style) return [];
  const out: Array<RunProperties> = [];
  for (let lvl = 1; lvl <= 9; lvl++) {
    const lvlPr = poChildren(style).find((c) => poIs(c, `a:lvl${lvl}pPr`));
    const defRPr = lvlPr ? poChildren(lvlPr).find((c) => poIs(c, 'a:defRPr')) : undefined;
    out.push(rPrToRunProps(defRPr));
  }
  return out;
}
