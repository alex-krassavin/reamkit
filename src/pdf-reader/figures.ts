// The figures a page draws, and the words set on them.
//
// A drawing is not a paragraph, and neither are the words written on it. A
// paper's state machine is a dozen boxes and arrows with a word or two in each
// box — "Interpret / Bytecodes", "Monitor", "side exit to existing trace" — set
// in 5pt Helvetica beside a column of 9pt Times. Read as the column's text, the
// labels were twenty-odd short lines of it, and they cut the page into so many
// columns that it was taken for a table; the drawing was anchored to the page,
// and went wherever its anchor went. comments.pdf came back five pages longer
// than it is, with its figure a page after the words that label it.
//
// So a figure is found before the page's text is read: the paths that make one
// drawing, the pictures in it, and the words set on it, which are its labels
// and not lines of the text.

import type { TextRun } from './content';
import type { PdfImage } from './images';
import type { PdfVector } from './vector';

/** One drawing on a page, with the words set on it. Page space, y-up. */
export interface PageFigure {
  readonly minX: number;
  readonly minY: number;
  readonly maxX: number;
  readonly maxY: number;
  /** The paths that draw it, in no particular order. */
  readonly vectors: ReadonlyArray<PdfVector>;
  /** The pictures it holds. */
  readonly images: ReadonlyArray<PdfImage>;
  /** The words set on it. */
  readonly labels: ReadonlyArray<TextRun>;
}

interface Box {
  minX: number;
  minY: number;
  maxX: number;
  maxY: number;
}

/** No thicker than this, a path is a line. */
const THIN_PT = 1.5;

/**
 * A line longer than this many ems RULES the page: the rule over a caption,
 * the line under a heading, a table's border. A drawing's own strokes are
 * shorter, and a long one inside a drawing — a chart's axis — is taken with
 * the drawing it stands in once that is found.
 */
const RULE_EM = 3;

/** How close, in ems, two paths stand to be parts of one drawing. */
const JOIN_EM = 0.5;

/** The fewest paths a figure is drawn with: a box and a frame are not one. */
const MIN_PATHS = 6;

/** Past this share of the sheet, a drawing is the page's ground or frame. */
const MOST_OF_PAGE = 0.5;

/** How far across a figure a line of its words may run and still be a label. */
const LABEL_FILL = 0.6;

/** Past this many lines that run across it, what a drawing holds is prose. */
const PROSE_LINES = 2;

/**
 * How far apart, in ems of the text, two parts of one figure may stand. The
 * parts of comments.pdf's Figure 6 stand twenty-six points apart at 9pt.
 */
const NEAR_EM = 3;

/** Type smaller than this share of the text's is a drawing's, not the text's. */
const SMALL_SHARE = 0.85;

/** How much larger than the text a figure's words may be set. */
const LARGEST_LABEL = 1.2;

/** How far, in its own ems, small type may stand off a drawing and label it. */
const CAPTION_EM = 2;

/** What a glyph mapped to no character reads as (see `UNMAPPED` in ./layout). */
const UNMAPPED = '\uFFFD';

/** Where an annotation's appearance stands in the page's painting order. */
const ANNOTATION_ORDER = Number.MAX_SAFE_INTEGER;

/**
 * How many ems of its own type a line of words on a drawing may run before it
 * is a line of prose, however wide the drawing: a label is a word or three.
 */
const LABEL_EM = 15;

/** Past this many paths a page is artwork through and through. */
const MOST_PATHS = 4000;

/**
 * The figures a page draws.
 *
 * A figure is paths that stand together — each within half an em of the next —
 * and DRAW something: a curve or a slanted line among them. Straight lines
 * and boxes alone are a ruling, which is a table's or a form's, and a page
 * reads those as text. It has words set on it, or it is only artwork, which
 * the page anchors as it always has; and none of its words run across it
 * line after line, which is prose set on a ground, not a label.
 *
 * @param vectors The page's paths, where the page shows them.
 * @param images  The page's pictures, likewise.
 * @param runs    The page's text, likewise.
 * @param sheet   The page's size.
 * @returns The figures, top of the page first.
 */
export function pageFigures(
  vectors: ReadonlyArray<PdfVector>,
  images: ReadonlyArray<PdfImage>,
  runs: ReadonlyArray<TextRun>,
  sheet: { readonly width: number; readonly height: number },
): Array<PageFigure> {
  // §9.10.2 — a run of glyphs that map to no character says nothing, and its
  // outlines are drawn as paths where it stands.
  const words = runs.filter((r) => r.text.replaceAll(UNMAPPED, '').trim().length > 0);
  if (words.length === 0 || vectors.length === 0 || vectors.length > MOST_PATHS) return [];
  const size = median(words.map((r) => r.fontSizePt).filter((s) => s > 0)) || 10;
  const gap = size * JOIN_EM;
  const pageArea = Math.max(1, sheet.width * sheet.height);
  const huge = (b: Box): boolean => (b.maxX - b.minX) * (b.maxY - b.minY) > pageArea * MOST_OF_PAGE;
  const thin = (v: PdfVector): boolean => Math.min(v.maxX - v.minX, v.maxY - v.minY) <= THIN_PT;
  const rule = (v: PdfVector): boolean =>
    thin(v) && Math.max(v.maxX - v.minX, v.maxY - v.minY) > size * RULE_EM;
  // §9.6.5 — a Type 3 face draws its glyphs as paths, and §9.6.6 a glyph the
  // file maps to no character is traced as one: those are letters, and a line
  // of them is not a drawing. TAMReview.pdf sets its every running head in
  // Type 3 bitmaps, and each one came back a figure.
  const lettering = words.filter((r) => r.type3 === true).map(inkOf);
  const letter = (v: PdfVector): boolean =>
    v.glyph === true ||
    lettering.some(
      (b) =>
        (v.minX + v.maxX) / 2 >= b.minX - 1 &&
        (v.minX + v.maxX) / 2 <= b.maxX + 1 &&
        (v.minY + v.maxY) / 2 >= b.minY - 1 &&
        (v.minY + v.maxY) / 2 <= b.maxY + 1,
    );
  // §12.5.5 — what an annotation's appearance draws is laid over the page,
  // not part of it: comments.pdf bars a strip across the head of a page, and
  // taken with the figure it touched it made one drawing of the page's text.
  const own = vectors.filter((v) => v.orderKey[0] !== ANNOTATION_ORDER);
  const seeds = own.filter((v) => !rule(v) && !huge(v) && !letter(v));

  // Paths within `gap` of each other, joined: sorted across the page, each is
  // compared only with those that start before it ends.
  const parent = seeds.map((_, k) => k);
  const find = (k: number): number => {
    while (parent[k] !== k) {
      parent[k] = parent[parent[k]!]!;
      k = parent[k]!;
    }
    return k;
  };
  const order = seeds.map((_, k) => k).sort((a, b) => seeds[a]!.minX - seeds[b]!.minX);
  for (let i = 0; i < order.length; i++) {
    const a = seeds[order[i]!]!;
    for (let j = i + 1; j < order.length; j++) {
      const b = seeds[order[j]!]!;
      if (b.minX > a.maxX + gap) break;
      if (b.minY <= a.maxY + gap && b.maxY >= a.minY - gap) {
        parent[find(order[i]!)] = find(order[j]!);
      }
    }
  }
  const clusters = new Map<number, Array<PdfVector>>();
  seeds.forEach((v, k) => {
    const root = find(k);
    const members = clusters.get(root);
    if (members) members.push(v);
    else clusters.set(root, [v]);
  });

  const found: Array<{ box: Box; vectors: Array<PdfVector> }> = [];
  // What is too little to be a figure on its own may still be part of one.
  const pieces: Array<{ box: Box; vectors: Array<PdfVector> }> = [];
  for (const members of clusters.values()) {
    const box = boxOf(members);
    if (huge(box)) continue;
    if (members.length < MIN_PATHS || !(members.some(draws) || barChart(members))) {
      pieces.push({ box, vectors: members });
    } else found.push({ box, vectors: members });
  }
  if (found.length === 0) return [];

  // Drawings that overlap are one: a key beside a diagram, a diagram on a
  // ground. So are the parts of one figure, set a little apart with nothing of
  // the text between them: Figure 7's (a) and (b) stand sixteen points apart
  // side by side, and set one under the other they would not be the figure.
  // Two figures with a caption between them stay two. A few paths standing
  // near one are its own, too: Figure 8 draws its outer tree's head five
  // points off the rest, and left on the page the head stayed where the page
  // had it while the figure moved. And a long line that reaches into one — a
  // chart's axis — is part of it.
  const near = size * NEAR_EM;
  const text = words.filter((r) => r.type3 !== true && r.fontSizePt >= size * SMALL_SHARE);
  const textBetween = (a: Box, b: Box): boolean => {
    const both = union(a, b);
    return text.some((r) => {
      const ink = inkOf(r);
      return touches(both, ink, 0) && !touches(a, ink, 0) && !touches(b, ink, 0);
    });
  };
  const rules = own.filter(rule);
  const taken = new Set<PdfVector>();
  let changed = true;
  while (changed) {
    changed = false;
    for (let i = 0; i < found.length; i++) {
      for (let j = found.length - 1; j > i; j--) {
        const [a, b] = [found[i]!.box, found[j]!.box];
        if (!touches(a, b, gap) && !(touches(a, b, near) && !textBetween(a, b))) continue;
        found[i]!.vectors.push(...found[j]!.vectors);
        found[i]!.box = union(a, b);
        found.splice(j, 1);
        changed = true;
      }
      for (let j = pieces.length - 1; j >= 0; j--) {
        const [a, b] = [found[i]!.box, pieces[j]!.box];
        if (!touches(a, b, gap) && !(touches(a, b, near) && !textBetween(a, b))) continue;
        found[i]!.vectors.push(...pieces[j]!.vectors);
        found[i]!.box = union(a, b);
        pieces.splice(j, 1);
        changed = true;
      }
      for (const r of rules) {
        if (taken.has(r) || !mostlyWithin(r, found[i]!.box)) continue;
        taken.add(r);
        found[i]!.vectors.push(r);
        found[i]!.box = union(found[i]!.box, r);
        changed = true;
      }
    }
  }

  // The words on a drawing are its labels, and so is small type just past its
  // edge: the "(a)" under a part, the "Closed" under a bar. What is set at the
  // size of the text is the text's, whatever stands near it.
  const figures: Array<PageFigure> = [];
  const labelled = new Set<TextRun>();
  for (const f of found) {
    const free = words.filter((r) => r.type3 !== true && !labelled.has(r));
    const labels = free.filter((r) => touches(f.box, inkOf(r), 0));
    let box = labels.reduce((b, r) => union(b, inkOf(r)), f.box);
    for (let grown = true; grown; ) {
      grown = false;
      for (const r of free) {
        if (labels.includes(r) || r.fontSizePt >= size * SMALL_SHARE) continue;
        const ink = inkOf(r);
        if (!touches(box, ink, r.fontSizePt * CAPTION_EM) || ink.minX < box.minX - r.fontSizePt)
          continue;
        if (ink.maxX > box.maxX + r.fontSizePt) continue;
        labels.push(r);
        box = union(box, ink);
        grown = true;
      }
    }
    if (labels.length === 0 || runsAcross(labels, box, size)) continue;
    // A label is set no larger than the text it stands among. What sets its
    // words larger is a page drawn round its words — bug1771477.pdf is a web
    // page printed, its "hello world" in 17pt on a card beside 10pt links —
    // and a page, not a figure on one.
    if (labels.some((r) => r.fontSizePt > size * LARGEST_LABEL)) continue;
    // …and the spaces between the words are the labels' too. Left on the
    // page, the one in comments.pdf's "loop edge" was taken into the line of
    // the column beside it, and moved that line's baseline three points.
    const spaces = runs.filter(
      (r) => r.text.trim().length === 0 && r.text.length > 0 && touches(box, inkOf(r), 0),
    );
    for (const r of [...labels, ...spaces]) labelled.add(r);
    const held = images.filter((img) => mostlyInside(img, box));
    figures.push({ ...box, vectors: f.vectors, images: held, labels: [...labels, ...spaces] });
  }
  return figures.sort((a, b) => b.maxY - a.maxY);
}

/** Whether a path draws: a curve, or a line that is neither level nor upright. */
function draws(v: PdfVector): boolean {
  let x = 0;
  let y = 0;
  for (const s of v.segs) {
    if (s.op === 'cubic') return true;
    if (s.op === 'line' && Math.abs(s.x - x) > THIN_PT && Math.abs(s.y - y) > THIN_PT) return true;
    if (s.op !== 'close') {
      x = s.x;
      y = s.y;
    }
  }
  return false;
}

/**
 * Whether paths draw a BAR CHART: filled rectangles, six and more, standing on
 * one line and as long as each other's figures make them — the one drawing
 * made of nothing but straight lines and boxes that a ruling is not. A
 * table's shading stands its boxes on one line too, but each as long as the
 * last.
 *
 * comments.pdf sets Figures 10 and 12 so, bars of gradient and of flat grey.
 * With no curve among them they were taken for a ruling: their labels were
 * read as lines of the text, Figure 12 and the column beside it came back as
 * one table, and Figure 10's caption, the first line left on its page, set
 * the page's top margin 352 points down.
 *
 * @param paths The paths of one cluster.
 */
function barChart(paths: ReadonlyArray<PdfVector>): boolean {
  const bars = paths
    .filter((v) => v.fillHex !== undefined || v.gradient !== undefined)
    .flatMap(rectanglesOf);
  // Bars stand up from a line, or run out from one.
  const standing = (base: (b: Box) => number, length: (b: Box) => number): boolean => {
    const lengths = new Map<number, Array<number>>();
    for (const b of bars) {
      const at = Math.round(base(b) / BAR_LINE_PT);
      const long = length(b);
      if (long > THIN_PT) lengths.set(at, [...(lengths.get(at) ?? []), long]);
    }
    return [...lengths.values()].some(
      (all) => all.length >= MIN_PATHS && Math.max(...all) >= Math.min(...all) * BAR_SPREAD,
    );
  };
  return (
    standing(
      (b) => b.minY,
      (b) => b.maxY - b.minY,
    ) ||
    standing(
      (b) => b.minX,
      (b) => b.maxX - b.minX,
    )
  );
}

/** How near, in points, bars' feet stand to stand on one line. */
const BAR_LINE_PT = 0.5;

/** How much longer the longest bar of a chart is than the shortest, at the least. */
const BAR_SPREAD = 2;

/** The upright rectangles a path is drawn with, one per subpath that is one. */
function rectanglesOf(v: PdfVector): Array<Box> {
  const out: Array<Box> = [];
  let corners: Array<{ x: number; y: number }> = [];
  let curved = false;
  const close = (): void => {
    const xs = new Set(corners.map((c) => Math.round(c.x * 10)));
    const ys = new Set(corners.map((c) => Math.round(c.y * 10)));
    if (!curved && corners.length >= 4 && corners.length <= 5 && xs.size === 2 && ys.size === 2) {
      out.push({
        minX: Math.min(...corners.map((c) => c.x)),
        maxX: Math.max(...corners.map((c) => c.x)),
        minY: Math.min(...corners.map((c) => c.y)),
        maxY: Math.max(...corners.map((c) => c.y)),
      });
    }
    corners = [];
    curved = false;
  };
  for (const s of v.segs) {
    if (s.op === 'move') {
      close();
      corners.push({ x: s.x, y: s.y });
    } else if (s.op === 'line') corners.push({ x: s.x, y: s.y });
    else if (s.op === 'cubic') curved = true;
  }
  close();
  return out;
}

/**
 * Whether the words on a drawing are prose: lines that run across most of it,
 * more than a caption's worth. A label is short, a paragraph set on a ground
 * fills it.
 *
 * A line is its words as far as the next gap wider than an em. Two boxes side
 * by side put their labels on one baseline — comments.pdf's "Record" and
 * "Enter" a hundred and twenty points apart — and taken as one line they ran
 * across the figure like a line of its text.
 */
function runsAcross(labels: ReadonlyArray<TextRun>, box: Box, size: number): boolean {
  const byLine: Array<{ y: number; em: number; inks: Array<Box> }> = [];
  for (const r of labels) {
    const line = byLine.find((l) => Math.abs(l.y - r.y) <= size * 0.3);
    if (line) {
      line.inks.push(inkOf(r));
      line.em = Math.max(line.em, r.fontSizePt);
    } else byLine.push({ y: r.y, em: r.fontSizePt, inks: [inkOf(r)] });
  }
  const width = box.maxX - box.minX;
  let across = 0;
  for (const { em, inks } of byLine) {
    inks.sort((a, b) => a.minX - b.minX);
    let from = inks[0]!.minX;
    let to = inks[0]!.maxX;
    for (const ink of [...inks.slice(1), undefined]) {
      if (ink !== undefined && ink.minX - to <= size) {
        to = Math.max(to, ink.maxX);
        continue;
      }
      if (to - from >= Math.min(width * LABEL_FILL, em * LABEL_EM)) across++;
      if (ink !== undefined) {
        from = ink.minX;
        to = ink.maxX;
      }
    }
  }
  return across >= PROSE_LINES;
}

/** The box a run's glyphs stand in: a fifth of the size below the baseline, the rest above. */
function inkOf(r: TextRun): Box {
  const s = r.fontSizePt;
  if (Math.abs(r.angleDeg ?? 0) < 0.5) {
    return {
      minX: Math.min(r.x, r.endX),
      minY: r.y - s * 0.2,
      maxX: Math.max(r.x, r.endX),
      maxY: r.y + s * 0.8,
    };
  }
  return {
    minX: Math.min(r.x, r.endX) - s * 0.8,
    minY: Math.min(r.y, r.endY) - s * 0.8,
    maxX: Math.max(r.x, r.endX) + s * 0.8,
    maxY: Math.max(r.y, r.endY) + s * 0.8,
  };
}

function boxOf(members: ReadonlyArray<Box>): Box {
  return members.reduce(union, { ...members[0]! });
}

function union(a: Box, b: Box): Box {
  return {
    minX: Math.min(a.minX, b.minX),
    minY: Math.min(a.minY, b.minY),
    maxX: Math.max(a.maxX, b.maxX),
    maxY: Math.max(a.maxY, b.maxY),
  };
}

function touches(a: Box, b: Box, slack: number): boolean {
  return (
    b.minX <= a.maxX + slack &&
    b.maxX >= a.minX - slack &&
    b.minY <= a.maxY + slack &&
    b.maxY >= a.minY - slack
  );
}

/**
 * Whether a rule lies along a drawing rather than meeting it: across it within
 * a point, and at least half its length inside it. A table's border that ends
 * where a figure begins is not the figure's.
 */
function mostlyWithin(r: Box, box: Box): boolean {
  const across = r.maxX - r.minX >= r.maxY - r.minY;
  const [from, to, lo, hi] = across
    ? [r.minX, r.maxX, box.minX, box.maxX]
    : [r.minY, r.maxY, box.minY, box.maxY];
  const [side, sideEnd, sideLo, sideHi] = across
    ? [r.minY, r.maxY, box.minY, box.maxY]
    : [r.minX, r.maxX, box.minX, box.maxX];
  if (sideEnd < sideLo - 1 || side > sideHi + 1) return false;
  return Math.min(to, hi) - Math.max(from, lo) >= (to - from) / 2;
}

/** Whether most of a picture lies inside a box. */
function mostlyInside(img: PdfImage, box: Box): boolean {
  const w = Math.max(0, Math.min(img.x + img.widthPt, box.maxX) - Math.max(img.x, box.minX));
  const h = Math.max(0, Math.min(img.y + img.heightPt, box.maxY) - Math.max(img.y, box.minY));
  return w * h > (img.widthPt * img.heightPt) / 2;
}

function median(values: ReadonlyArray<number>): number {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.floor(sorted.length / 2)]!;
}
