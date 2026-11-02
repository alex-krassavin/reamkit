// OpenType Layout run as a program: the GSUB and GPOS lookups of a face,
// applied to a buffer of glyphs the way a shaping engine applies them.
//
// `opentype-layout.ts` reads a face's layout tables into three flat maps — its
// ligatures, its pair kerning, its Arabic positional forms — and that is all
// Latin, Cyrillic and Arabic text has needed. A script of India needs the rest
// of the machinery. Its conjuncts and the shapes of its vowel signs are chosen
// by CONTEXT: a rule matches a sequence of glyphs and runs other lookups at
// places in it (GSUB LookupType 5 and 6). Each lookup steps over the glyphs its
// flags tell it to: the marks, or all but one class of them. The marks are then
// placed by anchors on the glyph they belong to (GPOS LookupType 4 and 6), off
// the pen position, above or below the line.
//
// This module is that machinery and nothing more. Which features run, in what
// order and on which glyphs is the shaper's business. Where the specification
// leaves a choice open, this follows HarfBuzz, the engine LibreOffice, Chrome
// and Android shape with, so that the glyphs come out as theirs do: which
// glyphs a match may skip, a join control included, the syllable a match keeps
// to, and how the places in a rule move when a lookup inside it changes the
// length of the buffer.

import type { ParsedTtf } from '@/core/font/ttf-parser';

/** One glyph in the buffer the lookups run over. */
export interface OtlGlyph {
  /** The glyph id. */
  gid: number;
  /** Index of the first character it was made from; a ligature takes the least. */
  cluster: number;
  /** The features that may act on it, one bit each (see {@link OtlStep.mask}). */
  mask: number;
  /** Its GDEF class: 0 unclassified, 1 base, 2 ligature, 3 mark, 4 component. */
  cls: number;
  /** The syllable it belongs to, or 0; a per-syllable lookup matches within one. */
  syllable: number;
  /** A join control, which a match may step over (see {@link OtlStep}). */
  joiner?: 'zwj' | 'zwnj';
  /** The ligature it is, or the ligature a mark inside it was skipped by; 0 none. */
  ligId: number;
  /** For a mark skipped inside a ligature: the component it followed, from 1. */
  ligComp: number;
  /**
   * What substitution made it, as {@link SUBSTITUTED}, {@link LIGATED} and
   * {@link MULTIPLIED} bits: a shaper asks whether a reph or a pre-base form
   * actually formed. Absent ⇒ none.
   */
  props?: number;
}

/** {@link OtlGlyph.props}: a GSUB lookup replaced the glyph. */
export const SUBSTITUTED = 0x1;
/** {@link OtlGlyph.props}: …by a ligature, last of all. */
export const LIGATED = 0x2;
/** {@link OtlGlyph.props}: …by a multiple substitution, last of all. */
export const MULTIPLIED = 0x4;

/** Where a glyph is drawn, in font units: its advance, and its offset off the pen. */
export interface OtlPosition {
  xAdvance: number;
  yAdvance: number;
  xOffset: number;
  yOffset: number;
  /** The glyph it hangs on, as the distance to it in the buffer; 0 none. */
  attachChain: number;
  /** How it hangs there: 1 a mark on its anchor, 2 a cursive join. */
  attachType: 0 | 1 | 2;
}

/** One lookup to run, with the glyphs it may act on and how it matches. */
export interface OtlStep {
  /** Its index in the table's LookupList. */
  readonly lookup: number;
  /** It acts on a glyph whose mask shares a bit with this one. */
  readonly mask: number;
  /** A ZWJ between the glyphs of a match is stepped over. */
  readonly autoZwj: boolean;
  /** A ZWNJ between the glyphs of a context is stepped over. */
  readonly autoZwnj: boolean;
  /** A match keeps to the syllable of the glyph it starts on. */
  readonly perSyllable: boolean;
}

/** A feature a shaper asks for: its tag, and how its lookups are to run. */
export interface OtlFeature {
  readonly tag: string;
  readonly mask: number;
  /** The font's own rules decide what a ZWJ does, so a match does not skip it. */
  readonly manualZwj?: boolean;
  /** …and what a ZWNJ does. */
  readonly manualZwnj?: boolean;
  readonly perSyllable?: boolean;
}

// LookupFlag bits.
const RIGHT_TO_LEFT = 0x1;
const IGNORE_FLAGS = 0xe; // IgnoreBaseGlyphs | IgnoreLigatures | IgnoreMarks
const IGNORE_MARKS = 0x8;
const USE_MARK_FILTERING_SET = 0x10;
const MARK_ATTACHMENT_TYPE = 0xff00;

/** GDEF glyph classes. */
const MARK = 3;

// A rule inside a rule inside a rule: a font that nests deeper is broken, or
// built to make the engine spin.
const MAX_NESTING = 6;
// The most glyphs a buffer may hold after substitution, as a multiple of what
// it started with; a multiple substitution that feeds itself stops here.
const MAX_GROWTH = 8;
// The most lookups one run applies before it gives up on the buffer.
const MAX_OPS_PER_GLYPH = 1024;

interface LangSys {
  readonly required: number;
  readonly features: ReadonlyArray<number>;
}

interface Script {
  readonly defaultLangSys?: LangSys;
  readonly langSys: ReadonlyMap<string, LangSys>;
}

interface Lookup {
  /** The type of its subtables, with an Extension unwrapped. */
  readonly type: number;
  readonly flag: number;
  readonly markSet: number;
  /** Absolute offsets of its subtables, an Extension's target in its place. */
  readonly subtables: ReadonlyArray<number>;
}

/** A GSUB or GPOS table, read as far as its lists. */
export interface LayoutTable {
  readonly kind: 'GSUB' | 'GPOS';
  readonly scripts: ReadonlyMap<string, Script>;
  readonly features: ReadonlyArray<{
    readonly tag: string;
    readonly lookups: ReadonlyArray<number>;
  }>;
  readonly lookups: ReadonlyArray<Lookup>;
}

const u16 = (b: Uint8Array, o: number): number => (b[o]! << 8) | b[o + 1]!;
const i16 = (b: Uint8Array, o: number): number => {
  const v = u16(b, o);
  return v >= 0x8000 ? v - 0x10000 : v;
};
const u32 = (b: Uint8Array, o: number): number =>
  b[o]! * 0x1000000 + ((b[o + 1]! << 16) | (b[o + 2]! << 8) | b[o + 3]!);
const tag = (b: Uint8Array, o: number): string =>
  String.fromCharCode(b[o]!, b[o + 1]!, b[o + 2]!, b[o + 3]!);

/**
 * A face's layout tables, read once: GDEF's classes and the lists of GSUB and
 * GPOS, with each coverage table and class definition read the first time a
 * lookup asks for it.
 */
export class OtlFace {
  private static readonly cache = new WeakMap<ParsedTtf, OtlFace>();

  /** The face's GSUB, when it has one that reads. */
  readonly gsub?: LayoutTable;
  /** …and its GPOS. */
  readonly gpos?: LayoutTable;

  private readonly raw: Uint8Array;
  private readonly glyphClasses: ReadonlyMap<number, number>;
  private readonly attachClasses: ReadonlyMap<number, number>;
  private readonly markSets: ReadonlyArray<number>;
  private readonly coverages = new Map<number, ReadonlyMap<number, number>>();
  private readonly classDefs = new Map<number, ReadonlyMap<number, number>>();

  private constructor(readonly parsed: ParsedTtf) {
    this.raw = parsed.raw;
    const gdef = parsed.tables.get('GDEF');
    let glyphClasses: ReadonlyMap<number, number> = new Map();
    let attachClasses: ReadonlyMap<number, number> = new Map();
    let markSets: Array<number> = [];
    if (gdef) {
      try {
        const b = gdef.offset;
        const minor = u16(this.raw, b + 2);
        const classOff = u16(this.raw, b + 4);
        const attachOff = u16(this.raw, b + 10);
        if (classOff) glyphClasses = this.classDef(b + classOff);
        if (attachOff) attachClasses = this.classDef(b + attachOff);
        const setsOff = minor >= 2 ? u16(this.raw, b + 12) : 0;
        if (setsOff) {
          const s = b + setsOff;
          const count = u16(this.raw, s + 2);
          for (let i = 0; i < count; i++) markSets.push(s + u32(this.raw, s + 4 + 4 * i));
        }
      } catch {
        glyphClasses = new Map();
        attachClasses = new Map();
        markSets = [];
      }
    }
    this.glyphClasses = glyphClasses;
    this.attachClasses = attachClasses;
    this.markSets = markSets;
    const gsub = this.readTable('GSUB');
    const gpos = this.readTable('GPOS');
    if (gsub) this.gsub = gsub;
    if (gpos) this.gpos = gpos;
  }

  /**
   * The layout tables of a face, read on first use and kept with it.
   *
   * @param parsed The face.
   * @returns Its {@link OtlFace}.
   */
  static of(parsed: ParsedTtf): OtlFace {
    let face = OtlFace.cache.get(parsed);
    if (!face) {
      face = new OtlFace(parsed);
      OtlFace.cache.set(parsed, face);
    }
    return face;
  }

  /** A glyph's GDEF class: 1 base, 2 ligature, 3 mark, 4 component, 0 none. */
  glyphClass(gid: number): number {
    return this.glyphClasses.get(gid) ?? 0;
  }

  /** A mark's attachment class, which a lookup's flag may select. */
  attachClass(gid: number): number {
    return this.attachClasses.get(gid) ?? 0;
  }

  /** Whether a glyph is in one of GDEF's mark glyph sets. */
  inMarkSet(set: number, gid: number): boolean {
    const off = this.markSets[set];
    return off !== undefined && this.coverage(off).has(gid);
  }

  /** @internal The coverage table at an absolute offset: glyph → coverage index. */
  coverage(off: number): ReadonlyMap<number, number> {
    let cov = this.coverages.get(off);
    if (cov) return cov;
    const b = this.raw;
    const map = new Map<number, number>();
    const format = u16(b, off);
    if (format === 1) {
      const n = u16(b, off + 2);
      for (let i = 0; i < n; i++) map.set(u16(b, off + 4 + 2 * i), i);
    } else if (format === 2) {
      const n = u16(b, off + 2);
      for (let i = 0; i < n; i++) {
        const r = off + 4 + 6 * i;
        const start = u16(b, r);
        const end = u16(b, r + 2);
        const index = u16(b, r + 4);
        for (let g = start; g <= end; g++) map.set(g, index + g - start);
      }
    }
    cov = map;
    this.coverages.set(off, cov);
    return cov;
  }

  /** @internal The class definition at an absolute offset: glyph → class, 0 unlisted. */
  classDef(off: number): ReadonlyMap<number, number> {
    let def = this.classDefs.get(off);
    if (def) return def;
    const b = this.raw;
    const map = new Map<number, number>();
    const format = u16(b, off);
    if (format === 1) {
      const start = u16(b, off + 2);
      const n = u16(b, off + 4);
      for (let i = 0; i < n; i++) {
        const c = u16(b, off + 6 + 2 * i);
        if (c) map.set(start + i, c);
      }
    } else if (format === 2) {
      const n = u16(b, off + 2);
      for (let i = 0; i < n; i++) {
        const r = off + 4 + 6 * i;
        const c = u16(b, r + 4);
        if (!c) continue;
        for (let g = u16(b, r); g <= u16(b, r + 2); g++) map.set(g, c);
      }
    }
    def = map;
    this.classDefs.set(off, def);
    return def;
  }

  /** @internal The raw bytes of the face. */
  get bytes(): Uint8Array {
    return this.raw;
  }

  private readTable(kind: 'GSUB' | 'GPOS'): LayoutTable | undefined {
    const info = this.parsed.tables.get(kind);
    if (!info) return undefined;
    try {
      const b = this.raw;
      const base = info.offset;
      const scriptList = base + u16(b, base + 4);
      const featureList = base + u16(b, base + 6);
      const lookupList = base + u16(b, base + 8);

      const langSys = (off: number): LangSys => {
        const n = u16(b, off + 4);
        const features: Array<number> = [];
        for (let i = 0; i < n; i++) features.push(u16(b, off + 6 + 2 * i));
        return { required: u16(b, off + 2), features };
      };
      const scripts = new Map<string, Script>();
      for (let i = 0, n = u16(b, scriptList); i < n; i++) {
        const r = scriptList + 2 + 6 * i;
        const s = scriptList + u16(b, r + 4);
        const defOff = u16(b, s);
        const langs = new Map<string, LangSys>();
        for (let k = 0, m = u16(b, s + 2); k < m; k++) {
          const lr = s + 4 + 6 * k;
          langs.set(tag(b, lr), langSys(s + u16(b, lr + 4)));
        }
        scripts.set(tag(b, r), {
          ...(defOff ? { defaultLangSys: langSys(s + defOff) } : {}),
          langSys: langs,
        });
      }

      const features: Array<{ tag: string; lookups: Array<number> }> = [];
      for (let i = 0, n = u16(b, featureList); i < n; i++) {
        const r = featureList + 2 + 6 * i;
        const f = featureList + u16(b, r + 4);
        const lookups: Array<number> = [];
        for (let k = 0, m = u16(b, f + 2); k < m; k++) lookups.push(u16(b, f + 4 + 2 * k));
        features.push({ tag: tag(b, r), lookups });
      }

      const extension = kind === 'GSUB' ? 7 : 9;
      const lookups: Array<Lookup> = [];
      for (let i = 0, n = u16(b, lookupList); i < n; i++) {
        const l = lookupList + u16(b, lookupList + 2 + 2 * i);
        const declared = u16(b, l);
        let type = declared;
        const flag = u16(b, l + 2);
        const count = u16(b, l + 4);
        const subtables: Array<number> = [];
        for (let k = 0; k < count; k++) {
          let st = l + u16(b, l + 6 + 2 * k);
          if (declared === extension) {
            // An Extension subtable: a format, the type it stands for, and a
            // 32-bit offset to the subtable it stands for.
            type = u16(b, st + 2);
            st += u32(b, st + 4);
          }
          subtables.push(st);
        }
        const markSet = flag & USE_MARK_FILTERING_SET ? u16(b, l + 6 + 2 * count) : 0;
        lookups.push({ type, flag, markSet, subtables });
      }
      return { kind, scripts, features, lookups };
    } catch {
      return undefined;
    }
  }
}

/**
 * The script a table answers for, of the tags a shaper prefers: the first the
 * table has, else its default script, else Latin.
 *
 * @param table      The GSUB or GPOS table.
 * @param scriptTags The tags, most preferred first (`dev2`, then `deva`).
 * @returns The tag found, or `undefined` when the table has none of them.
 */
export function pickScript(
  table: LayoutTable | undefined,
  scriptTags: ReadonlyArray<string>,
): string | undefined {
  if (!table) return undefined;
  for (const t of [...scriptTags, 'DFLT', 'dflt', 'latn']) if (table.scripts.has(t)) return t;
  return undefined;
}

/**
 * The lookups a set of features runs, merged and in LookupList order — the
 * order they apply in, whichever feature named them.
 *
 * A lookup two of the features name runs once, on the glyphs either feature
 * may act on. The language system's required feature, if it has one, runs on
 * every glyph.
 *
 * @param table      The GSUB or GPOS table.
 * @param scriptTags The script tags to look for, most preferred first.
 * @param features   The features, with the glyphs each acts on.
 * @param language   A language system tag; the script's default otherwise.
 * @returns The steps to run, in order.
 */
export function planLookups(
  table: LayoutTable | undefined,
  scriptTags: ReadonlyArray<string>,
  features: ReadonlyArray<OtlFeature>,
  language?: string,
): Array<OtlStep> {
  if (!table) return [];
  const scriptTag = pickScript(table, scriptTags);
  const script = scriptTag === undefined ? undefined : table.scripts.get(scriptTag);
  const lang =
    (language !== undefined ? script?.langSys.get(language) : undefined) ?? script?.defaultLangSys;
  if (!lang) return [];
  const steps = new Map<number, OtlStep>();
  const add = (lookup: number, f: OtlFeature): void => {
    if (lookup >= table.lookups.length) return;
    const had = steps.get(lookup);
    steps.set(lookup, {
      lookup,
      mask: (had?.mask ?? 0) | f.mask,
      autoZwj: (had?.autoZwj ?? true) && f.manualZwj !== true,
      autoZwnj: (had?.autoZwnj ?? true) && f.manualZwnj !== true,
      perSyllable: (had?.perSyllable ?? true) && f.perSyllable === true,
    });
  };
  const required = table.features[lang.required];
  if (lang.required !== 0xffff && required) {
    for (const l of required.lookups) add(l, { tag: required.tag, mask: ~0 });
  }
  for (const f of features) {
    for (const index of lang.features) {
      const feature = table.features[index];
      if (feature?.tag !== f.tag) continue;
      for (const l of feature.lookups) add(l, f);
    }
  }
  return [...steps.values()].sort((a, b) => a.lookup - b.lookup);
}

/** The state one lookup runs with. */
interface Run<TGlyph extends OtlGlyph> {
  readonly face: OtlFace;
  readonly table: LayoutTable;
  readonly buf: Array<TGlyph>;
  readonly pos: Array<OtlPosition> | undefined;
  readonly step: OtlStep;
  lookup: Lookup;
  /** The glyph the lookup is applied at. */
  idx: number;
  depth: number;
  ops: number;
  readonly maxLength: number;
}

/**
 * Whether a lookup sees a glyph at all, by its flags: the classes it ignores,
 * and of the marks the one class, or the one set, it keeps.
 */
function visible(face: OtlFace, g: OtlGlyph, flag: number, markSet: number): boolean {
  const cls = g.cls;
  if (cls >= 1 && cls <= 3 && flag & IGNORE_FLAGS & (1 << cls)) return false;
  if (cls !== MARK) return true;
  if (flag & USE_MARK_FILTERING_SET) return face.inMarkSet(markSet, g.gid);
  if (flag & MARK_ATTACHMENT_TYPE) return face.attachClass(g.gid) === flag >> 8;
  return true;
}

/**
 * What a matcher wants of the glyph it lands on. `undefined` takes any glyph,
 * and only as a "maybe": a join control it could step over is stepped over,
 * where a matcher that names its glyph takes it if it is the one named.
 */
type Want = ((g: OtlGlyph, nth: number) => boolean) | undefined;

/**
 * Step from `from` towards `dir`, `count` times, onto the glyphs a match takes:
 * HarfBuzz's skipping iterator. A glyph the lookup's flags hide is stepped
 * over. A join control is stepped over unless it is what the rule asks for,
 * and only where the lookup leaves it to the engine: a ZWJ in a context
 * always, in the input when the feature does; a ZWNJ in GPOS, and in a GSUB
 * context when the feature does. In the input, every glyph must carry the
 * step's mask, and with a per-syllable step every glyph keeps to the syllable
 * of the glyph the lookup is applied at.
 *
 * @returns The positions it landed on, or `undefined` where the match fails.
 */
function walk<TGlyph extends OtlGlyph>(
  r: Run<TGlyph>,
  from: number,
  dir: 1 | -1,
  count: number,
  want: Want,
  context: boolean,
  flag = r.lookup.flag,
): Array<number> | undefined {
  const out: Array<number> = [];
  const gpos = r.table.kind === 'GPOS';
  const ignoreZwnj = gpos || (context && r.step.autoZwnj);
  const ignoreZwj = context || r.step.autoZwj;
  const syllable =
    r.step.perSyllable && !gpos && from === r.idx ? (r.buf[r.idx]?.syllable ?? 0) : 0;
  let i = from;
  while (out.length < count) {
    i += dir;
    const g = r.buf[i];
    if (!g) return undefined;
    if (!visible(r.face, g, flag, r.lookup.markSet)) continue;
    const maybeSkip =
      g.joiner !== undefined &&
      (ignoreZwnj || g.joiner !== 'zwnj') &&
      (ignoreZwj || g.joiner !== 'zwj');
    const eligible =
      (context || (g.mask & r.step.mask) !== 0) && (syllable === 0 || g.syllable === syllable);
    if (eligible && (want ? want(g, out.length) : !maybeSkip)) {
      out.push(i);
      continue;
    }
    if (!maybeSkip) return undefined;
  }
  return out;
}

/**
 * Run a stage of GSUB lookups over the buffer, each over the whole of it
 * before the next begins.
 *
 * @param face  The face.
 * @param steps The lookups, from {@link planLookups} over its GSUB.
 * @param buf   The glyphs; substituted in place, and may grow or shrink.
 */
export function substitute<TGlyph extends OtlGlyph>(
  face: OtlFace,
  steps: ReadonlyArray<OtlStep>,
  buf: Array<TGlyph>,
): void {
  const table = face.gsub;
  if (!table) return;
  const maxLength = Math.max(64, buf.length * MAX_GROWTH);
  for (const step of steps) {
    const lookup = table.lookups[step.lookup];
    if (!lookup) continue;
    const starts = firstGlyphs(face, table, step.lookup);
    const r: Run<TGlyph> = {
      face,
      table,
      buf,
      pos: undefined,
      step,
      lookup,
      idx: 0,
      depth: 0,
      ops: 0,
      maxLength,
    };
    // GSUB LookupType 8 is applied from the end of the buffer to its start.
    if (lookup.type === 8) {
      for (let i = buf.length - 1; i >= 0; i--) {
        const g = buf[i]!;
        if (!(g.mask & step.mask) || !starts.has(g.gid)) continue;
        if (!visible(face, g, lookup.flag, lookup.markSet)) continue;
        r.idx = i;
        applyLookup(r);
      }
      continue;
    }
    let i = 0;
    while (i < buf.length) {
      const g = buf[i]!;
      if (
        g.mask & step.mask &&
        starts.has(g.gid) &&
        visible(face, g, lookup.flag, lookup.markSet)
      ) {
        r.idx = i;
        r.lookup = lookup;
        const next = applyLookup(r);
        if (next >= 0) {
          i = next;
          if (++r.ops > MAX_OPS_PER_GLYPH * Math.max(1, buf.length)) break;
          continue;
        }
      }
      i++;
    }
  }
}

/**
 * Position the buffer: each glyph's own advance, then a stage of GPOS lookups,
 * then every mark and cursive join carried along with the glyph it hangs on.
 *
 * @param face  The face.
 * @param steps The lookups, from {@link planLookups} over its GPOS.
 * @param buf   The glyphs, after substitution.
 * @returns A position per glyph, in font units, y up.
 */
export function position<TGlyph extends OtlGlyph>(
  face: OtlFace,
  steps: ReadonlyArray<OtlStep>,
  buf: Array<TGlyph>,
): Array<OtlPosition> {
  const advances = face.parsed.advanceWidths;
  const pos: Array<OtlPosition> = buf.map((g) => ({
    xAdvance: advances[g.gid] ?? 0,
    yAdvance: 0,
    xOffset: 0,
    yOffset: 0,
    attachChain: 0,
    attachType: 0,
  }));
  const table = face.gpos;
  if (table) {
    for (const step of steps) {
      const lookup = table.lookups[step.lookup];
      if (!lookup) continue;
      const starts = firstGlyphs(face, table, step.lookup);
      const r: Run<TGlyph> = {
        face,
        table,
        buf,
        pos,
        step,
        lookup,
        idx: 0,
        depth: 0,
        ops: 0,
        maxLength: buf.length,
      };
      let i = 0;
      while (i < buf.length) {
        const g = buf[i]!;
        if (
          g.mask & step.mask &&
          starts.has(g.gid) &&
          visible(face, g, lookup.flag, lookup.markSet)
        ) {
          r.idx = i;
          r.lookup = lookup;
          const next = applyLookup(r);
          if (next >= 0) {
            i = Math.max(next, i + 1);
            continue;
          }
        }
        i++;
      }
    }
  }
  for (let i = 0; i < pos.length; i++) settle(pos, i, MAX_NESTING * 4);
  return pos;
}

const firstGlyphCache = new WeakMap<LayoutTable, Map<number, ReadonlySet<number>>>();

/**
 * Every glyph a lookup can start at: the union of its subtables' first
 * coverage. A glyph outside it is passed over without reading a subtable, as
 * HarfBuzz's digests pass it over; most of a few hundred lookups touch none of
 * a word's glyphs.
 */
function firstGlyphs(face: OtlFace, table: LayoutTable, index: number): ReadonlySet<number> {
  let byLookup = firstGlyphCache.get(table);
  if (!byLookup) {
    byLookup = new Map();
    firstGlyphCache.set(table, byLookup);
  }
  const had = byLookup.get(index);
  if (had) return had;
  const b = face.bytes;
  const lookup = table.lookups[index]!;
  const out = new Set<number>();
  const contextual =
    table.kind === 'GSUB' ? lookup.type === 5 || lookup.type === 6 : lookup.type >= 7;
  for (const st of lookup.subtables) {
    let cov = st + u16(b, st + 2);
    if (contextual && u16(b, st) === 3) {
      if (lookup.type === 6 || lookup.type === 8) {
        const back = u16(b, st + 2);
        cov = st + u16(b, st + 6 + 2 * back);
      } else {
        cov = st + u16(b, st + 6);
      }
    }
    for (const g of face.coverage(cov).keys()) out.add(g);
  }
  byLookup.set(index, out);
  return out;
}

/**
 * Whether a feature's lookups would substitute this exact sequence of glyphs
 * — the question a shaper asks before it decides that a consonant takes its
 * below-base form, or that a Ra starts a reph.
 *
 * @param face        The face.
 * @param steps       The feature's lookups, from {@link planLookups}.
 * @param gids        The glyphs, in order.
 * @param zeroContext Only a rule with no backtrack and no lookahead counts.
 * @returns `true` when one of them would.
 */
export function wouldSubstitute(
  face: OtlFace,
  steps: ReadonlyArray<OtlStep>,
  gids: ReadonlyArray<number>,
  zeroContext: boolean,
): boolean {
  const table = face.gsub;
  if (!table || gids.length === 0) return false;
  const b = face.bytes;
  const n = gids.length;
  const first = gids[0]!;
  for (const step of steps) {
    const lookup = table.lookups[step.lookup];
    if (!lookup) continue;
    for (const st of lookup.subtables) {
      const format = u16(b, st);
      switch (lookup.type) {
        case 1:
        case 2:
        case 3:
          if (n === 1 && face.coverage(st + u16(b, st + 2)).has(first)) return true;
          break;
        case 4: {
          const index = face.coverage(st + u16(b, st + 2)).get(first);
          if (index === undefined) break;
          const set = st + u16(b, st + 6 + 2 * index);
          for (let k = 0, m = u16(b, set); k < m; k++) {
            const lig = set + u16(b, set + 2 + 2 * k);
            if (u16(b, lig + 2) !== n) continue;
            let ok = true;
            for (let c = 1; c < n && ok; c++) ok = u16(b, lig + 2 + 2 * c) === gids[c];
            if (ok) return true;
          }
          break;
        }
        case 5:
        case 6: {
          const rules = contextRules(face, lookup.type, format, st, first);
          for (const rule of rules) {
            if (rule.input.length + 1 !== n) continue;
            if (zeroContext && (rule.backtrack.length > 0 || rule.lookahead.length > 0)) continue;
            let ok = true;
            for (let c = 1; c < n && ok; c++) ok = rule.input[c - 1]!(gids[c]!);
            if (ok) return true;
          }
          break;
        }
        default:
          break;
      }
    }
  }
  return false;
}

/** Try the lookup's subtables at `r.idx`; the index to go on from, or -1. */
function applyLookup<TGlyph extends OtlGlyph>(r: Run<TGlyph>): number {
  for (const st of r.lookup.subtables) {
    const next = r.table.kind === 'GSUB' ? applySubstitution(r, st) : applyPositioning(r, st);
    if (next >= 0) return next;
  }
  return -1;
}

function applySubstitution<TGlyph extends OtlGlyph>(r: Run<TGlyph>, st: number): number {
  const b = r.face.bytes;
  const g = r.buf[r.idx]!;
  const format = u16(b, st);
  switch (r.lookup.type) {
    case 1: {
      if (format !== 1 && format !== 2) return -1;
      const index = r.face.coverage(st + u16(b, st + 2)).get(g.gid);
      if (index === undefined) return -1;
      const gid = format === 1 ? (g.gid + i16(b, st + 4)) & 0xffff : u16(b, st + 6 + 2 * index);
      replace(r, r.idx, gid);
      return r.idx + 1;
    }
    case 2: {
      if (format !== 1) return -1;
      const index = r.face.coverage(st + u16(b, st + 2)).get(g.gid);
      if (index === undefined) return -1;
      const seq = st + u16(b, st + 6 + 2 * index);
      const count = u16(b, seq);
      if (r.buf.length + count - 1 > r.maxLength) return -1;
      // One glyph is a plain substitution, as HarfBuzz counts it.
      if (count === 1) {
        replace(r, r.idx, u16(b, seq + 2));
        return r.idx + 1;
      }
      const made: Array<TGlyph> = [];
      const props = (g.props ?? 0) | SUBSTITUTED | MULTIPLIED;
      for (let k = 0; k < count; k++) {
        const gid = u16(b, seq + 2 + 2 * k);
        made.push({ ...g, gid, cls: r.face.glyphClass(gid), props });
      }
      r.buf.splice(r.idx, 1, ...made);
      return r.idx + count;
    }
    case 3: {
      // An alternate is picked by a feature's value; on, a feature picks the first.
      if (format !== 1) return -1;
      const index = r.face.coverage(st + u16(b, st + 2)).get(g.gid);
      if (index === undefined) return -1;
      const set = st + u16(b, st + 6 + 2 * index);
      if (u16(b, set) === 0) return -1;
      replace(r, r.idx, u16(b, set + 2));
      return r.idx + 1;
    }
    case 4:
      return format === 1 ? ligate(r, st) : -1;
    case 5:
    case 6:
      return applyContext(r, st);
    case 8:
      return format === 1 ? reverseChain(r, st) : -1;
    default:
      return -1;
  }
}

/** Put another glyph in a slot, with the class GDEF gives the new one. */
function replace<TGlyph extends OtlGlyph>(r: Run<TGlyph>, at: number, gid: number): void {
  const g = r.buf[at]!;
  g.gid = gid;
  g.cls = r.face.glyphClass(gid);
  g.props = (g.props ?? 0) | SUBSTITUTED;
}

let nextLigId = 1;

/**
 * Make `buf[start..end)` one cluster, the least of theirs. A cluster cut at
 * either edge comes in whole, as HarfBuzz merges them: a vowel sign that was
 * one cluster with the last letter of a conjunct is one with the conjunct.
 *
 * @param buf   The glyphs.
 * @param start The first glyph of the stretch.
 * @param end   The glyph after its last.
 */
export function mergeClusters(buf: Array<OtlGlyph>, start: number, end: number): void {
  if (end - start < 2) return;
  let cluster = buf[start]!.cluster;
  for (let i = start + 1; i < end; i++) cluster = Math.min(cluster, buf[i]!.cluster);
  if (cluster !== buf[end - 1]!.cluster) {
    while (end < buf.length && buf[end - 1]!.cluster === buf[end]!.cluster) end++;
  }
  if (cluster !== buf[start]!.cluster) {
    while (start > 0 && buf[start - 1]!.cluster === buf[start]!.cluster) start--;
  }
  for (let i = start; i < end; i++) buf[i]!.cluster = cluster;
}

/** GSUB LookupType 4: the first ligature of the set whose components follow. */
function ligate<TGlyph extends OtlGlyph>(r: Run<TGlyph>, st: number): number {
  const b = r.face.bytes;
  const index = r.face.coverage(st + u16(b, st + 2)).get(r.buf[r.idx]!.gid);
  if (index === undefined) return -1;
  const set = st + u16(b, st + 6 + 2 * index);
  for (let k = 0, m = u16(b, set); k < m; k++) {
    const lig = set + u16(b, set + 2 + 2 * k);
    const count = u16(b, lig + 2);
    const at =
      count <= 1
        ? []
        : walk(r, r.idx, 1, count - 1, (g, nth) => g.gid === u16(b, lig + 4 + 2 * nth), false);
    if (!at) continue;
    const last = at.length > 0 ? at[at.length - 1]! : r.idx;
    // The whole stretch is one cluster now, the marks the match stepped over
    // with it, and each of those marks remembers which component it followed.
    mergeClusters(r.buf, r.idx, last + 1);
    const ligId = nextLigId++;
    let comp = 1;
    for (let i = r.idx + 1; i <= last; i++) {
      const g = r.buf[i]!;
      if (at.includes(i)) comp++;
      else {
        g.ligId = ligId;
        g.ligComp = comp;
      }
    }
    const first = r.buf[r.idx]!;
    first.ligId = ligId;
    first.ligComp = 0;
    replace(r, r.idx, u16(b, lig));
    // A ligature forgives an earlier multiplication: it is what the glyph is now.
    first.props = ((first.props ?? 0) | LIGATED) & ~MULTIPLIED;
    for (let c = at.length - 1; c >= 0; c--) r.buf.splice(at[c]!, 1);
    return last + 1 - at.length;
  }
  return -1;
}

/** GSUB LookupType 8: one substitution, by context, from the end backwards. */
function reverseChain<TGlyph extends OtlGlyph>(r: Run<TGlyph>, st: number): number {
  const b = r.face.bytes;
  const index = r.face.coverage(st + u16(b, st + 2)).get(r.buf[r.idx]!.gid);
  if (index === undefined) return -1;
  let o = st + 4;
  const back = u16(b, o);
  const backCov = Array.from({ length: back }, (_, k) => st + u16(b, o + 2 + 2 * k));
  o += 2 + 2 * back;
  const ahead = u16(b, o);
  const aheadCov = Array.from({ length: ahead }, (_, k) => st + u16(b, o + 2 + 2 * k));
  o += 2 + 2 * ahead;
  const covers =
    (offs: Array<number>): Want =>
    (g, nth) =>
      r.face.coverage(offs[nth]!).has(g.gid);
  if (!walk(r, r.idx, -1, back, covers(backCov), true)) return -1;
  if (!walk(r, r.idx, 1, ahead, covers(aheadCov), true)) return -1;
  replace(r, r.idx, u16(b, o + 2 + 2 * index));
  return r.idx + 1;
}

/** A context rule, its glyphs read as tests whatever format named them. */
interface ContextRule {
  readonly backtrack: ReadonlyArray<(gid: number) => boolean>;
  /** The input past its first glyph. */
  readonly input: ReadonlyArray<(gid: number) => boolean>;
  readonly lookahead: ReadonlyArray<(gid: number) => boolean>;
  /** The SequenceLookupRecords: where in the input, and which lookup. */
  readonly records: ReadonlyArray<readonly [number, number]>;
}

/**
 * The rules of a context subtable (LookupType 5/6 of GSUB, 7/8 of GPOS) that
 * may start on `first`, in the order they are tried.
 */
function contextRules(
  face: OtlFace,
  type: number,
  format: number,
  st: number,
  first: number,
): ReadonlyArray<ContextRule> {
  // Read once per subtable and first glyph: a rule is tried at every glyph a
  // lookup passes over, and a face of India runs a few hundred of them.
  let bySubtable = ruleCache.get(face);
  if (!bySubtable) {
    bySubtable = new Map();
    ruleCache.set(face, bySubtable);
  }
  let byGlyph = bySubtable.get(st);
  if (!byGlyph) {
    byGlyph = new Map();
    bySubtable.set(st, byGlyph);
  }
  let rules = byGlyph.get(first);
  if (!rules) {
    rules = readContextRules(face, type, format, st, first);
    byGlyph.set(first, rules);
  }
  return rules;
}

const ruleCache = new WeakMap<OtlFace, Map<number, Map<number, ReadonlyArray<ContextRule>>>>();

function readContextRules(
  face: OtlFace,
  type: number,
  format: number,
  st: number,
  first: number,
): Array<ContextRule> {
  const b = face.bytes;
  const chained = type === 6 || type === 8;
  const records = (o: number, n: number): Array<readonly [number, number]> =>
    Array.from({ length: n }, (_, k) => [u16(b, o + 4 * k), u16(b, o + 4 * k + 2)] as const);
  const glyphIs = (gid: number) => (g: number) => g === gid;
  const out: Array<ContextRule> = [];
  if (format === 1 || format === 2) {
    let index: number | undefined;
    let ruleSets: number;
    let classes: Array<ReadonlyMap<number, number>> = [];
    if (format === 1) {
      index = face.coverage(st + u16(b, st + 2)).get(first);
      ruleSets = st + 6;
    } else if (!chained) {
      if (!face.coverage(st + u16(b, st + 2)).has(first)) return out;
      const input = face.classDef(st + u16(b, st + 4));
      classes = [input, input, input];
      index = input.get(first) ?? 0;
      ruleSets = st + 8;
    } else {
      if (!face.coverage(st + u16(b, st + 2)).has(first)) return out;
      const def = (o: number): ReadonlyMap<number, number> =>
        u16(b, st + o) ? face.classDef(st + u16(b, st + o)) : new Map();
      classes = [def(4), def(6), def(8)];
      index = classes[1]!.get(first) ?? 0;
      ruleSets = st + 12;
    }
    if (index === undefined) return out;
    const setCount = u16(b, ruleSets - 2);
    if (index >= setCount) return out;
    const setOff = u16(b, ruleSets + 2 * index);
    if (!setOff) return out;
    const set = st + setOff;
    const test =
      format === 1
        ? (_which: number) => glyphIs
        : (which: number) => (cls: number) => (g: number) => (classes[which]!.get(g) ?? 0) === cls;
    for (let k = 0, m = u16(b, set); k < m; k++) {
      let o = set + u16(b, set + 2 + 2 * k);
      if (!chained) {
        const glyphs = u16(b, o);
        const count = u16(b, o + 2);
        const input = Array.from({ length: Math.max(0, glyphs - 1) }, (_, c) =>
          test(1)(u16(b, o + 4 + 2 * c)),
        );
        out.push({
          backtrack: [],
          input,
          lookahead: [],
          records: records(o + 4 + 2 * input.length, count),
        });
        continue;
      }
      const back = u16(b, o);
      const backtrack = Array.from({ length: back }, (_, c) => test(0)(u16(b, o + 2 + 2 * c)));
      o += 2 + 2 * back;
      const glyphs = u16(b, o);
      const input = Array.from({ length: Math.max(0, glyphs - 1) }, (_, c) =>
        test(1)(u16(b, o + 2 + 2 * c)),
      );
      o += 2 + 2 * input.length;
      const ahead = u16(b, o);
      const lookahead = Array.from({ length: ahead }, (_, c) => test(2)(u16(b, o + 2 + 2 * c)));
      o += 2 + 2 * ahead;
      out.push({ backtrack, input, lookahead, records: records(o + 2, u16(b, o)) });
    }
    return out;
  }
  if (format !== 3) return out;
  const covers = (off: number) => (g: number) => face.coverage(off).has(g);
  if (!chained) {
    const glyphs = u16(b, st + 2);
    const count = u16(b, st + 4);
    const covs = Array.from({ length: glyphs }, (_, c) => st + u16(b, st + 6 + 2 * c));
    if (covs.length === 0 || !face.coverage(covs[0]!).has(first)) return out;
    out.push({
      backtrack: [],
      input: covs.slice(1).map(covers),
      lookahead: [],
      records: records(st + 6 + 2 * glyphs, count),
    });
    return out;
  }
  let o = st + 2;
  const back = u16(b, o);
  const backtrack = Array.from({ length: back }, (_, c) => covers(st + u16(b, o + 2 + 2 * c)));
  o += 2 + 2 * back;
  const glyphs = u16(b, o);
  const covs = Array.from({ length: glyphs }, (_, c) => st + u16(b, o + 2 + 2 * c));
  o += 2 + 2 * glyphs;
  const ahead = u16(b, o);
  const lookahead = Array.from({ length: ahead }, (_, c) => covers(st + u16(b, o + 2 + 2 * c)));
  o += 2 + 2 * ahead;
  if (covs.length === 0 || !face.coverage(covs[0]!).has(first)) return out;
  out.push({
    backtrack,
    input: covs.slice(1).map(covers),
    lookahead,
    records: records(o + 2, u16(b, o)),
  });
  return out;
}

/**
 * A context lookup (GSUB 5/6, GPOS 7/8): the first rule whose glyphs are all
 * there runs its lookups at the places it names, and the run goes on after
 * the matched input.
 */
function applyContext<TGlyph extends OtlGlyph>(r: Run<TGlyph>, st: number): number {
  const b = r.face.bytes;
  const rules = contextRules(r.face, r.lookup.type, u16(b, st), st, r.buf[r.idx]!.gid);
  for (const rule of rules) {
    const input = walk(r, r.idx, 1, rule.input.length, (g, nth) => rule.input[nth]!(g.gid), false);
    if (!input) continue;
    const end = input.length > 0 ? input[input.length - 1]! : r.idx;
    if (!walk(r, r.idx, -1, rule.backtrack.length, (g, nth) => rule.backtrack[nth]!(g.gid), true)) {
      continue;
    }
    if (!walk(r, end, 1, rule.lookahead.length, (g, nth) => rule.lookahead[nth]!(g.gid), true)) {
      continue;
    }
    return runRecords(r, [r.idx, ...input], rule.records);
  }
  return -1;
}

/**
 * Run a rule's SequenceLookupRecords at the places of its input, and return
 * the index after the input. A lookup that grows or shrinks the buffer moves
 * the places after it, by HarfBuzz's reckoning: glyphs it added came right
 * after the place it ran at, and glyphs it removed were the places after it.
 */
function runRecords<TGlyph extends OtlGlyph>(
  r: Run<TGlyph>,
  places: Array<number>,
  records: ReadonlyArray<readonly [number, number]>,
): number {
  const match = [...places];
  let end = match[match.length - 1]! + 1;
  if (r.depth >= MAX_NESTING) return end;
  const outer = { lookup: r.lookup, idx: r.idx };
  for (const [seqIndex, lookupIndex] of records) {
    if (seqIndex >= match.length) continue;
    const lookup = r.table.lookups[lookupIndex];
    if (!lookup) continue;
    const at = match[seqIndex]!;
    if (at >= r.buf.length) continue;
    const before = r.buf.length;
    r.lookup = lookup;
    r.idx = at;
    r.depth++;
    const applied = applyLookup(r) >= 0;
    r.depth--;
    r.lookup = outer.lookup;
    r.idx = outer.idx;
    if (!applied) continue;
    let delta = r.buf.length - before;
    if (delta === 0) continue;
    end += delta;
    if (end < at) {
      delta += at - end;
      end = at;
    }
    let next = seqIndex + 1;
    if (delta < 0) {
      delta = Math.max(delta, next - match.length);
      next -= delta;
    }
    const tail = match.slice(next);
    match.length = next + delta;
    for (let j = seqIndex + 1; j < next + delta; j++) match[j] = match[j - 1]! + 1;
    for (const t of tail) match.push(t + delta);
  }
  return end;
}

/** A ValueRecord's size in bytes, by its format. */
function valueSize(format: number): number {
  let n = 0;
  for (let f = format & 0xff; f; f >>= 1) n += f & 1;
  return 2 * n;
}

/** Add a ValueRecord to a glyph's position; device tables are not read. */
function addValue(b: Uint8Array, o: number, format: number, p: OtlPosition): void {
  if (format & 0x1) {
    p.xOffset += i16(b, o);
    o += 2;
  }
  if (format & 0x2) {
    p.yOffset += i16(b, o);
    o += 2;
  }
  if (format & 0x4) {
    p.xAdvance += i16(b, o);
    o += 2;
  }
  // A yAdvance moves the pen only in vertical text.
}

/** An anchor's point; formats 2 and 3 add a contour point or a device table, not read. */
function anchor(b: Uint8Array, o: number): readonly [number, number] {
  return [i16(b, o + 2), i16(b, o + 4)];
}

function applyPositioning<TGlyph extends OtlGlyph>(r: Run<TGlyph>, st: number): number {
  const b = r.face.bytes;
  const pos = r.pos!;
  const g = r.buf[r.idx]!;
  const format = u16(b, st);
  switch (r.lookup.type) {
    case 1: {
      const index = r.face.coverage(st + u16(b, st + 2)).get(g.gid);
      if (index === undefined) return -1;
      const vf = u16(b, st + 4);
      if (format === 1) addValue(b, st + 6, vf, pos[r.idx]!);
      else if (format === 2) addValue(b, st + 8 + index * valueSize(vf), vf, pos[r.idx]!);
      else return -1;
      return r.idx + 1;
    }
    case 2:
      return pairPos(r, st, format);
    case 3:
      return format === 1 ? cursivePos(r, st) : -1;
    case 4:
    case 5:
    case 6:
      return format === 1 ? markPos(r, st) : -1;
    case 7:
    case 8:
      return applyContext(r, st);
    default:
      return -1;
  }
}

/** GPOS LookupType 2: the current glyph and the next one the lookup sees. */
function pairPos<TGlyph extends OtlGlyph>(r: Run<TGlyph>, st: number, format: number): number {
  const b = r.face.bytes;
  const pos = r.pos!;
  const first = r.buf[r.idx]!;
  const index = r.face.coverage(st + u16(b, st + 2)).get(first.gid);
  if (index === undefined) return -1;
  const next = walk(r, r.idx, 1, 1, undefined, false);
  if (!next) return -1;
  const j = next[0]!;
  const second = r.buf[j]!.gid;
  const vf1 = u16(b, st + 4);
  const vf2 = u16(b, st + 6);
  const s1 = valueSize(vf1);
  const s2 = valueSize(vf2);
  let record: number | undefined;
  if (format === 1) {
    const set = st + u16(b, st + 10 + 2 * index);
    const size = 2 + s1 + s2;
    // The second glyphs are sorted by id: search them.
    let lo = 0;
    let hi = u16(b, set) - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const at = set + 2 + mid * size;
      const gid = u16(b, at);
      if (gid === second) {
        record = at + 2;
        break;
      }
      if (gid < second) lo = mid + 1;
      else hi = mid - 1;
    }
  } else if (format === 2) {
    const c1 = r.face.classDef(st + u16(b, st + 8)).get(first.gid) ?? 0;
    const c2 = r.face.classDef(st + u16(b, st + 10)).get(second) ?? 0;
    const n1 = u16(b, st + 12);
    const n2 = u16(b, st + 14);
    if (c1 >= n1 || c2 >= n2) return -1;
    record = st + 16 + (c1 * n2 + c2) * (s1 + s2);
  } else return -1;
  if (record === undefined) return -1;
  addValue(b, record, vf1, pos[r.idx]!);
  addValue(b, record + s1, vf2, pos[j]!);
  return vf2 ? j + 1 : j;
}

/** GPOS LookupType 3: the current glyph's entry joined to the exit of the one before. */
function cursivePos<TGlyph extends OtlGlyph>(r: Run<TGlyph>, st: number): number {
  const b = r.face.bytes;
  const pos = r.pos!;
  const cov = r.face.coverage(st + u16(b, st + 2));
  const record = (gid: number): number | undefined => {
    const index = cov.get(gid);
    return index === undefined ? undefined : st + 6 + 4 * index;
  };
  const here = record(r.buf[r.idx]!.gid);
  if (here === undefined || !u16(b, here)) return -1;
  const prev = walk(r, r.idx, -1, 1, undefined, false);
  if (!prev) return -1;
  const i = prev[0]!;
  const there = record(r.buf[i]!.gid);
  if (there === undefined || !u16(b, there + 2)) return -1;
  const j = r.idx;
  const [exitX, exitY] = anchor(b, st + u16(b, there + 2));
  const [entryX, entryY] = anchor(b, st + u16(b, here));
  pos[i]!.xAdvance = exitX + pos[i]!.xOffset;
  const d = entryX + pos[j]!.xOffset;
  pos[j]!.xAdvance -= d;
  pos[j]!.xOffset -= d;
  // The child hangs off its parent; right-to-left, the later glyph is the parent.
  let child = i;
  let parent = j;
  let y = entryY - exitY;
  if (!(r.lookup.flag & RIGHT_TO_LEFT)) {
    child = j;
    parent = i;
    y = -y;
  }
  reverseCursive(pos, child, parent);
  pos[child]!.attachType = 2;
  pos[child]!.attachChain = parent - child;
  pos[child]!.yOffset = y;
  if (pos[parent]!.attachChain === -pos[child]!.attachChain) {
    pos[parent]!.attachChain = 0;
    pos[parent]!.yOffset = 0;
  }
  return r.idx + 1;
}

/** Turn a cursive chain round, so the whole of it now hangs off a new parent. */
function reverseCursive(pos: Array<OtlPosition>, i: number, newParent: number): void {
  const p = pos[i]!;
  const chain = p.attachChain;
  if (!chain || p.attachType !== 2) return;
  p.attachChain = 0;
  const j = i + chain;
  const q = pos[j];
  if (j === newParent || !q) return;
  reverseCursive(pos, j, newParent);
  q.yOffset = -p.yOffset;
  q.attachChain = -chain;
  q.attachType = 2;
}

/**
 * GPOS LookupType 4, 5 and 6: a mark on the anchor of the base, the ligature
 * component or the mark it belongs to.
 */
function markPos<TGlyph extends OtlGlyph>(r: Run<TGlyph>, st: number): number {
  const b = r.face.bytes;
  const pos = r.pos!;
  const mark = r.buf[r.idx]!;
  const markIndex = r.face.coverage(st + u16(b, st + 2)).get(mark.gid);
  if (markIndex === undefined) return -1;
  const classCount = u16(b, st + 6);
  const markArray = st + u16(b, st + 8);
  const targetArray = st + u16(b, st + 10);
  const targetCov = r.face.coverage(st + u16(b, st + 4));

  let j = -1;
  if (r.lookup.type === 6) {
    // The mark before it, with the lookup's mark filtering but not its ignores.
    const prev = walk(r, r.idx, -1, 1, undefined, false, r.lookup.flag & ~IGNORE_FLAGS);
    if (!prev) return -1;
    j = prev[0]!;
    if (r.buf[j]!.cls !== MARK) return -1;
  } else {
    // The nearest glyph before it that is not a mark.
    const prev = walk(r, r.idx, -1, 1, undefined, false, IGNORE_MARKS);
    if (!prev) return -1;
    j = prev[0]!;
  }
  const targetIndex = targetCov.get(r.buf[j]!.gid);
  if (targetIndex === undefined) return -1;

  const markRecord = markArray + 2 + 4 * markIndex;
  const markClass = u16(b, markRecord);
  const markAnchor = anchor(b, markArray + u16(b, markRecord + 2));
  if (markClass >= classCount) return -1;
  let anchorOff: number;
  if (r.lookup.type === 5) {
    const attach = targetArray + u16(b, targetArray + 2 + 2 * targetIndex);
    const components = u16(b, attach);
    if (components === 0) return -1;
    const target = r.buf[j]!;
    const comp =
      mark.ligId !== 0 && mark.ligId === target.ligId && mark.ligComp > 0
        ? Math.min(mark.ligComp, components)
        : components;
    anchorOff = u16(b, attach + 2 + 2 * ((comp - 1) * classCount + markClass));
    if (!anchorOff) return -1;
    anchorOff += attach;
  } else {
    anchorOff = u16(b, targetArray + 2 + 2 * (targetIndex * classCount + markClass));
    if (!anchorOff) return -1;
    anchorOff += targetArray;
  }
  const [bx, by] = anchor(b, anchorOff);
  const p = pos[r.idx]!;
  p.xOffset = bx - markAnchor[0];
  p.yOffset = by - markAnchor[1];
  p.attachType = 1;
  p.attachChain = j - r.idx;
  return r.idx + 1;
}

/**
 * Carry a glyph along with the one it hangs on: a cursive join takes its
 * parent's rise, a mark its base's offset less the advances between them.
 */
function settle(pos: Array<OtlPosition>, i: number, depth: number): void {
  const p = pos[i]!;
  const chain = p.attachChain;
  if (!chain) return;
  p.attachChain = 0;
  const j = i + chain;
  const q = pos[j];
  if (!q || depth === 0) return;
  settle(pos, j, depth - 1);
  if (p.attachType === 2) {
    p.yOffset += q.yOffset;
    return;
  }
  p.xOffset += q.xOffset;
  p.yOffset += q.yOffset;
  for (let k = j; k < i; k++) {
    p.xOffset -= pos[k]!.xAdvance;
    p.yOffset -= pos[k]!.yAdvance;
  }
}
