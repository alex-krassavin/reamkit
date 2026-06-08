// Unicode UAX #9 — Unicode Bidirectional Algorithm.
//
// Implements the resolution rules over a sequence of code points:
//   P2/P3   paragraph embedding level
//   X1-X10  explicit levels & directions (embeddings, overrides, isolates)
//   W1-W7   weak type resolution
//   N0-N2   neutral & isolate-formatting resolution
//   I1-I2   implicit levels
//   L1-L2   reordering (whitespace reset + level reversal)
//
// The implementation follows the reference pseudocode including isolate
// handling (LRI/RLI/FSI/PDI). It does not implement the optional
// retain-explicit-formatting variant — explicit format characters are removed
// from the visual output per X9.

import type { BidiClass } from '@/bidi/char-types';
import { bidiClass } from '@/bidi/char-types';

const MAX_DEPTH = 125;

export type Direction = 'ltr' | 'rtl' | 'auto';

export interface BidiResult {
  // One entry per input code point.
  readonly levels: ReadonlyArray<number>;
  // The resolved paragraph embedding level (0 = LTR, 1 = RTL).
  readonly paragraphLevel: number;
  // Working types after all resolution (useful for tests / debugging).
  readonly types: ReadonlyArray<BidiClass>;
}

// BD8/BD9 — find the matching PDI for an isolate initiator at index i, or the
// length of the array if none. Also used to skip isolate spans in P2.
function matchingPDI(types: ReadonlyArray<BidiClass>, openIdx: number): number {
  let depth = 1;
  for (let i = openIdx + 1; i < types.length; i++) {
    const t = types[i]!;
    if (t === 'LRI' || t === 'RLI' || t === 'FSI') depth++;
    else if (t === 'PDI') {
      depth--;
      if (depth === 0) return i;
    }
  }
  return types.length;
}

// P2/P3 — compute a paragraph (or isolate) base level from the first strong
// character, skipping over isolated sub-sequences.
function computeBaseLevel(types: ReadonlyArray<BidiClass>, start: number, end: number): number {
  for (let i = start; i < end; i++) {
    const t = types[i]!;
    if (t === 'LRI' || t === 'RLI' || t === 'FSI') {
      i = matchingPDI(types, i); // skip to matching PDI (loop ++ moves past it)
      continue;
    }
    if (t === 'L') return 0;
    if (t === 'R' || t === 'AL') return 1;
  }
  return 0;
}

interface StatusEntry {
  readonly level: number;
  readonly override: 'neutral' | 'L' | 'R';
  readonly isolate: boolean;
}

export function computeBidi(codePoints: ReadonlyArray<number>, dir: Direction): BidiResult {
  const n = codePoints.length;
  const origTypes: Array<BidiClass> = new Array(n);
  for (let i = 0; i < n; i++) origTypes[i] = bidiClass(codePoints[i]!);

  const paragraphLevel = dir === 'ltr' ? 0 : dir === 'rtl' ? 1 : computeBaseLevel(origTypes, 0, n);

  const types: Array<BidiClass> = origTypes.slice();
  const levels: Array<number> = new Array(n).fill(paragraphLevel);

  // ---- X1-X8: explicit levels and directions ----
  const stack: Array<StatusEntry> = [
    { level: paragraphLevel, override: 'neutral', isolate: false },
  ];
  let overflowIsolate = 0;
  let overflowEmbedding = 0;
  let validIsolate = 0;

  const nextOddLevel = (lvl: number) => (lvl % 2 === 0 ? lvl + 1 : lvl + 2);
  const nextEvenLevel = (lvl: number) => (lvl % 2 === 0 ? lvl + 2 : lvl + 1);
  const top = () => stack[stack.length - 1]!;

  for (let i = 0; i < n; i++) {
    const t = types[i]!;
    switch (t) {
      case 'RLE':
      case 'LRE':
      case 'RLO':
      case 'LRO': {
        const isRTL = t === 'RLE' || t === 'RLO';
        const newLevel = isRTL ? nextOddLevel(top().level) : nextEvenLevel(top().level);
        // Format char takes the embedding level of the entry it follows.
        levels[i] = top().level;
        types[i] = 'BN';
        if (newLevel <= MAX_DEPTH && overflowIsolate === 0 && overflowEmbedding === 0) {
          stack.push({
            level: newLevel,
            override: t === 'RLO' ? 'R' : t === 'LRO' ? 'L' : 'neutral',
            isolate: false,
          });
        } else if (overflowIsolate === 0) {
          overflowEmbedding++;
        }
        break;
      }
      case 'RLI':
      case 'LRI':
      case 'FSI': {
        // The isolate initiator gets the current level/override first.
        const ov = top().override;
        levels[i] = top().level;
        if (ov !== 'neutral') types[i] = ov;
        let isRTL = t === 'RLI';
        if (t === 'FSI') {
          const pdi = matchingPDI(types, i);
          isRTL = computeBaseLevel(types, i + 1, pdi) === 1;
        }
        const newLevel = isRTL ? nextOddLevel(top().level) : nextEvenLevel(top().level);
        if (newLevel <= MAX_DEPTH && overflowIsolate === 0 && overflowEmbedding === 0) {
          validIsolate++;
          stack.push({ level: newLevel, override: 'neutral', isolate: true });
        } else {
          overflowIsolate++;
        }
        break;
      }
      case 'PDI': {
        if (overflowIsolate > 0) {
          overflowIsolate--;
        } else if (validIsolate > 0) {
          overflowEmbedding = 0;
          while (!top().isolate) stack.pop();
          stack.pop();
          validIsolate--;
        }
        const ov = top().override;
        levels[i] = top().level;
        if (ov !== 'neutral') types[i] = ov;
        break;
      }
      case 'PDF': {
        levels[i] = top().level;
        types[i] = 'BN';
        if (overflowIsolate > 0) {
          // ignore
        } else if (overflowEmbedding > 0) {
          overflowEmbedding--;
        } else if (!top().isolate && stack.length >= 2) {
          stack.pop();
        }
        break;
      }
      case 'B': {
        // Paragraph separator — reset to paragraph level (we handle a single
        // paragraph, so this only appears at the very end).
        levels[i] = paragraphLevel;
        break;
      }
      default: {
        const ov = top().override;
        levels[i] = top().level;
        if (ov !== 'neutral') types[i] = ov;
        break;
      }
    }
  }

  // ---- X9: remove explicit format & BN characters from rule processing ----
  // We keep them in place (with their assigned level) but skip them when
  // building isolating run sequences and resolving W/N/I.

  // ---- X10 / BD13: build isolating run sequences ----
  const sequences = buildIsolatingRunSequences(types, levels, paragraphLevel);

  for (const seq of sequences) {
    resolveWeakTypes(types, seq);
    resolveNeutralTypes(types, levels, seq, codePoints);
    resolveImplicitLevels(types, levels, seq);
  }

  // ---- L1: reset whitespace / separators at line/paragraph ends ----
  applyL1(origTypes, types, levels, paragraphLevel);

  return { levels, paragraphLevel, types };
}

interface IsolatingRunSequence {
  // Indices into the code point array, in logical order, excluding removed
  // (BN / explicit) characters.
  readonly indices: Array<number>;
  readonly sos: 'L' | 'R';
  readonly eos: 'L' | 'R';
}

function isRemoved(t: BidiClass): boolean {
  return t === 'BN' || t === 'RLE' || t === 'LRE' || t === 'RLO' || t === 'LRO' || t === 'PDF';
}

// BD7 — a level run is a maximal contiguous substring of characters with the
// same embedding level (ignoring removed chars). BD13 — chain level runs into
// isolating run sequences via isolate initiator / matching PDI links.
function buildIsolatingRunSequences(
  types: ReadonlyArray<BidiClass>,
  levels: ReadonlyArray<number>,
  paragraphLevel: number,
): Array<IsolatingRunSequence> {
  const n = types.length;

  // Build level runs over non-removed characters.
  const runs: Array<Array<number>> = [];
  let current: Array<number> = [];
  let currentLevel = -1;
  for (let i = 0; i < n; i++) {
    if (isRemoved(types[i]!)) continue;
    if (current.length === 0) {
      current = [i];
      currentLevel = levels[i]!;
    } else if (levels[i] === currentLevel) {
      current.push(i);
    } else {
      runs.push(current);
      current = [i];
      currentLevel = levels[i]!;
    }
  }
  if (current.length > 0) runs.push(current);

  // Map each run's first index → run, to chain isolate initiators to PDIs.
  // BD13: a sequence starts at a run whose first char is not a PDI matching an
  // isolate initiator, and continues through runs linked by initiator→PDI.
  const runByFirst = new Map<number, number>();
  for (let r = 0; r < runs.length; r++) runByFirst.set(runs[r]![0]!, r);

  const used = new Array<boolean>(runs.length).fill(false);
  const sequences: Array<IsolatingRunSequence> = [];

  for (let r = 0; r < runs.length; r++) {
    if (used[r]) continue;
    const firstIdx = runs[r]![0]!;
    const firstType = types[firstIdx]!;
    // Skip runs that begin with a PDI continuing an earlier initiator — they
    // are appended to that initiator's sequence below.
    if (firstType === 'PDI' && matchesAnInitiator(types, firstIdx)) continue;

    const seqIndices: Array<number> = [];
    let cur = r;
    for (;;) {
      used[cur] = true;
      seqIndices.push(...runs[cur]!);
      const lastIdx = runs[cur]![runs[cur]!.length - 1]!;
      const lastType = types[lastIdx]!;
      if (lastType === 'LRI' || lastType === 'RLI' || lastType === 'FSI') {
        const pdi = matchingPDI(types, lastIdx);
        if (pdi < n) {
          const nextRun = runByFirst.get(pdi);
          if (nextRun !== undefined && !used[nextRun]) {
            cur = nextRun;
            continue;
          }
        }
      }
      break;
    }

    // sos/eos per X10: compare the sequence's level with the level of the
    // preceding/following non-removed char (or the paragraph level).
    const seqLevel = levels[seqIndices[0]!]!;
    const prevLevel = levelBefore(types, levels, seqIndices[0]!, paragraphLevel);
    const lastSeqIdx = seqIndices[seqIndices.length - 1]!;
    const lastType = types[lastSeqIdx]!;
    let nextLevel: number;
    if (lastType === 'LRI' || lastType === 'RLI' || lastType === 'FSI') {
      // An unmatched isolate initiator: eos uses the paragraph level.
      nextLevel = paragraphLevel;
    } else {
      nextLevel = levelAfter(types, levels, lastSeqIdx, paragraphLevel);
    }
    const sos = Math.max(seqLevel, prevLevel) % 2 === 0 ? 'L' : 'R';
    const eos = Math.max(seqLevel, nextLevel) % 2 === 0 ? 'L' : 'R';
    sequences.push({ indices: seqIndices, sos, eos });
  }

  return sequences;
}

function matchesAnInitiator(types: ReadonlyArray<BidiClass>, pdiIdx: number): boolean {
  // Walk backwards counting isolates to see if this PDI closes one.
  let depth = 1;
  for (let i = pdiIdx - 1; i >= 0; i--) {
    const t = types[i]!;
    if (t === 'PDI') depth++;
    else if (t === 'LRI' || t === 'RLI' || t === 'FSI') {
      depth--;
      if (depth === 0) return true;
    }
  }
  return false;
}

function levelBefore(
  types: ReadonlyArray<BidiClass>,
  levels: ReadonlyArray<number>,
  idx: number,
  paragraphLevel: number,
): number {
  for (let i = idx - 1; i >= 0; i--) {
    if (!isRemoved(types[i]!)) return levels[i]!;
  }
  return paragraphLevel;
}

function levelAfter(
  types: ReadonlyArray<BidiClass>,
  levels: ReadonlyArray<number>,
  idx: number,
  paragraphLevel: number,
): number {
  for (let i = idx + 1; i < types.length; i++) {
    if (!isRemoved(types[i]!)) return levels[i]!;
  }
  return paragraphLevel;
}

// ---- W1-W7 ----
function resolveWeakTypes(types: Array<BidiClass>, seq: IsolatingRunSequence): void {
  const idx = seq.indices;

  // W1: NSM → type of previous char (or sos). Isolate format chars → ON.
  const isIsolateType = (t: BidiClass): boolean =>
    t === 'LRI' || t === 'RLI' || t === 'FSI' || t === 'PDI';
  let prevType: BidiClass = seq.sos;
  for (const i of idx) {
    const t = types[i]!;
    if (t === 'NSM') {
      types[i] = isIsolateType(prevType) ? 'ON' : prevType;
    }
    prevType = types[i]!;
  }

  // W2: EN → AN if the previous strong type is AL.
  let lastStrong: BidiClass = seq.sos;
  for (const i of idx) {
    const t = types[i]!;
    if (t === 'R' || t === 'L' || t === 'AL') lastStrong = t;
    else if (t === 'EN' && lastStrong === 'AL') types[i] = 'AN';
  }

  // W3: AL → R.
  for (const i of idx) {
    if (types[i] === 'AL') types[i] = 'R';
  }

  // W4: single ES between EN→EN becomes EN; single CS between EN→EN or AN→AN
  // becomes that number type.
  for (let k = 1; k < idx.length - 1; k++) {
    const i = idx[k]!;
    const t = types[i]!;
    const prev = types[idx[k - 1]!]!;
    const next = types[idx[k + 1]!]!;
    if (t === 'ES' && prev === 'EN' && next === 'EN') types[i] = 'EN';
    else if (t === 'CS' && prev === 'EN' && next === 'EN') types[i] = 'EN';
    else if (t === 'CS' && prev === 'AN' && next === 'AN') types[i] = 'AN';
  }

  // W5: a run of ET adjacent to EN becomes EN.
  for (let k = 0; k < idx.length; k++) {
    if (types[idx[k]!] !== 'ET') continue;
    let runEnd = k;
    while (runEnd < idx.length && types[idx[runEnd]!] === 'ET') runEnd++;
    const before = k > 0 ? types[idx[k - 1]!]! : seq.sos;
    const after = runEnd < idx.length ? types[idx[runEnd]!]! : seq.eos;
    if (before === 'EN' || after === 'EN') {
      for (let j = k; j < runEnd; j++) types[idx[j]!] = 'EN';
    }
    k = runEnd - 1;
  }

  // W6: remaining ES/ET/CS → ON.
  for (const i of idx) {
    const t = types[i]!;
    if (t === 'ES' || t === 'ET' || t === 'CS') types[i] = 'ON';
  }

  // W7: EN → L if the previous strong type is L.
  lastStrong = seq.sos;
  for (const i of idx) {
    const t = types[i]!;
    if (t === 'R' || t === 'L') lastStrong = t;
    else if (t === 'EN' && lastStrong === 'L') types[i] = 'L';
  }
}

// ---- N0: paired brackets (BD14–BD16) ----
// Canonical Bidi_Paired_Bracket opening→closing pairs (the common ASCII set plus
// the representative Unicode brackets). Enough for ordinary parenthesised text.
const BRACKET_OPEN_TO_CLOSE: ReadonlyMap<number, number> = new Map([
  [0x28, 0x29],
  [0x5b, 0x5d],
  [0x7b, 0x7d],
  [0x2045, 0x2046],
  [0x207d, 0x207e],
  [0x208d, 0x208e],
  [0x2308, 0x2309],
  [0x230a, 0x230b],
  [0x2329, 0x232a],
  [0x27e6, 0x27e7],
  [0x27e8, 0x27e9],
  [0x27ea, 0x27eb],
  [0x27ec, 0x27ed],
  [0x27ee, 0x27ef],
  [0x2983, 0x2984],
  [0x2985, 0x2986],
  [0x2987, 0x2988],
  [0x2989, 0x298a],
  [0x298b, 0x298c],
  [0x298d, 0x298e],
  [0x298f, 0x2990],
  [0x2991, 0x2992],
  [0x2993, 0x2994],
  [0x2995, 0x2996],
  [0x2997, 0x2998],
  [0x3008, 0x3009],
  [0x300a, 0x300b],
  [0x300c, 0x300d],
  [0x300e, 0x300f],
  [0x3010, 0x3011],
  [0x3014, 0x3015],
  [0x3016, 0x3017],
  [0x3018, 0x3019],
  [0x301a, 0x301b],
  [0xff08, 0xff09],
  [0xff3b, 0xff3d],
  [0xff5b, 0xff5d],
  [0xff5f, 0xff60],
  [0xff62, 0xff63],
]);

// BD16 matches with canonical equivalence: U+2329/U+232A ≡ U+3008/U+3009.
function canonBracket(cp: number): number {
  if (cp === 0x2329) return 0x3008;
  if (cp === 0x232a) return 0x3009;
  return cp;
}

const CLOSE_BRACKETS = new Set<number>();
for (const c of BRACKET_OPEN_TO_CLOSE.values()) CLOSE_BRACKETS.add(canonBracket(c));

// Rule N0: resolve the direction of paired brackets so they mirror correctly.
// A pair takes the embedding direction e if a strong e appears between them;
// otherwise the opposite direction o if a strong o appears and the preceding
// context is also o; otherwise it stays neutral for N1/N2.
function resolvePairedBrackets(
  types: Array<BidiClass>,
  levels: ReadonlyArray<number>,
  seq: IsolatingRunSequence,
  codePoints: ReadonlyArray<number>,
): void {
  const idx = seq.indices;
  if (idx.length === 0) return;
  const e: 'L' | 'R' = levels[idx[0]!]! % 2 === 0 ? 'L' : 'R';
  const o: 'L' | 'R' = e === 'L' ? 'R' : 'L';

  // BD16: stack-based pairing (depth capped at 63 per spec).
  const stack: Array<{ close: number; pos: number }> = [];
  const pairs: Array<{ open: number; close: number }> = [];
  for (let p = 0; p < idx.length; p++) {
    if (types[idx[p]!] !== 'ON') continue;
    const cp = canonBracket(codePoints[idx[p]!]!);
    const closing = BRACKET_OPEN_TO_CLOSE.get(cp);
    if (closing !== undefined) {
      if (stack.length === 63) break;
      stack.push({ close: canonBracket(closing), pos: p });
    } else if (CLOSE_BRACKETS.has(cp)) {
      for (let s = stack.length - 1; s >= 0; s--) {
        if (stack[s]!.close === cp) {
          pairs.push({ open: stack[s]!.pos, close: p });
          stack.length = s;
          break;
        }
      }
    }
  }
  pairs.sort((a, b) => a.open - b.open);

  for (const { open, close } of pairs) {
    let foundE = false;
    let foundO = false;
    for (let p = open + 1; p < close; p++) {
      const d = directionOf(types[idx[p]!]!);
      if (d === e) {
        foundE = true;
        break;
      }
      if (d === o) foundO = true;
    }
    let setDir: 'L' | 'R' | null = null;
    if (foundE) {
      setDir = e;
    } else if (foundO) {
      let prev: 'L' | 'R' | 'other' = 'other';
      for (let p = open - 1; p >= 0; p--) {
        const d = directionOf(types[idx[p]!]!);
        if (d !== 'other') {
          prev = d;
          break;
        }
      }
      if (prev === 'other') prev = directionOf(seq.sos);
      setDir = prev === o ? o : e;
    }
    if (setDir) {
      types[idx[open]!] = setDir;
      types[idx[close]!] = setDir;
    }
  }
}

// ---- N0-N2 ----
function resolveNeutralTypes(
  types: Array<BidiClass>,
  levels: ReadonlyArray<number>,
  seq: IsolatingRunSequence,
  codePoints: ReadonlyArray<number>,
): void {
  const idx = seq.indices;
  const isNeutralOrIsolate = (t: BidiClass): boolean =>
    t === 'B' ||
    t === 'S' ||
    t === 'WS' ||
    t === 'ON' ||
    t === 'LRI' ||
    t === 'RLI' ||
    t === 'FSI' ||
    t === 'PDI';

  // N0: paired-bracket resolution runs first (may set brackets to L/R).
  resolvePairedBrackets(types, levels, seq, codePoints);

  // N1: a sequence of neutrals between two characters of the same direction
  // (treating EN/AN as R) takes that direction.
  // N2: remaining neutrals take the embedding direction.
  let k = 0;
  while (k < idx.length) {
    if (!isNeutralOrIsolate(types[idx[k]!]!)) {
      k++;
      continue;
    }
    let runEnd = k;
    while (runEnd < idx.length && isNeutralOrIsolate(types[idx[runEnd]!]!)) runEnd++;

    const leftDir = directionOf(k > 0 ? types[idx[k - 1]!]! : seq.sos);
    const rightDir = directionOf(runEnd < idx.length ? types[idx[runEnd]!]! : seq.eos);

    let resolved: BidiClass;
    if (leftDir === rightDir && (leftDir === 'L' || leftDir === 'R')) {
      resolved = leftDir;
    } else {
      // N2: embedding direction from the run's level.
      resolved = levels[idx[k]!]! % 2 === 0 ? 'L' : 'R';
    }
    for (let j = k; j < runEnd; j++) types[idx[j]!] = resolved;
    k = runEnd;
  }
}

function directionOf(t: BidiClass): 'L' | 'R' | 'other' {
  if (t === 'L') return 'L';
  if (t === 'R' || t === 'EN' || t === 'AN') return 'R';
  return 'other';
}

// ---- I1-I2 ----
function resolveImplicitLevels(
  types: ReadonlyArray<BidiClass>,
  levels: Array<number>,
  seq: IsolatingRunSequence,
): void {
  for (const i of seq.indices) {
    const t = types[i]!;
    const lvl = levels[i]!;
    if (lvl % 2 === 0) {
      // even (L) level
      if (t === 'R') levels[i] = lvl + 1;
      else if (t === 'AN' || t === 'EN') levels[i] = lvl + 2;
    } else {
      // odd (R) level
      if (t === 'L' || t === 'EN' || t === 'AN') levels[i] = lvl + 1;
    }
  }
}

// ---- L1 ----
// On each line, and at paragraph end, reset to the paragraph level: segment
// separators (S), paragraph separators (B), and any sequence of whitespace /
// isolate formatting characters that precedes them or ends the line.
function applyL1(
  origTypes: ReadonlyArray<BidiClass>,
  types: ReadonlyArray<BidiClass>,
  levels: Array<number>,
  paragraphLevel: number,
): void {
  const n = levels.length;
  const isResettableWS = (t: BidiClass): boolean =>
    t === 'WS' || t === 'LRI' || t === 'RLI' || t === 'FSI' || t === 'PDI' || isRemoved(t);

  let i = 0;
  while (i < n) {
    const ot = origTypes[i]!;
    if (ot === 'B' || ot === 'S') {
      levels[i] = paragraphLevel;
      // Reset preceding run of whitespace.
      let j = i - 1;
      while (j >= 0 && isResettableWS(origTypes[j]!)) {
        levels[j] = paragraphLevel;
        j--;
      }
    }
    i++;
  }
  // Trailing whitespace at end of the paragraph.
  let j = n - 1;
  while (j >= 0 && isResettableWS(origTypes[j]!)) {
    levels[j] = paragraphLevel;
    j--;
  }
}

// ---- L2 ----
// Given the resolved levels for a (sub)range, return the visual order as a
// permutation `visual[k] = logical index`. Reverses contiguous runs from the
// highest level down to the lowest odd level.
export function reorderVisual(levels: ReadonlyArray<number>): Array<number> {
  const n = levels.length;
  const order = Array.from({ length: n }, (_, i) => i);
  if (n === 0) return order;

  let highest = 0;
  let lowestOdd = MAX_DEPTH + 1;
  for (const lvl of levels) {
    if (lvl > highest) highest = lvl;
    if (lvl % 2 === 1 && lvl < lowestOdd) lowestOdd = lvl;
  }

  for (let level = highest; level >= lowestOdd; level--) {
    let k = 0;
    while (k < n) {
      if (levels[order[k]!]! < level) {
        k++;
        continue;
      }
      let runEnd = k;
      while (runEnd < n && levels[order[runEnd]!]! >= level) runEnd++;
      // Reverse order[k..runEnd).
      let a = k;
      let b = runEnd - 1;
      while (a < b) {
        const tmp = order[a]!;
        order[a] = order[b]!;
        order[b] = tmp;
        a++;
        b--;
      }
      k = runEnd;
    }
  }
  return order;
}
