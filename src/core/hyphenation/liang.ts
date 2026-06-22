// Liang (1983) "Word Hy-phen-a-tion by Com-pu-ter" — pattern-based
// hyphenation. A pattern like "hy3ph" has letters "hyph" and a weight of 3
// at the gap between "y" and "p"; when this pattern matches a substring of
// a (padded) word, that weight is OR-ed (max-merged) into the corresponding
// position. Odd final weights are hyphenation points; even values block.
//
// We use a compact trie keyed by character, storing the weight array on
// terminal nodes. Lookup is O(word.length × max_pattern_length) — fast
// enough for arbitrarily many patterns.

interface TrieNode {
  // Patterns terminating here. Each weight array's length equals the pattern
  // text length + 1 (one weight per gap, including before the first char and
  // after the last).
  weights: Array<number> | null;
  children: Map<string, TrieNode>;
}

function makeNode(): TrieNode {
  return { weights: null, children: new Map() };
}

interface ParsedPattern {
  readonly text: string;
  readonly weights: ReadonlyArray<number>;
}

function parsePattern(pattern: string): ParsedPattern {
  let text = '';
  const weights: Array<number> = [];
  let pending = 0;
  for (const ch of pattern) {
    if (ch >= '0' && ch <= '9') {
      pending = ch.charCodeAt(0) - 48;
    } else {
      weights.push(pending);
      text += ch;
      pending = 0;
    }
  }
  weights.push(pending);
  return { text, weights };
}

function insertPattern(root: TrieNode, pattern: string): void {
  const { text, weights } = parsePattern(pattern);
  let node = root;
  for (const ch of text) {
    let child = node.children.get(ch);
    if (!child) {
      child = makeNode();
      node.children.set(ch, child);
    }
    node = child;
  }
  node.weights = [...weights];
}

/** A configured hyphenator for one language. */
export interface Hyphenator {
  /**
   * Return the 0-indexed break positions inside `word`. A position `p` means the
   * word can be split between `word[p-1]` and `word[p]`; e.g. "computer" with
   * breaks `[3,5]` → "com·pu·ter".
   */
  hyphenate: (word: string) => Array<number>;
}

/** Options for {@link createHyphenator}. */
export interface HyphenatorOptions {
  /** Minimum characters before the first break (TeX default: 2). */
  readonly leftMin?: number;
  /** Minimum characters after the last break (TeX default: 3 for English, often 2 elsewhere). */
  readonly rightMin?: number;
  /**
   * Explicit exceptions that override pattern output. Each is a word with
   * hyphens marking the allowed breaks, e.g. `"as-so-ciate"`.
   */
  readonly exceptions?: ReadonlyArray<string>;
}

/**
 * Build a Liang pattern-based {@link Hyphenator} from a pattern list.
 *
 * @param patterns The TeX-style hyphenation patterns (e.g. `"hy3ph"`).
 * @param options  Left/right minimums and explicit exceptions.
 * @returns The hyphenator.
 */
export function createHyphenator(
  patterns: ReadonlyArray<string>,
  options: HyphenatorOptions = {},
): Hyphenator {
  const root = makeNode();
  for (const p of patterns) insertPattern(root, p);

  const leftMin = options.leftMin ?? 2;
  const rightMin = options.rightMin ?? 2;

  // Build exception lookup: lowercased word → break positions.
  const exceptionMap = new Map<string, Array<number>>();
  for (const e of options.exceptions ?? []) {
    const breaks: Array<number> = [];
    let stripped = '';
    for (const ch of e) {
      if (ch === '-') breaks.push(stripped.length);
      else stripped += ch;
    }
    exceptionMap.set(stripped.toLowerCase(), breaks);
  }

  return {
    hyphenate(word: string): Array<number> {
      if (word.length < leftMin + rightMin) return [];
      const lower = word.toLowerCase();
      const exception = exceptionMap.get(lower);
      if (exception) return exception.slice();
      const padded = `.${lower}.`;
      // weights[i] is the weight at the gap before padded[i]; weights[padded.length]
      // is the gap after the last char.
      const weights = new Array<number>(padded.length + 1).fill(0);
      for (let i = 0; i < padded.length; i++) {
        let node: TrieNode = root;
        for (let j = i; j < padded.length; j++) {
          const next = node.children.get(padded[j]!);
          if (!next) break;
          node = next;
          const w = node.weights;
          if (w) {
            for (let k = 0; k < w.length; k++) {
              const pos = i + k;
              if (pos < weights.length && w[k]! > weights[pos]!) {
                weights[pos] = w[k]!;
              }
            }
          }
        }
      }
      const breaks: Array<number> = [];
      // word[i] is padded[i+1]. Break between word[i-1] and word[i] means
      // splitting between padded[i] and padded[i+1] → weights[i+1].
      for (let i = leftMin; i <= word.length - rightMin; i++) {
        if (weights[i + 1]! % 2 === 1) breaks.push(i);
      }
      return breaks;
    },
  };
}

/**
 * Split a TeX-style pattern bundle (a whitespace-separated string) into the
 * pattern array {@link createHyphenator} expects.
 *
 * @param bundle The whitespace-separated patterns.
 * @returns The pattern array (empty entries removed).
 */
export function splitPatternBundle(bundle: string): Array<string> {
  return bundle.split(/\s+/).filter((s) => s.length > 0);
}
