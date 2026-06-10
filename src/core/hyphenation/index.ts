// Hyphenation entry-points. The Liang algorithm lives in liang.ts; the
// language-specific pattern data sits in patterns-*.ts and is loaded on
// demand to keep the cold-start cost low for callers that never hyphenate.

import type { Hyphenator } from '@/core/hyphenation/liang';
import { createHyphenator, splitPatternBundle } from '@/core/hyphenation/liang';

export type { Hyphenator, HyphenatorOptions } from '@/core/hyphenation/liang';
export { createHyphenator, splitPatternBundle } from '@/core/hyphenation/liang';

export type SupportedLanguage = 'en-us' | 'ru';

interface PatternBundle {
  readonly PATTERNS: string;
  readonly EXCEPTIONS: string;
}

interface LanguageDescriptor {
  readonly leftMin: number;
  readonly rightMin: number;
  load: () => Promise<PatternBundle>;
}

const REGISTRY: Record<SupportedLanguage, LanguageDescriptor> = {
  'en-us': {
    leftMin: 2,
    rightMin: 3,
    load: () => import('@/core/hyphenation/patterns-en-us'),
  },
  ru: {
    leftMin: 2,
    rightMin: 2,
    load: () => import('@/core/hyphenation/patterns-ru'),
  },
};

const cache = new Map<SupportedLanguage, Hyphenator>();

// Lazy-load a language's hyphenator. The patterns module is dynamically
// imported the first time a language is requested.
export async function getHyphenator(language: SupportedLanguage): Promise<Hyphenator> {
  const cached = cache.get(language);
  if (cached) return cached;
  const desc = REGISTRY[language];
  const mod = await desc.load();
  const h = createHyphenator(splitPatternBundle(mod.PATTERNS), {
    leftMin: desc.leftMin,
    rightMin: desc.rightMin,
    exceptions: splitPatternBundle(mod.EXCEPTIONS),
  });
  cache.set(language, h);
  return h;
}

// Synchronous variant: construct from already-loaded pattern strings (useful
// when patterns are bundled or fetched out-of-band).
export function createLanguageHyphenator(
  language: SupportedLanguage,
  bundle: { patterns: string; exceptions?: string },
): Hyphenator {
  const desc = REGISTRY[language];
  return createHyphenator(splitPatternBundle(bundle.patterns), {
    leftMin: desc.leftMin,
    rightMin: desc.rightMin,
    ...(bundle.exceptions ? { exceptions: splitPatternBundle(bundle.exceptions) } : {}),
  });
}

// Synchronously load and create a hyphenator using a top-level `require`-style
// import. Uses the patterns bundled with this package — the cost is paid up
// front (parse the pattern string + build the trie).
export async function loadHyphenator(language: SupportedLanguage): Promise<Hyphenator> {
  return getHyphenator(language);
}
