// Hyphenation entry-points. The Liang algorithm lives in liang.ts; the
// language-specific pattern data sits in patterns-*.ts and is loaded on
// demand to keep the cold-start cost low for callers that never hyphenate.

import type { Hyphenator } from '@/core/hyphenation/liang';
import { createHyphenator, splitPatternBundle } from '@/core/hyphenation/liang';

export type { Hyphenator, HyphenatorOptions } from '@/core/hyphenation/liang';
export { createHyphenator, splitPatternBundle } from '@/core/hyphenation/liang';

/** A language for which bundled hyphenation patterns exist. */
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

/**
 * Lazily build (and cache) a language's {@link Hyphenator}. The patterns module
 * is dynamically imported the first time a language is requested, so callers
 * that never hyphenate pay no cold-start cost.
 *
 * @param language The language to load.
 * @returns The cached hyphenator.
 */
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

/**
 * Synchronous variant of {@link getHyphenator}: build a {@link Hyphenator} from
 * already-loaded pattern strings (useful when patterns are bundled or fetched
 * out-of-band).
 *
 * @param language The language (selects the left/right minimums).
 * @param bundle   The pattern and optional exception strings.
 * @returns The hyphenator.
 */
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

/**
 * Alias of {@link getHyphenator} — load a hyphenator for `language` using the
 * patterns bundled with this package.
 *
 * @param language The language to load.
 * @returns The hyphenator.
 */
export async function loadHyphenator(language: SupportedLanguage): Promise<Hyphenator> {
  return getHyphenator(language);
}
