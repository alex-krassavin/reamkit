// What search engines and link previews read off a page's <head>, put right
// where Starlight's defaults say less than the page does. Starlight builds the
// head before any route middleware runs, so these edit its built entries.
import { defineRouteMiddleware } from '@astrojs/starlight/route-data';

/** One entry of Starlight's built head. */
interface HeadEntry {
  tag: string;
  attrs?: Record<string, string | boolean | undefined>;
  content?: string;
}

/** The longest description a results page shows whole. */
const DESCRIPTION_CHARS = 155;

export const onRequest = defineRouteMiddleware((context) => {
  const route = context.locals.starlightRoute;
  const head = route.head as Array<HeadEntry>;
  const path = context.url.pathname;

  const meta = (key: 'name' | 'property', value: string): HeadEntry | undefined =>
    head.find((h) => h.tag === 'meta' && h.attrs?.[key] === value);
  const setMeta = (key: 'name' | 'property', value: string, content: string): void => {
    const found = meta(key, value);
    if (found?.attrs) found.attrs.content = content;
    else head.push({ tag: 'meta', attrs: { [key]: value, content } });
  };

  // ogp.me: a locale is language_TERRITORY, not a bare language code.
  setMeta('property', 'og:locale', 'en_US');

  // The home page is the site, not an article on it.
  if (path === '/') setMeta('property', 'og:type', 'website');

  // A missing page is nothing to index (GitHub Pages also answers it 404).
  if (route.id === '404') setMeta('name', 'robots', 'noindex');

  if (path.startsWith('/api/')) {
    // The reference is generated with no description of its own, and two
    // hundred pages fell back to the site's: each says what it documents now,
    // from the first sentence of its own doc comment.
    const name = path === '/api/' ? 'API reference' : route.entry.data.title;
    if (path === '/api/') route.entry.data.title = name;
    const summary = firstSentence(route.entry.body ?? '');
    const description = clip(
      path === '/api/'
        ? 'The public API of Ream (reamkit): the Ream class and its convert functions, the readers and writers, and the typed document model behind them.'
        : `${name} — Ream API reference${summary ? `: ${summary}` : '.'}`,
    );
    route.entry.data.description = description;
    setMeta('name', 'description', description);
    setMeta('property', 'og:description', description);
    // A type's bare name — Border, Table, Run — says nothing in a results
    // list on its own; its title says whose reference it is.
    const title = head.find((h) => h.tag === 'title');
    const shown = path === '/api/' ? `${name} | Ream` : `${name} — API reference | Ream`;
    if (title) title.content = shown;
    setMeta('property', 'og:title', path === '/api/' ? name : `${name} — Ream API`);
  }
});

/**
 * The first sentence of a generated reference page's prose: past the
 * "Defined in" line, headings, lists and tables, its links and code marks
 * read as their text.
 */
function firstSentence(body: string): string | undefined {
  for (const block of body.split(/\n\s*\n/)) {
    const text = block.trim();
    if (!text || /^(Defined in:|#|[-*|>]|```|\*\*\*|\d+\.)/.test(text)) continue;
    const plain = text
      .replace(/!?\[([^\]]*)\]\([^)]*\)/g, '$1')
      .replace(/[`*_]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
    if (!plain) continue;
    const end = plain.search(/[.!?](\s|$)/);
    return end >= 0 ? plain.slice(0, end + 1) : plain;
  }
  return undefined;
}

/** `text` cut at a word to fit a results page, an ellipsis where it was cut. */
function clip(text: string): string {
  if (text.length <= DESCRIPTION_CHARS) return text;
  const cut = text.slice(0, DESCRIPTION_CHARS - 1);
  const space = cut.lastIndexOf(' ');
  return `${(space > 80 ? cut.slice(0, space) : cut).replace(/[\s,;:—-]+$/, '')}…`;
}
