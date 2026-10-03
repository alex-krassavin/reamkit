// After `astro build`: the whole sitemap as one file at /sitemap.xml, the
// address crawlers and Search Console try first. @astrojs/sitemap writes an
// index (sitemap-index.xml) pointing at its parts (sitemap-0.xml, …); this
// gathers the parts' <url> entries into a single <urlset> beside them, under
// the first part's own namespaces.
import { readFileSync, writeFileSync } from 'node:fs';

const dist = new URL('../dist/', import.meta.url);
const index = readFileSync(new URL('sitemap-index.xml', dist), 'utf8');
const parts = [...index.matchAll(/<loc>[^<]*\/(sitemap-\d+\.xml)<\/loc>/g)].map((m) => m[1]);
if (parts.length === 0) throw new Error('sitemap-index.xml names no sitemap-N.xml part');

let open = '';
const urls = [];
for (const part of parts) {
  const xml = readFileSync(new URL(part, dist), 'utf8');
  open ||= xml.match(/<urlset\b[^>]*>/)?.[0] ?? '';
  for (const m of xml.matchAll(/<url>[\s\S]*?<\/url>/g)) urls.push(m[0]);
}
if (!open || urls.length === 0) throw new Error(`no <url> entries in ${parts.join(', ')}`);
// A sitemap holds at most 50 000 URLs; past that the index is the only way.
if (urls.length > 50000) throw new Error(`${urls.length} URLs: too many for one sitemap`);

writeFileSync(
  new URL('sitemap.xml', dist),
  `<?xml version="1.0" encoding="UTF-8"?>${open}${urls.join('')}</urlset>`,
);
console.log(`sitemap.xml: ${urls.length} URLs from ${parts.join(', ')}`);
