// Fetch a licence-clean corpus of real documents into corpus/external/
// (gitignored), with a provenance manifest. Covers every format Ream reads:
// the OOXML trio (.docx/.xlsx/.pptx), the legacy binaries (.doc/.xls/.ppt) and
// .pdf. Sources: Apache POI test-data (Apache-2.0), LibreOffice regression
// corpora (MPL-2.0), and Mozilla pdf.js's committed test PDFs (Apache-2.0) —
// all crafted to exercise format edge cases and real-world bug reports.
//
// SECURITY: macro-enabled formats (.docm/.xlsm/.pptm) are skipped, but treat
// ALL fetched documents as untrusted (the pdf.js set deliberately includes
// fuzzed/malformed files) — validate with CORPUS_SANDBOX=docker so the
// reference render runs isolated, and CORPUS_ISOLATE_OURS=1 so our own parse
// runs under a wall-clock + heap cap.
//
// Usage: tsx scripts/corpus/fetch-corpus.ts [--limit 60] [--source <id-prefix>]

import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, renameSync, writeFileSync } from 'node:fs';
import { basename, dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { SOURCES } from './sources';
import type { Source } from './sources';

const here = dirname(fileURLToPath(import.meta.url));
const outRoot = resolve(here, '../../corpus/external');

const limitArg = process.argv.indexOf('--limit');
const LIMIT = limitArg >= 0 ? Number(process.argv[limitArg + 1]) : 60;
// --source <id-prefix>: fetch only matching sources (e.g. `--source lo-`).
const sourceArg = process.argv.indexOf('--source');
const SOURCE_PREFIX = sourceArg >= 0 ? (process.argv[sourceArg + 1] ?? '') : '';

interface GhEntry {
  readonly name: string;
  readonly type: string;
  readonly download_url: string | null;
  readonly sha: string;
  readonly size: number;
}

const headers = { Accept: 'application/vnd.github+json', 'User-Agent': 'docgen-corpus' };

/**
 * Whether a listing's name is a file name of its directory and nothing more.
 * The name is the network's word, and `resolve(dir, '../../x.docx')` would
 * write outside the corpus.
 */
function isPlainName(name: string): boolean {
  return name !== '.' && name !== '..' && !name.includes('\\') && basename(name) === name;
}

/** Git's blob id of the bytes — what the contents listing's `sha` names. */
function gitBlobSha(bytes: Uint8Array): string {
  return createHash('sha1')
    .update(`blob ${String(bytes.length)}\0`)
    .update(bytes)
    .digest('hex');
}

async function listDir(s: Source): Promise<Array<GhEntry>> {
  const url = `https://api.github.com/repos/${s.repo}/contents/${s.path}?ref=${s.ref}`;
  const res = await fetch(url, { headers });
  if (!res.ok) throw new Error(`GitHub API ${res.status} for ${url}`);
  return (await res.json()) as Array<GhEntry>;
}

interface ManifestEntry {
  readonly source: string;
  readonly license: string;
  readonly repo: string;
  readonly path: string;
  readonly name: string;
  readonly sha: string;
  readonly size: number;
}

async function main(): Promise<void> {
  mkdirSync(outRoot, { recursive: true });
  const manifest: Array<ManifestEntry> = [];
  for (const s of SOURCES) {
    if (SOURCE_PREFIX && !s.id.startsWith(SOURCE_PREFIX)) continue;
    const dir = resolve(outRoot, s.id);
    mkdirSync(dir, { recursive: true });
    const all = await listDir(s);
    const picked = all
      .filter(
        (e) =>
          e.type === 'file' &&
          s.ext.test(e.name) &&
          !/\.(docm|xlsm|pptm)$/i.test(e.name) && // never fetch macro-enabled files
          e.download_url,
      )
      .slice(0, LIMIT);
    console.error(
      `${s.id}: ${picked.length}/${all.length} files (${s.repo}/${s.path}, ${s.license})`,
    );
    for (const e of picked) {
      if (!isPlainName(e.name)) {
        console.error(`  skip ${JSON.stringify(e.name)}: not a plain file name`);
        continue;
      }
      const dest = resolve(dir, e.name);
      if (!existsSync(dest)) {
        const res = await fetch(e.download_url!, { headers });
        if (!res.ok) {
          console.error(`  skip ${e.name}: HTTP ${res.status}`);
          continue;
        }
        // Only the blob the listing named is written — its size and its git
        // id — so a cut-short or substituted download never enters the corpus.
        const bytes = new Uint8Array(await res.arrayBuffer());
        if (bytes.length !== e.size || gitBlobSha(bytes) !== e.sha) {
          console.error(`  skip ${e.name}: not the blob the listing names`);
          continue;
        }
        // Written beside it and renamed into place: under its own name a file
        // is always whole, so the check above never keeps a half-written one.
        const part = `${dest}.${String(process.pid)}.part`;
        writeFileSync(part, bytes);
        renameSync(part, dest);
      }
      manifest.push({
        source: s.id,
        license: s.license,
        repo: s.repo,
        path: `${s.path}/${e.name}`,
        name: e.name,
        sha: e.sha,
        size: e.size,
      });
      process.stderr.write('.');
    }
    process.stderr.write('\n');
  }
  writeFileSync(resolve(outRoot, 'manifest.json'), JSON.stringify(manifest, null, 2) + '\n');
  console.error(`\nFetched ${manifest.length} files → corpus/external/ (gitignored).`);
  console.error('Validate (sandboxed) e.g.:');
  console.error(
    '  CORPUS_SANDBOX=docker CORPUS_DIR=corpus/external/poi-docx npx tsx scripts/corpus/run.ts',
  );
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
