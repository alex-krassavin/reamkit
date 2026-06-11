// Fetch a licence-clean corpus of real .docx/.xlsx into corpus/external/
// (gitignored), with a provenance manifest. Source: Apache POI's test-data
// (Apache-2.0) — files crafted to exercise OOXML edge cases.
//
// SECURITY: macro-enabled formats (.docm/.xlsm) are skipped, but treat ALL
// fetched documents as untrusted — validate them with CORPUS_SANDBOX=docker so
// the reference render (LibreOffice) and our own parse run isolated.
//
// Usage: tsx scripts/corpus/fetch-corpus.ts [--limit 60]

import { existsSync, mkdirSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const outRoot = resolve(here, '../../corpus/external');

const limitArg = process.argv.indexOf('--limit');
const LIMIT = limitArg >= 0 ? Number(process.argv[limitArg + 1]) : 60;
// --source <id-prefix>: fetch only matching sources (e.g. `--source lo-`).
const sourceArg = process.argv.indexOf('--source');
const SOURCE_PREFIX = sourceArg >= 0 ? (process.argv[sourceArg + 1] ?? '') : '';

interface Source {
  readonly id: string;
  readonly repo: string; // owner/repo
  readonly path: string; // directory within the repo
  readonly ref: string; // branch or tag
  readonly ext: RegExp; // which files to take
  readonly license: string;
}

const SOURCES: ReadonlyArray<Source> = [
  {
    id: 'poi-docx',
    repo: 'apache/poi',
    path: 'test-data/document',
    ref: 'trunk',
    ext: /\.docx$/i,
    license: 'Apache-2.0',
  },
  {
    id: 'poi-xlsx',
    repo: 'apache/poi',
    path: 'test-data/spreadsheet',
    ref: 'trunk',
    ext: /\.xlsx$/i,
    license: 'Apache-2.0',
  },
  // LibreOffice's OOXML regression corpora — thousands of real-world-shaped
  // documents distilled from actual bug reports. (GovDocs1 was evaluated and
  // rejected: it predates OOXML — a legacy .doc/.xls corpus.)
  {
    id: 'lo-docx-export',
    repo: 'LibreOffice/core',
    path: 'sw/qa/extras/ooxmlexport/data',
    ref: 'master',
    ext: /\.docx$/i,
    license: 'MPL-2.0',
  },
  {
    id: 'lo-docx-import',
    repo: 'LibreOffice/core',
    path: 'sw/qa/extras/ooxmlimport/data',
    ref: 'master',
    ext: /\.docx$/i,
    license: 'MPL-2.0',
  },
  {
    id: 'lo-xlsx',
    repo: 'LibreOffice/core',
    path: 'sc/qa/unit/data/xlsx',
    ref: 'master',
    ext: /\.xlsx$/i,
    license: 'MPL-2.0',
  },
];

interface GhEntry {
  readonly name: string;
  readonly type: string;
  readonly download_url: string | null;
  readonly sha: string;
  readonly size: number;
}

const headers = { Accept: 'application/vnd.github+json', 'User-Agent': 'docgen-corpus' };

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
          !/\.(docm|xlsm)$/i.test(e.name) && // never fetch macro-enabled files
          e.download_url,
      )
      .slice(0, LIMIT);
    console.error(
      `${s.id}: ${picked.length}/${all.length} files (${s.repo}/${s.path}, ${s.license})`,
    );
    for (const e of picked) {
      const dest = resolve(dir, e.name);
      if (!existsSync(dest)) {
        const res = await fetch(e.download_url!, { headers });
        if (!res.ok) {
          console.error(`  skip ${e.name}: HTTP ${res.status}`);
          continue;
        }
        writeFileSync(dest, new Uint8Array(await res.arrayBuffer()));
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
