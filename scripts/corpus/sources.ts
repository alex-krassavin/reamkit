// Provenance table for the fetched corpora — the single place that records
// where each document set comes from and under what licence.
//
// Split out of fetch-corpus.ts (which runs main() on import) so other tooling
// can read the table without triggering a fetch: sync-real-fixtures.ts cites
// it when it writes the NOTICE for the documents adopted into tests/.

export interface Source {
  readonly id: string;
  readonly repo: string; // owner/repo
  readonly path: string; // directory within the repo
  readonly ref: string; // branch or tag
  readonly ext: RegExp; // which files to take
  readonly license: string;
}

export const SOURCES: ReadonlyArray<Source> = [
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
  // --- PresentationML (.pptx) ---
  {
    id: 'poi-pptx',
    repo: 'apache/poi',
    path: 'test-data/slideshow',
    ref: 'trunk',
    ext: /\.pptx$/i,
    license: 'Apache-2.0',
  },
  {
    id: 'lo-pptx',
    repo: 'LibreOffice/core',
    path: 'sd/qa/unit/data/pptx',
    ref: 'master',
    ext: /\.pptx$/i,
    license: 'MPL-2.0',
  },
  // --- Legacy binary Word (.doc — MS-DOC / WW8) ---
  {
    id: 'poi-doc',
    repo: 'apache/poi',
    path: 'test-data/document',
    ref: 'trunk',
    ext: /\.doc$/i,
    license: 'Apache-2.0',
  },
  {
    id: 'lo-doc',
    repo: 'LibreOffice/core',
    path: 'sw/qa/extras/ww8export/data',
    ref: 'master',
    ext: /\.doc$/i,
    license: 'MPL-2.0',
  },
  // --- Legacy binary Excel (.xls — BIFF8) ---
  {
    id: 'poi-xls',
    repo: 'apache/poi',
    path: 'test-data/spreadsheet',
    ref: 'trunk',
    ext: /\.xls$/i,
    license: 'Apache-2.0',
  },
  {
    id: 'lo-xls',
    repo: 'LibreOffice/core',
    path: 'sc/qa/unit/data/xls',
    ref: 'master',
    ext: /\.xls$/i,
    license: 'MPL-2.0',
  },
  // --- Legacy binary PowerPoint (.ppt — MS-PPT / HSLF) ---
  {
    id: 'poi-ppt',
    repo: 'apache/poi',
    path: 'test-data/slideshow',
    ref: 'trunk',
    ext: /\.ppt$/i,
    license: 'Apache-2.0',
  },
  {
    id: 'lo-ppt',
    repo: 'LibreOffice/core',
    path: 'sd/qa/unit/data/ppt',
    ref: 'master',
    ext: /\.ppt$/i,
    license: 'MPL-2.0',
  },
  // --- PDF (read path) — Mozilla pdf.js's committed real-world test corpus,
  // including deliberately malformed / fuzzed files. Treat as hostile: render
  // the reference sandboxed and isolate our own parse (CORPUS_ISOLATE_OURS=1).
  {
    id: 'pdfjs',
    repo: 'mozilla/pdf.js',
    path: 'test/pdfs',
    ref: 'master',
    ext: /\.pdf$/i,
    license: 'Apache-2.0',
  },
];
