// Adopt documents from the fetched corpus into tests/fixtures/real/.
//
// The synthetic builders (tests/fixtures/build-*.ts) emit exactly the dialect
// our parsers expect, so a test written against them proves the parser can read
// our own writer — not that it can read Excel. These adopted documents are the
// counterweight: real files, from real producers, checked into the repo so the
// suite stays hermetic and offline.
//
// Each entry says WHY it was adopted. A fixture nobody can justify is a fixture
// nobody will maintain; the reason is what tells a future reader whether a
// change to it is a fix or a regression.
//
// Usage:
//   npx tsx scripts/corpus/sync-real-fixtures.ts           # verify (CI-safe)
//   npx tsx scripts/corpus/sync-real-fixtures.ts --adopt   # (re)copy + NOTICE

import { createHash } from 'node:crypto';
import { copyFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { SOURCES } from './sources';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '../..');
const fixtureDir = resolve(root, 'tests/fixtures/real');

const ADOPT = process.argv.includes('--adopt');

interface Adopted {
  /** File name, identical in the corpus and in tests/fixtures/real/. */
  readonly file: string;
  /** Corpus id from SOURCES — carries the repo, path and licence. */
  readonly source: string;
  /** What this document pins. Written for whoever has to judge a future diff. */
  readonly why: string;
}

// Kept alphabetical by file so the NOTICE diffs cleanly.
const ADOPTED: ReadonlyArray<Adopted> = [
  {
    file: '45540_form_Header.xlsx',
    source: 'poi-xlsx',
    why: 'Forty captionless ActiveX check boxes over a form — a control drawn with its `<control name>` writes an identifier across the page.',
  },
  {
    file: '47737.xlsx',
    source: 'poi-xlsx',
    why: 'Two sheets on `<pageSetup scale>` with no fit-to-page — a scaled sheet still paginates across its columns — and a second sheet whose only text is its header, which Excel refuses to print at all.',
  },
  {
    file: '49156.xlsx',
    source: 'poi-xlsx',
    why: 'Print area combined with manual row breaks — pagination driven by the document, not the page size.',
  },
  {
    file: '50299.xlsx',
    source: 'poi-xlsx',
    why: 'A rectangle whose fill and outline live only in `<xdr:style>` — gallery references into the theme, with nothing in its spPr — beside ten empty cells that carry nothing but a fill.',
  },
  {
    file: '53105.xlsx',
    source: 'poi-xlsx',
    why: 'Declares all 16 384 columns, so the grid materialization cap fires and must report the clip.',
  },
  {
    file: 'AverageTaxRates.xlsx',
    source: 'poi-xlsx',
    why: 'fitToPage scaling plus manual breaks across three sheets.',
  },
  {
    file: 'RepeatingRowsCols.xlsx',
    source: 'poi-xlsx',
    why: 'Print_Titles across four sheets — the header rows must repeat on every continuation page.',
  },
  {
    file: 'Spill.xlsx',
    source: 'lo-xlsx',
    why: 'A dynamic array whose spill is blocked: the cells cache a legacy `#VALUE!` and point at a rich value that names the real error.',
  },
  {
    file: 'bnc762542.xlsx',
    source: 'lo-xlsx',
    why: 'A3 landscape with fitToPage — paper size 8, the largest in the set.',
  },
  {
    file: 'singlecontrol.xlsx',
    source: 'lo-xlsx',
    why: 'One check box anchored 7331pt down a sheet with no cells — the drawings have to paginate on their own, downwards.',
  },
  {
    file: 'simple-monthly-budget.xlsx',
    source: 'poi-xlsx',
    why: 'An ordinary real-world workbook (landscape, fitToPage) rather than a bug reproduction.',
  },
  {
    file: 'Encrypted_LO_Standard_abc.docx',
    source: 'lo-docx-export',
    why: 'MS-OFFCRYPTO standard encryption (EncryptionInfo 3.2): AES-ECB under a key spun from 50 000 SHA-1 rounds. Password `abc`.',
  },
  {
    file: 'Encrypted_MSO2013_abc.docx',
    source: 'lo-docx-export',
    why: 'MS-OFFCRYPTO agile encryption (4.4) as Office 2013 writes it — SHA-512, AES-CBC, and a certificate key encryptor beside the password one. Password `abc`.',
  },
  {
    file: 'bar-chart.pptx',
    source: 'poi-pptx',
    why: "A deck whose chart carries its data as `ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx`, STORED — so the embedded workbook's own `xl/workbook.xml` lies in the outer file's bytes, where a substring sniff read it as a workbook.",
  },
  {
    file: 'protected_passtika.xlsx',
    source: 'poi-xlsx',
    why: 'An encrypted WORKBOOK — the same container question as the two documents above, on the other reader. Password `tika`.',
  },
  {
    file: 'tdf100034.xlsx',
    source: 'lo-xlsx',
    why: 'Letter paper (size 1) with a print area over two sheets — guards the A4-vs-Letter default.',
  },
  {
    file: 'tdf167019.xlsx',
    source: 'lo-xlsx',
    why: 'A4 landscape with both a print area and print titles.',
  },
  {
    file: 'tdf171828_fail_to_import_file.xlsx',
    source: 'lo-xlsx',
    why: 'Three sheets on three different papers (A4 landscape, Letter portrait, A4 landscape) — the mixed-geometry workbook.',
  },
  {
    file: 'tdf58243.xlsx',
    source: 'lo-xlsx',
    why: 'Print area, print titles and fitToPage together — the densest print-model document in the corpus.',
  },
  {
    file: 'open-as-read-only.xlsx',
    source: 'lo-xlsx',
    why: 'One cell in a one-column used range (`<dimension ref="A1"/>`) holding a sentence far wider than it — the plainest case of text overflowing past the end of the grid.',
  },
  {
    file: 'duplicate-filename.xlsx',
    source: 'poi-xlsx',
    why: 'Declares t="inlineStr" but writes the text into <v>; also ships two ZIP entries for the same part name.',
  },
  {
    file: 'tdf111980_radioButtons.xlsx',
    source: 'lo-xlsx',
    why: 'Reaches its ActiveX controls through §18.3.1.19 <control> rather than <oleObject>, with the state in binary activeX#.bin property bags.',
  },
  {
    file: 'tdf115159.xlsx',
    source: 'lo-xlsx',
    why: 'Two untouched tabs beside one sheet of data — an empty sheet must not print a page of its own.',
  },
  {
    file: 'tdf122336.xlsx',
    source: 'lo-xlsx',
    why: 'Namespace-prefixed SpreadsheetML (<x:worksheet>), GUID-shaped r:id values, and unparseable cell refs (r="11_2").',
  },
  {
    file: 'tdf76115.xlsx',
    source: 'lo-xlsx',
    why: 'Backslash ZIP separators, and keeps its worksheet at xl/sheet1.xml instead of xl/worksheets/.',
  },
  {
    file: 'tdf82984_zip64XLSXImport.xlsx',
    source: 'lo-xlsx',
    why: 'Zip64: every entry declares the 0xFFFFFFFF size sentinel rather than its real size.',
  },
  {
    file: 'too-many-cols-rows.xlsx',
    source: 'lo-xlsx',
    why: 'A 2.5 KB sheet declaring A1:XFE16777217 — the amplification case behind the total-cell budget.',
  },
];

const sha256 = (bytes: Uint8Array): string =>
  createHash('sha256').update(bytes).digest('hex').slice(0, 16);

function sourceFor(id: string): (typeof SOURCES)[number] {
  const source = SOURCES.find((s) => s.id === id);
  if (!source) throw new Error(`unknown corpus source "${id}"`);
  return source;
}

function corpusPath(entry: Adopted): string {
  return resolve(root, 'corpus/external', entry.source, entry.file);
}

function fixturePath(entry: Adopted): string {
  return resolve(fixtureDir, entry.file);
}

function renderNotice(digests: ReadonlyMap<string, string>): string {
  const lines = [
    '# Third-party test documents',
    '',
    'Real `.xlsx` files adopted from upstream test corpora, checked in so the',
    'suite stays hermetic and offline. They are here because the synthetic',
    'builders in `tests/fixtures/build-*.ts` can only produce the dialect our own',
    'parsers emit — these carry the dialects real producers emit.',
    '',
    'Each remains under its original licence, reproduced below. Regenerate this',
    'file with `npx tsx scripts/corpus/sync-real-fixtures.ts --adopt`.',
    '',
  ];
  for (const id of [...new Set(ADOPTED.map((a) => a.source))].sort()) {
    const source = sourceFor(id);
    lines.push(`## ${source.repo} — ${source.license}`);
    lines.push('');
    lines.push(`Upstream path: \`${source.path}\` (ref \`${source.ref}\`).`);
    lines.push('');
    lines.push('| File | sha256 (16) | Why it is here |');
    lines.push('|---|---|---|');
    for (const entry of ADOPTED.filter((a) => a.source === id)) {
      lines.push(`| \`${entry.file}\` | \`${digests.get(entry.file) ?? '?'}\` | ${entry.why} |`);
    }
    lines.push('');
  }
  return lines.join('\n');
}

function main(): void {
  if (ADOPT) {
    mkdirSync(fixtureDir, { recursive: true });
    const missing = ADOPTED.filter((a) => !existsSync(corpusPath(a)));
    if (missing.length > 0) {
      console.error(
        `Missing from the fetched corpus:\n  ${missing.map((m) => `${m.source}/${m.file}`).join('\n  ')}\n\n` +
          'Fetch it first: npx tsx scripts/corpus/fetch-corpus.ts --source lo-xlsx',
      );
      process.exit(1);
    }
    for (const entry of ADOPTED) copyFileSync(corpusPath(entry), fixturePath(entry));
  }

  const digests = new Map<string, string>();
  const absent: Array<string> = [];
  for (const entry of ADOPTED) {
    const path = fixturePath(entry);
    if (!existsSync(path)) {
      absent.push(entry.file);
      continue;
    }
    digests.set(entry.file, sha256(new Uint8Array(readFileSync(path))));
  }

  if (absent.length > 0) {
    console.error(
      `Adopted fixtures missing from tests/fixtures/real/:\n  ${absent.join('\n  ')}\n\n` +
        'Run with --adopt (needs the fetched corpus).',
    );
    process.exit(1);
  }

  const notice = renderNotice(digests);
  const noticePath = resolve(fixtureDir, 'NOTICE.md');
  if (ADOPT) {
    writeFileSync(noticePath, notice);
    console.log(`Adopted ${ADOPTED.length} document(s) → tests/fixtures/real/`);
    return;
  }
  // Verify mode: the NOTICE must match the bytes actually checked in, so a
  // fixture cannot be swapped without its provenance following it.
  const current = existsSync(noticePath) ? readFileSync(noticePath, 'utf8') : '';
  if (current !== notice) {
    console.error('tests/fixtures/real/NOTICE.md is stale — re-run with --adopt.');
    process.exit(1);
  }
  console.log(`✅ ${ADOPTED.length} adopted fixture(s) present, NOTICE current.`);
}

main();
