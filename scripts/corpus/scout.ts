// Rank a corpus directory by how far our render sits from LibreOffice's.
//
// The golden set (tests/fixtures/real) is fifteen documents chosen because each
// one broke us. The wide corpus is six hundred, and nothing about it says which
// to look at next — the invariant sweep proves we do not crash or lose content
// against OURSELVES, not that the page matches anybody. This renders both sides
// and sorts by the distance, so "take the next file" has an answer.
//
// Read the number as a pointer, not a verdict. A page count that differs by ten
// is usually LibreOffice printing blank column bands we deliberately trim, and
// a character count that differs by a hundred is often its own clock in a
// footer. The list says where to LOOK; `corpus:visual` says what is there.
//
// Usage:
//   npx tsx scripts/corpus/scout.ts corpus/external/lo-xlsx [limit] [offset]

import { mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { basename, resolve } from 'node:path';

import { referenceToPdf, stext } from './lib';
import type { FontBytesByVariant } from '@/core/font';
import { Ream } from '@/core/converter/ream';

const FONTS: FontBytesByVariant = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
  italic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Italic.ttf')),
  boldItalic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-BoldItalic.ttf')),
};

interface Shape {
  readonly pages: number;
  readonly chars: number;
}

// Whitespace is not content, and counting it ranks by the wrong thing:
// LibreOffice pads an empty cell with spaces where we emit none, so 49156.xlsx
// read as 302 characters against 813 and sat near the top of this list with a
// text layer that is character-for-character identical to the reference's.
const visible = (chars: ReadonlyArray<{ readonly c: string }>): number =>
  chars.filter((ch) => ch.c.trim().length > 0).length;

const shapeOf = (pdf: string, out: string): Shape => {
  const pages = stext(pdf, out);
  return { pages: pages.length, chars: pages.reduce((n, p) => n + visible(p.chars), 0) };
};

async function main(): Promise<void> {
  const dir = process.argv[2];
  if (dir === undefined) throw new Error('usage: scout.ts <corpus-dir> [limit] [offset]');
  const limit = Number(process.argv[3] ?? 40);
  const offset = Number(process.argv[4] ?? 0);
  const work = resolve(`corpus/.scout-${String(process.pid)}`);
  rmSync(work, { recursive: true, force: true });
  mkdirSync(work, { recursive: true });
  const files = readdirSync(dir)
    .filter((f) => f.toLowerCase().endsWith('.xlsx'))
    .sort()
    .slice(offset, offset + limit);
  const rows: Array<{ score: number; line: string }> = [];
  for (const name of files) {
    const src = resolve(dir, name);
    let ours: Shape;
    try {
      // LibreOffice heads a page with the file's NAME when the sheet asks for
      // `&F`; a byte-oriented API has to be told it, or the comparison is
      // against a header we were never given.
      const pdf = await Ream.parse(new Uint8Array(readFileSync(src))).convert('pdf', {
        fileName: basename(src),
        fonts: FONTS,
      });
      const p = resolve(work, 'ours.pdf');
      writeFileSync(p, pdf);
      ours = shapeOf(p, resolve(work, 'ours.stext'));
    } catch (e) {
      rows.push({ score: -1, line: `  threw  ${name} — ${(e as Error).message.slice(0, 70)}` });
      continue;
    }
    let ref: Shape;
    try {
      ref = shapeOf(referenceToPdf(src, work), resolve(work, 'ref.stext'));
    } catch (e) {
      rows.push({ score: -1, line: `  ref!   ${name} — ${(e as Error).message.slice(0, 60)}` });
      continue;
    }
    // Characters first: a page-count gap is usually a pagination policy we chose
    // on purpose, while missing text is missing text.
    const score = Math.abs(ours.chars - ref.chars) + Math.abs(ours.pages - ref.pages);
    rows.push({
      score,
      line:
        `${String(score).padStart(7)}  ${name.padEnd(46)} ` +
        `pages ${ours.pages}/${ref.pages}  chars ${ours.chars}/${ref.chars}`,
    });
  }
  rmSync(work, { recursive: true, force: true });
  for (const r of rows.sort((a, b) => b.score - a.score)) console.log(r.line);
}

await main();
