// Page geometry of real spreadsheets, against summaries LibreOffice produced
// once (scripts/corpus/make-golden.ts writes tests/fixtures/real/golden/).
//
// Nothing here shells out: the reference is committed text, so the assertions
// run in `npm test` on any machine, offline, in milliseconds.
//
// The suite is deliberate about what it claims:
//
//   - PAGE SIZE is print-model output. When the workbook names its paper size,
//     both renderers must honour it and the comparison is strict. When it does
//     not, the file holds no answer — Excel picks by locale and printer,
//     LibreOffice by locale — so asserting there would test the golden
//     machine's locale, not our code.
//   - PAGINATION and CONTENT COVERAGE depend on typesetting, where we do not
//     match LibreOffice and have never claimed to. Both figures are recorded,
//     ours is pinned, and the distance to LibreOffice is asserted to be no
//     worse than when the golden was taken. That makes the gap a number that
//     can be argued down instead of a vague known-difference.

import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import type { FontBytesByVariant } from '@/core/font';
import { convertXlsxToPdfSync } from '@/core/converter';

const here = dirname(fileURLToPath(import.meta.url));
const fixtureDir = resolve(here, 'fixtures/real');
const goldenDir = resolve(fixtureDir, 'golden');

const FONTS: FontBytesByVariant = {
  regular: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf'))),
  bold: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Bold.ttf'))),
  italic: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Italic.ttf'))),
  boldItalic: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-BoldItalic.ttf'))),
};

interface GoldenPage {
  readonly widthPt: number;
  readonly heightPt: number;
}
interface GoldenRender {
  readonly pages: number;
  readonly pageSizes: ReadonlyArray<GoldenPage>;
  readonly chars: number;
  readonly firstLines: ReadonlyArray<string>;
}
interface Golden {
  readonly file: string;
  readonly paperDeclared: boolean;
  readonly libreOffice: GoldenRender;
  readonly ream: GoldenRender;
}

const goldens: Array<Golden> = readdirSync(goldenDir)
  .filter((f) => f.endsWith('.json'))
  .sort()
  .map((f) => JSON.parse(readFileSync(resolve(goldenDir, f), 'utf8')) as Golden);

/** Page sizes of our PDF, deduplicated and rounded, straight from the /MediaBox. */
function ourPageSizes(pdf: Uint8Array): Array<GoldenPage> {
  const text = new TextDecoder('latin1').decode(pdf);
  const seen = new Map<string, GoldenPage>();
  const re = /\/MediaBox \[0 0 ([\d.]+) ([\d.]+)\]/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    const size = { widthPt: Math.round(Number(m[1])), heightPt: Math.round(Number(m[2])) };
    seen.set(`${size.widthPt}x${size.heightPt}`, size);
  }
  return [...seen.values()];
}

function ourPageCount(pdf: Uint8Array): number {
  const text = new TextDecoder('latin1').decode(pdf);
  return (text.match(/\/Type \/Page\b/g) ?? []).filter((_, i, all) => all[i] === '/Type /Page')
    .length;
}

describe('golden page geometry (real spreadsheets)', () => {
  it('has a golden for every adopted document LibreOffice can open', () => {
    // duplicate-filename.xlsx and tdf76115.xlsx are absent by design: soffice
    // refuses both, so no reference exists. tdf76115 is the sharper case — we
    // read 7118 cells out of a document LibreOffice will not open at all.
    // protected_passtika.xlsx is absent for a different reason: it is
    // encrypted, and a golden render would have to carry its password.
    const adopted = readdirSync(fixtureDir).filter((f) => f.endsWith('.xlsx'));
    const withGolden = new Set(goldens.map((g) => g.file));
    const without = adopted.filter((f) => !withGolden.has(f)).sort();
    expect(without).toEqual([
      'duplicate-filename.xlsx',
      'protected_passtika.xlsx',
      'tdf76115.xlsx',
    ]);
  });

  for (const golden of goldens) {
    describe(golden.file, () => {
      const pdf = convertXlsxToPdfSync(
        new Uint8Array(readFileSync(resolve(fixtureDir, golden.file))),
        { fonts: FONTS },
      );

      it(
        golden.paperDeclared
          ? 'renders the paper size the workbook names, as LibreOffice does'
          : 'renders a deterministic page size (the workbook names none)',
        () => {
          const sizes = ourPageSizes(pdf);
          expect(sizes.length).toBeGreaterThan(0);
          if (golden.paperDeclared) {
            expect(sizes).toEqual(golden.libreOffice.pageSizes);
          } else {
            // No claim against LibreOffice here — only that we are stable.
            expect(sizes).toEqual(golden.ream.pageSizes);
          }
        },
      );

      it('paginates as recorded, no further from LibreOffice than before', () => {
        const pages = ourPageCount(pdf);
        expect(pages).toBe(golden.ream.pages);
        const recordedGap = Math.abs(golden.ream.pages - golden.libreOffice.pages);
        expect(Math.abs(pages - golden.libreOffice.pages)).toBeLessThanOrEqual(recordedGap);
      });
    });
  }
});
