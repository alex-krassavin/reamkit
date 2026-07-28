# Corpus validation

Compares our converter's output against LibreOffice's (`soffice`) as the "gold
standard", per the project's corpus-driven validation plan.

## Prerequisites (local only — not runtime deps)

- **LibreOffice** — reference renderer: `brew install --cask libreoffice`
- **mutool** (MuPDF) — rasterise + structured-text extraction: `brew install mupdf`
- **Roboto** installed in `~/Library/Fonts` so LibreOffice substitutes the same
  font the harness uses (copy from `tests/fixtures/fonts/Roboto-*.ttf`). Without
  this the visual diff is dominated by font-shape differences.

## Usage

```sh
npm run corpus:build   # regenerate corpus/inputs/ (synthetic docx/xlsx)
npm run corpus         # run the diff harness → corpus/report.md
npm run corpus -- --keep --dpi 150   # keep intermediate PNGs, higher DPI
```

## Metrics

Per document the harness reports:

- **TextSim** — LCS character similarity of extracted text vs reference
  (font-independent; catches missing/extra/reordered text). Higher is better.
- **Drift** — median baseline-y delta between matched lines (vertical layout
  fidelity). Lower is better. Confounded by reading-order differences (e.g.
  headers/footers), so treat large values on those docs with caution.
- **Visual** — worst-page pixel mismatch ratio of the RGB rasters (needs
  matching page size + font). Lower is better.
- **Pages** — page-count agreement.

The verdict is `✅` only when *all* of them agree: same page count, same page
size, TextSim > 95%, worst-page visual mismatch < 10% and geometry similarity
≥ 50%. `🈳` marks a document where **both** sides extracted no text — two empty
strings score a perfect similarity, so this must never read as a pass; it is
evidence of nothing and usually means the reference render failed too.

## Invariant sweep (xlsx) — no reference renderer needed

`npm run corpus` answers "does our render match LibreOffice?". The invariant
sweep answers a cheaper and stricter question that needs no oracle at all:
*did we silently lose the file?*

```sh
npm run corpus:xlsx:invariants            # check against the committed baseline
npm run corpus:xlsx:invariants -- --update   # bank progress / re-seed
XLSX_CORPUS_DIRS=corpus/external/lo-xlsx npm run corpus:xlsx:invariants  # narrow
```

Three invariants per document, each parsed in a child process under a 512 MB
heap cap and a wall-clock timeout:

1. **no silent loss** — a file with value cells must project to a non-empty
   document, or the reader must report a `Loss` saying what it dropped;
2. **bounded resources** — an unbounded allocation driven by a declared
   `dimension` is a denial-of-service vector, not a cosmetic bug;
3. **no unexpected throw** — well-formed packages parse; deliberately
   corrupt/fuzzed input may throw, and the baseline records which.

Violations are diffed against `corpus/xlsx-invariants-baseline.json`, so the
burn-down is visible in review and a regression fails the run. A baselined
entry is a known gap, not an accepted one.

## Notes

- Inputs declare an explicit A4 `sectPr` so both engines agree on page geometry;
  documents without one make LibreOffice fall back to a locale paper size.
- This is not part of `npm test` (it shells out to external binaries); run it
  on demand.
