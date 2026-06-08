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

## Notes

- xlsx comparison is intentionally apples-to-oranges: we render a sheet as a
  bordered grid plus a sheet-name title, whereas LibreOffice Calc prints via its
  own print-area model. Divergence there is expected.
- Inputs declare an explicit A4 `sectPr` so both engines agree on page geometry;
  documents without one make LibreOffice fall back to a locale paper size.
- This is not part of `npm test` (it shells out to external binaries); run it
  on demand.
