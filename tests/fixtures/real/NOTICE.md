# Third-party test documents

Real `.xlsx` files adopted from upstream test corpora, checked in so the
suite stays hermetic and offline. They are here because the synthetic
builders in `tests/fixtures/build-*.ts` can only produce the dialect our own
parsers emit — these carry the dialects real producers emit.

Each remains under its original licence, reproduced below. Regenerate this
file with `npx tsx scripts/corpus/sync-real-fixtures.ts --adopt`.

## LibreOffice/core — MPL-2.0

Upstream path: `sc/qa/unit/data/xlsx` (ref `master`).

| File | sha256 (16) | Why it is here |
|---|---|---|
| `bnc762542.xlsx` | `89afbf1b49458804` | A3 landscape with fitToPage — paper size 8, the largest in the set. |
| `tdf100034.xlsx` | `8efb54cb804b713f` | Letter paper (size 1) with a print area over two sheets — guards the A4-vs-Letter default. |
| `tdf167019.xlsx` | `298671dfe0882a2f` | A4 landscape with both a print area and print titles. |
| `tdf58243.xlsx` | `442be369ae4f768d` | Print area, print titles and fitToPage together — the densest print-model document in the corpus. |
| `tdf122336.xlsx` | `d0d409d88cb22f11` | Namespace-prefixed SpreadsheetML (<x:worksheet>), GUID-shaped r:id values, and unparseable cell refs (r="11_2"). |
| `tdf76115.xlsx` | `1428f774dd01d4c5` | Backslash ZIP separators, and keeps its worksheet at xl/sheet1.xml instead of xl/worksheets/. |
| `tdf82984_zip64XLSXImport.xlsx` | `32299060140eb381` | Zip64: every entry declares the 0xFFFFFFFF size sentinel rather than its real size. |
| `too-many-cols-rows.xlsx` | `2e0bb99477a00d8b` | A 2.5 KB sheet declaring A1:XFE16777217 — the amplification case behind the total-cell budget. |

## apache/poi — Apache-2.0

Upstream path: `test-data/spreadsheet` (ref `trunk`).

| File | sha256 (16) | Why it is here |
|---|---|---|
| `49156.xlsx` | `7376a8118afc503a` | Print area combined with manual row breaks — pagination driven by the document, not the page size. |
| `53105.xlsx` | `7086065eb8727133` | Declares all 16 384 columns, so the grid materialization cap fires and must report the clip. |
| `AverageTaxRates.xlsx` | `094b2facaf85870a` | fitToPage scaling plus manual breaks across three sheets. |
| `RepeatingRowsCols.xlsx` | `ff67241b278977c9` | Print_Titles across four sheets — the header rows must repeat on every continuation page. |
| `simple-monthly-budget.xlsx` | `cae00c6894b95743` | An ordinary real-world workbook (landscape, fitToPage) rather than a bug reproduction. |
| `duplicate-filename.xlsx` | `5aa65f91139a76cd` | Declares t="inlineStr" but writes the text into <v>; also ships two ZIP entries for the same part name. |
