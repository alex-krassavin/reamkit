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
| `tdf122336.xlsx` | `d0d409d88cb22f11` | Namespace-prefixed SpreadsheetML (<x:worksheet>), GUID-shaped r:id values, and unparseable cell refs (r="11_2"). |
| `tdf76115.xlsx` | `1428f774dd01d4c5` | Backslash ZIP separators, and keeps its worksheet at xl/sheet1.xml instead of xl/worksheets/. |
| `tdf82984_zip64XLSXImport.xlsx` | `32299060140eb381` | Zip64: every entry declares the 0xFFFFFFFF size sentinel rather than its real size. |
| `too-many-cols-rows.xlsx` | `2e0bb99477a00d8b` | A 2.5 KB sheet declaring A1:XFE16777217 — the amplification case behind the total-cell budget. |

## apache/poi — Apache-2.0

Upstream path: `test-data/spreadsheet` (ref `trunk`).

| File | sha256 (16) | Why it is here |
|---|---|---|
| `53105.xlsx` | `7086065eb8727133` | Declares all 16 384 columns, so the grid materialization cap fires and must report the clip. |
| `duplicate-filename.xlsx` | `5aa65f91139a76cd` | Declares t="inlineStr" but writes the text into <v>; also ships two ZIP entries for the same part name. |
