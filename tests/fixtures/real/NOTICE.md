# Third-party test documents

Real `.xlsx` files adopted from upstream test corpora, checked in so the
suite stays hermetic and offline. They are here because the synthetic
builders in `tests/fixtures/build-*.ts` can only produce the dialect our own
parsers emit — these carry the dialects real producers emit.

Each remains under its original licence, reproduced below. Regenerate this
file with `npx tsx scripts/corpus/sync-real-fixtures.ts --adopt`.

## LibreOffice/core — MPL-2.0

Upstream path: `sw/qa/extras/ooxmlexport/data` (ref `master`).

| File | sha256 (16) | Why it is here |
|---|---|---|
| `Encrypted_LO_Standard_abc.docx` | `d6b55065e5c8e6de` | MS-OFFCRYPTO standard encryption (EncryptionInfo 3.2): AES-ECB under a key spun from 50 000 SHA-1 rounds. Password `abc`. |
| `Encrypted_MSO2013_abc.docx` | `2db365c48e5f3b03` | MS-OFFCRYPTO agile encryption (4.4) as Office 2013 writes it — SHA-512, AES-CBC, and a certificate key encryptor beside the password one. Password `abc`. |

## LibreOffice/core — MPL-2.0

Upstream path: `sd/qa/unit/data/pptx` (ref `master`).

| File | sha256 (16) | Why it is here |
|---|---|---|
| `master-bg-color.pptx` | `88cba9c63fdb464b` | A master background of `schemeClr bg1` under a map that says bg1 means dk2 — the deck is blue, and read without the map it is white. |

## LibreOffice/core — MPL-2.0

Upstream path: `sc/qa/unit/data/xlsx` (ref `master`).

| File | sha256 (16) | Why it is here |
|---|---|---|
| `Spill.xlsx` | `c1e17e7d0d8b8f1c` | A dynamic array whose spill is blocked: the cells cache a legacy `#VALUE!` and point at a rich value that names the real error. |
| `bnc762542.xlsx` | `89afbf1b49458804` | A3 landscape with fitToPage — paper size 8, the largest in the set. |
| `singlecontrol.xlsx` | `a13fd1a411a499f4` | One check box anchored 7331pt down a sheet with no cells — the drawings have to paginate on their own, downwards. |
| `tdf100034.xlsx` | `8efb54cb804b713f` | Letter paper (size 1) with a print area over two sheets — guards the A4-vs-Letter default. |
| `tdf167019.xlsx` | `298671dfe0882a2f` | A4 landscape with both a print area and print titles. |
| `tdf171828_fail_to_import_file.xlsx` | `ebfabb870c52d902` | Three sheets on three different papers (A4 landscape, Letter portrait, A4 landscape) — the mixed-geometry workbook. |
| `tdf58243.xlsx` | `442be369ae4f768d` | Print area, print titles and fitToPage together — the densest print-model document in the corpus. |
| `open-as-read-only.xlsx` | `0a6729f7e3fbe08a` | One cell in a one-column used range (`<dimension ref="A1"/>`) holding a sentence far wider than it — the plainest case of text overflowing past the end of the grid. |
| `tdf111980_radioButtons.xlsx` | `78067fa92760a582` | Reaches its ActiveX controls through §18.3.1.19 <control> rather than <oleObject>, with the state in binary activeX#.bin property bags. |
| `tdf115159.xlsx` | `98053ff4bab193b5` | Two untouched tabs beside one sheet of data — an empty sheet must not print a page of its own. |
| `tdf122336.xlsx` | `d0d409d88cb22f11` | Namespace-prefixed SpreadsheetML (<x:worksheet>), GUID-shaped r:id values, and unparseable cell refs (r="11_2"). |
| `tdf76115.xlsx` | `1428f774dd01d4c5` | Backslash ZIP separators, and keeps its worksheet at xl/sheet1.xml instead of xl/worksheets/. |
| `tdf82984_zip64XLSXImport.xlsx` | `32299060140eb381` | Zip64: every entry declares the 0xFFFFFFFF size sentinel rather than its real size. |
| `too-many-cols-rows.xlsx` | `2e0bb99477a00d8b` | A 2.5 KB sheet declaring A1:XFE16777217 — the amplification case behind the total-cell budget. |

## apache/poi — Apache-2.0

Upstream path: `test-data/slideshow` (ref `trunk`).

| File | sha256 (16) | Why it is here |
|---|---|---|
| `bar-chart.pptx` | `79e1d218bfb2903e` | A deck whose chart carries its data as `ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx`, STORED — so the embedded workbook's own `xl/workbook.xml` lies in the outer file's bytes, where a substring sniff read it as a workbook. |

## apache/poi — Apache-2.0

Upstream path: `test-data/spreadsheet` (ref `trunk`).

| File | sha256 (16) | Why it is here |
|---|---|---|
| `45540_form_Header.xlsx` | `5e9a5cc4f70614fa` | Forty captionless ActiveX check boxes over a form — a control drawn with its `<control name>` writes an identifier across the page. |
| `47737.xlsx` | `4bcb8da52e258c61` | Two sheets on `<pageSetup scale>` with no fit-to-page — a scaled sheet still paginates across its columns — and a second sheet whose only text is its header, which Excel refuses to print at all. |
| `49156.xlsx` | `7376a8118afc503a` | Print area combined with manual row breaks — pagination driven by the document, not the page size. |
| `50299.xlsx` | `eb32ad8da197c3e3` | A rectangle whose fill and outline live only in `<xdr:style>` — gallery references into the theme, with nothing in its spPr — beside ten empty cells that carry nothing but a fill. |
| `53105.xlsx` | `7086065eb8727133` | Declares all 16 384 columns, so the grid materialization cap fires and must report the clip. |
| `AverageTaxRates.xlsx` | `094b2facaf85870a` | fitToPage scaling plus manual breaks across three sheets. |
| `RepeatingRowsCols.xlsx` | `ff67241b278977c9` | Print_Titles across four sheets — the header rows must repeat on every continuation page. |
| `simple-monthly-budget.xlsx` | `cae00c6894b95743` | An ordinary real-world workbook (landscape, fitToPage) rather than a bug reproduction. |
| `protected_passtika.xlsx` | `e58713895915de62` | An encrypted WORKBOOK — the same container question as the two documents above, on the other reader. Password `tika`. |
| `duplicate-filename.xlsx` | `5aa65f91139a76cd` | Declares t="inlineStr" but writes the text into <v>; also ships two ZIP entries for the same part name. |
