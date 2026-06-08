# Ream (reamkit) — handoff

> **Статус:** опубликовано — `reamkit@0.1.0-alpha.0` в npm (MIT, provenance, dist-tag `alpha`). Репо: github.com/alex-krassavin/reamkit. Тулинг под `@tanstack/config`. M0–M6 завершены.

## Цель проекта

Библиотека для конвертации документов **Word (docx) и Excel (xlsx) → PDF**. Реализация **с нуля по официальным спецификациям**, без обёрток над LibreOffice / ZetaJS / Aspose / Microsoft Office.

Первичный таргет — браузер (клиентская конвертация); поскольку реализация на чистом JS, тот же код запускается в Node.js, serverless и edge-окружениях без изменений.

## Что отвергнуто и почему

| Подход | Почему не подходит |
|---|---|
| **ZetaJS / ZetaOffice WASM** в браузере | Артефакты ~250 МБ, старт 10–15 с, потолок памяти на мобильных Safari ~500 МБ |
| **LibreOffice headless в контейнере** (serverless) | Образ ~1 ГБ, медленный cold start, зависимость от системного бинаря |
| **Gotenberg / shelfio Lambda Layer** | Та же проблема — обёртка над LO, не наш код |
| **docx-wasm (NativeDocuments), Apryse, Nutrient** | Коммерческие SDK с лицензионными отчислениями |
| **DOCX → HTML → PDF** (mammoth.js + headless Chromium) | Промежуточная HTML-модель теряет ~10–30% качества для сложного форматирования |
| **DOCX → XSL-FO → PDF** (docx4j + Apache FOP) | JVM-зависимость, известные косяки рендера |
| **Bounded subset под наши шаблоны** | Пользователь явно отверг подход «компромисс по фичам» |

**Решение:** идём по пути наибольшего сопротивления — пишем конвертер по спецификациям. Существующие реализации (LibreOffice C++ source, Apache POI) используем только как **референс для изучения подводных камней**, не как зависимость и не копируя код (LO под MPL-2.0, лицензионно несовместимо).

## Спецификации — источники истины

### Входные форматы

- **ECMA-376 5th Edition (2016)** — OOXML, бесплатно: https://ecma-international.org/publications-and-standards/standards/ecma-376/
  - Part 1: Fundamentals & Markup Language Reference (~5000 стр) — WordprocessingML §17, SpreadsheetML §18, DrawingML §20, OfficeMathML §22, SharedML §23
  - Part 2: Open Packaging Conventions (OPC) (~130 стр)
  - Part 3: Markup Compatibility & Extensibility (MCE) (~30 стр)
  - Part 4: Transitional Migration Features (~1500 стр) — legacy VML и т.п.
- **Microsoft Open Specifications** (обязательны — Word/Excel отклоняются от ECMA):
  - [MS-DOCX] Word Extensions to OOXML
  - [MS-XLSX] Excel Extensions
  - [MS-OE376] Office Implementation Information for ECMA-376
  - [MS-OI29500] Office Implementation Information for ISO/IEC-29500

### Выходной формат

- **ISO 32000-1:2008** (= Adobe PDF Reference 1.7, бесплатно) — стартовая точка
- **ISO 32000-2:2020** (PDF 2.0) — целевой стандарт
- **ISO 19005-1/2/3** (PDF/A) — архивный профиль для Phase 2

### Вспомогательные стандарты

- **Unicode 15.x + UAX #14** (Line Breaking), **UAX #9** (BiDi), **UAX #29** (Text Segmentation)
- **OpenType Specification** (Microsoft/Adobe) — TTF/OTF
- **ISO/IEC 14496-22** — OFF (Open Font Format)
- **Knuth & Plass (1981)** — алгоритм абзацной вёрстки
- **Liang (1983)** — алгоритм переноса слов
- **RFC 1951 (Deflate)**, **ITU-T T.6 (Group 4 fax)** — для PDF streams и OOXML ZIP

## Технический стек

**Язык:** JavaScript / TypeScript. Один и тот же код работает в браузере, Node.js, serverless и edge.

**Минимальные внешние зависимости** (используем как кирпичи для incidental complexity, не для основной логики конвертации):
- `fflate` — ZIP/Deflate (для OPC контейнера)
- `fast-xml-parser` — XML-парсинг (XML — отдельный W3C-стандарт, не часть нашей задачи)
- `typescript`, `tsx`, `vitest` — тулинг

**Чего НЕ берём готового** (это сердце проекта):
- PDF-writer'ы (`pdf-lib`, `pdfkit`)
- Layout-движки
- Text shapers
- Конвертеры docx/xlsx

## Архитектура — модули

Каждый модуль маппится на конкретный раздел стандарта:

```
src/
  opc/              # ECMA-376 Part 2 — OPC packaging
  ooxml/
    wordproc/       # ECMA-376 Part 1 §17 — WordprocessingML
    spreadsheet/    # ECMA-376 Part 1 §18 — SpreadsheetML
    drawing/        # ECMA-376 Part 1 §20 — DrawingML
    math/           # ECMA-376 Part 1 §22 — OfficeMathML
    mce/            # ECMA-376 Part 3 — Markup Compatibility
  document-model/   # Типизированная in-memory модель (Document, Section, Paragraph, Run, Table, ...)
  style-cascade/    # §17.7 — docDefaults → styles → direct formatting
  numbering/        # §17.9 — abstractNum, lvlOverride
  font-engine/      # OpenType — парсинг TTF/OTF, метрики, шейпинг
  text-shaper/      # UAX + GSUB/GPOS — Unicode → массив глифов
  line-breaker/     # Knuth-Plass + Liang
  layout-engine/    # Box model → strings → blocks → pages
  pdf/
    objects/        # ISO 32000-2 §7 — типы объектов
    writer/         # ISO 32000-2 §7-8 — content streams, xref
    fonts/          # §9 — Font subsetting, Type0/CIDFontType2 embedding
    images/         # §8.9 — XObject Image, DCTDecode/FlateDecode/CCITTFaxDecode
    graphics/       # §8.5 — path/fill/stroke
  converter/        # Оркестратор: OOXML → document-model → layout → PDF
```

## Этапы (milestones)

| Этап | Содержание | Definition of Done |
|---|---|---|
| **M0** — фундамент | OPC reader, text extractor, минимальный PDF writer с Helvetica builtin | hello.docx → PDF с тем же plain text, `qpdf --check` зелёный |
| **M1** — текст по-настоящему | OOXML §17.3/17.4/17.7, style cascade, font engine (TTF emboed Type0+CIDFontType2), greedy line breaking, inline-формат (bold/italic/underline/color/size) | Корпус ~50 деловых docx, визуальный diff с Word/LO < 5% per page |
| **M2** — структура | Таблицы (§17.4 + auto-layout), списки (§17.9), headers/footers, sections, изображения (PNG/JPEG) | Корпус ~300 документов, diff < 8% |
| **M3** — Excel | SpreadsheetML парсер, Excel number formats, сеточный layout, **print area ✅, page breaks ✅, print-model ✅** (gridlines/fit-to-page/print-titles/centered) | Корпус ~100 xlsx, выход совпадает с Excel print preview |
| **M4** — продвинутый текст | GSUB/GPOS (лигатуры, kerning), Knuth-Plass, Liang hyphenation (ru/en), BiDi, footnotes | RTL и продвинутая типографика на корпусе работают |
| **M5** — DrawingML, charts, math | **Shapes ✅ + Charts ✅ + OfficeMathML ✅** — полный from-scratch матнабор (дроби/индексы/радикалы/n-арные/функции/пределы/разделители/матрицы/акценты), inline + display | **M5 завершён полностью** |
| **M6** — compliance | **Tagged PDF ✅ + PDF/A-1b/1a/2b/2u/2a/3b/3u/3a ✅** — логическая структура (StructTreeRoot, marked content, Document/H1–H6/P/L/LI/LBody/Table/TR/TD/Figure), /Lang, alt-текст, артефакты; PDF/A-2/3 (PDF 1.7, прозрачность, вложенные файлы); **все 8 профилей формально проходят veraPDF 1.30**; **цифровые подписи ✅** (PKCS#7 detached, openssl-verified) | **M6 завершён** |

## Валидация (с первого дня)

**Corpus-driven разработка:**
1. ~1000 документов, каждый размечен по фиче ECMA-376
2. Эталонные рендеры: `soffice --convert-to pdf` как «золотой стандарт»
3. **Структурный diff**: текст с координатами из обоих PDF
4. **Визуальный diff**: PNG-рендер (через mupdf) + pixelmatch/PerceptualDiff
5. CI прогоняет весь корпус, бот публикует таблицу регрессий
6. Трекинг покрытия спеки: сколько элементов реализовано / частично / нет

## Текущий статус

### История диалога

1. Изначальная идея пользователя: скачать LibreOffice source, изучить, портировать в WASM.
2. Я предложил ZetaJS (готовый WASM-порт от Allotropia, MIT) → отвергнуто (тяжёлая загрузка).
3. Предложил серверный LibreOffice / Gotenberg → отвергнуто (дорого в контейнере).
4. Пользователь предложил собственный конвертер по сопоставлению спецификаций Word ↔ PDF.
5. Я попытался предложить bounded subset / DOCX→HTML→PDF как пути меньшего сопротивления → отвергнуто.
6. Зафиксировано: **строго по спецификациям**, **JS/TS**, **с нуля**.

### Tasks (на момент handoff)

| ID | Статус | Описание |
|---|---|---|
| #1–5 | completed | Research-фаза |
| #6–10 | completed | **M0** — bootstrap, OPC, text extractor, минимальный PDF writer, CLI, smoke |
| #11 | completed | M1.1: TTF parser (head/hhea/hmtx/maxp/cmap form4+12/name/post/OS2) |
| #12 | completed | M1.2: PDF Type0 + CIDFontType2 + FontFile2 + ToUnicode |
| #13 | completed | M1.3: WinAnsi → Identity-H/CID в renderer |
| #14 | completed | M1.4: Roboto Regular прокинут через converter, кириллица |
| #15 | completed | M1.5: TTF subsetting (closure composites + glyf zeroing + checksumAdjustment) |
| #16 | completed | M1.6: Renderer two-pass для subsetting |
| #17 | completed | M1.7: OOXML §17.4.1 rPr parser (b/i/u/sz/color/strike/rFonts/vertAlign) |
| #18 | completed | M1.8: OOXML §17.3.1 pPr parser (jc/spacing/ind/pStyle) |
| #19 | completed | M1.9: styles.xml parser + style cascade (docDefaults → para → char → direct, basedOn chain) |
| #20 | completed | M1.10: Document model + FontRegistry (regular/bold/italic/boldItalic с fallback) |
| #21 | completed | M1.11: Styled PDF emit — mixed fonts, цвет (rg), Tf size, alignment center/right |
| #22 | completed | M2.1: preserveOrder XML helpers (PO walker + po-to-flat адаптер для свойств) |
| #23 | completed | M2.2: Document model — BodyElement union, Table/Row/Cell + Border/CellMargins типы |
| #24 | completed | M2.3: OOXML §17.4 table parser (tbl/tblGrid/tr/tc + tblPr/trPr/tcPr, vMerge, gridSpan, borders, margins) |
| #25 | completed | M2.4: Layout таблиц — column widths + cell wrap + row heights |
| #26 | completed | M2.5: PDF emit таблиц — границы (m/l/S операторы) + позиционирование текста |
| #27 | completed | M2.6: insideH/insideV + position-aware border resolution + no-default-fallback |
| #28 | completed | M2.7: Dedup border drawing (top+left per cell, bottom только в последней строке, right в последней колонке) |
| #29 | completed | M2.8: Auto-fit column widths (max single-line content width per column) |
| #30 | completed | M2.9: Justify alignment ("both") — per-token Tm positioning, последняя строка left |
| #31 | completed | M2.10: gridSpan render — ячейка занимает сумму spanned-колонок |
| #32 | completed | M2.11: vMerge render — pre-process roles (start/middle/end), suppress content и borders на continuation-ячейках |
| #33 | pending | M2.12: Row split через page break для строк выше страницы (текущее: атомарный перенос строки целиком) |
| #34 | completed | M2.13: Document model — Numbering, NumberingLevel, NumberingReference |
| #35 | completed | M2.14: numbering.xml parser + numPr parsing + hanging indent (`<w:ind w:hanging>`) |
| #36 | completed | M2.15: NumberingState + marker rendering (decimal/lowerLetter/upperLetter/lowerRoman/upperRoman/bullet с Symbol-PUA → U+2022 substitution) |
| #37 | completed | M2.16: Document model — SectionProperties, PageSize, PageMargins, HeaderFooterReference |
| #38 | completed | M2.17: OPC.getPartRelationships + resolveRelatedPart + r:/xml: namespace support в po-helpers |
| #39 | completed | M2.18: parseSection (sectPr из body — pgSz/pgMar/header/footer refs) + parseHeaderFooter (w:hdr/w:ftr) |
| #40 | completed | M2.19: Renderer — pgSz/pgMar override defaults, layout header/footer один раз и emit на каждой странице |
| #41 | completed | M2.20: Document model — ImageBlock как третий kind в BodyElement |
| #42 | completed | M2.21: Drawing parser — wp:inline/anchor + wp:extent (EMU) + a:blip r:embed → ImageBlock |
| #43 | completed | M2.22: JPEG XObject через `/DCTDecode` (SOFn marker → width/height/components) |
| #44 | completed | M2.23: PNG decoder — inflate IDAT, undo фильтров None/Sub/Up/Average/Paeth + `/FlateDecode` XObject, RGBA/Gray+A → SMask |
| #45 | completed | M2.24: Renderer — image XObject embed, layout с EMU→pt + clamp по contentWidth, page /XObject resource, emit `q w 0 0 h x y cm /Im Do Q` |
| #46 | completed | M3.1: Cell reference helpers (A1 ↔ row/col) |
| #47 | completed | M3.2: workbook.xml + sharedStrings.xml parsers |
| #48 | completed | M3.3: worksheet.xml parser (rows × cells с типами n/s/str/b/d/e/inlineStr) |
| #49 | completed | M3.4: convertXlsxToPdf — построить Table из cells, переиспользует styled renderer + auto-fit |
| #50 | completed | M3.5: CLI dispatch по расширению (.xlsx → xlsx, иначе docx) |
| #51 | completed | M3.6: xl/styles.xml parser — fonts (sz/b/i/u/color/name), fills (solid fgColor), cellXfs, numFmts |
| #52 | completed | M3.7: Excel number formatter — built-in numFmtId 0-49 + custom (`#,##0.00`, `0%`, литералы в кавычках, `[colour]` отбрасываются) |
| #53 | completed | M3.8: WorksheetCell.styleIndex + apply font/alignment/numFmt per cell |
| #54 | completed | M3.9: Merged cells (`mergeCells/mergeCell ref`) → gridSpan + vMerge=restart на origin, continue на вертикальных |
| #55 | completed | M3.10: `pageBreakBefore` в ParagraphProperties + cascade + renderer форсит startNewPage |
| #56 | completed | M3.11: xlsx все sheets — sheet title + table на каждый, pageBreakBefore для sheetIdx>0 |
| #57 | completed | M3.12: `CellShading { colorHex }` + `CellProperties.shading?` в document-model |
| #58 | completed | M3.13: styles-parser `parseBorders` — top/right/bottom/left с style (thin/medium/thick/hair/dashed/dotted/double/...) + color |
| #59 | completed | M3.14: Renderer DrawCommand kind `'fill'` — отдельный q/Q блок с `re/f` ДО images/borders/text |
| #60 | completed | M3.15: xlsx `shadingFromXf` (fillId → solid fgColor → shading) + `bordersFromXf` (borderId → CellBorders с маппингом Excel styles) |
| #61 | completed | M3.16: Excel date format — `excelSerialToDate` (epoch 1899-12-30) + built-in IDs 14-22, 45-47 + custom format tokenizer (m/mmm/mmmm/d/dd/yyyy/h/hh/m минуты/s/AM-PM/[h]) с disambiguation `m`=month vs minute |
| #62 | completed | M3.17: xlsx row heights — `<row ht customHeight>` парсер (pt→twips) + TableRow.properties с heightTwips/heightRule='atLeast' |
| #63 | completed | M3.18: xlsx page setup — `<pageMargins>` (inches→twips) + `<pageSetup paperSize orientation>` (Letter/A3/A4/A5/Legal/Tabloid + landscape swap) → SectionProperties, проброшен в renderStyledPdf через options.section |
| #64 | completed | M3.19: xlsx date1904 mode — `parseWorkbook` теперь возвращает `{ sheets, date1904 }`, флаг проброшен в applyNumberFormat → excelSerialToDate с epoch 1904-01-01 |
| #65 | completed | M2.12: Row split — таблица строка выше pageContentHeight разделяется на чанки между страницами; first/last chunk сохраняют top/bottom борders и padding |
| #66 | completed | M2 leftover: First/even headers/footers — `<w:titlePg/>` в sectPr и `<w:evenAndOddHeaders/>` в settings.xml; renderer выбирает default/first/even по индексу страницы |
| #67 | completed | M2 leftover: Multiple sections — `parseSections` собирает sectPr из pPr + финальный body-level; renderer применяет per-section pgSz/pgMar/headers/footers/titlePg, MediaBox меняется по секциям |
| #68 | completed | M2 leftover: Inline images — `Run.inlineImage` несёт картинку прямо в текстовый run; tokenize выдаёт image-токены, emit делает ET/cm/Do/BT внутри строки |
| #69 | completed | M4.1: Knuth-Plass line breaking — total-fit DP с 4 fitness classes, tolerance ratio 4, заменил greedy `wrap()` в styled-page-renderer |
| #70 | completed | M4.2: Liang hyphenation — pattern-based с trie; bundled en-US (4938 patterns) и ru (7021 patterns) с lazy load; интегрирован в KP как flagged penalty items; hyphen suffix добавляется к последнему фрагменту строки |
| #71 | completed | M4.3: OpenType GSUB/GPOS — парсер ligature substitution (Type 4) и pair positioning (Format 1/2), feature filter `liga`/`clig`/`kern`; shapeText применяется в textWidthPt и encodeTextAsCidHex |
| #72 | completed | KP perf: cap active node list — max 200 узлов, сортировка по totalDemerits, отброс хвоста. Тесты бегут на 40% быстрее (~1s vs 1.7s) |
| #73 | completed | PDF /Info dictionary — Title/Author/Subject/Keywords/Producer/CreationDate/ModDate; UTF-16BE hex для Unicode; парсер docProps/core.xml для извлечения metadata из docx/xlsx, options.info для override |
| #74 | completed | M4.4: BiDi (UAX #9) — `src/bidi/` (char-types + P/X/W/N/I/L rules с isolate-support); модель `ParagraphProperties.bidi` / `RunProperties.rtl`; парсеры `w:bidi`/`w:rtl`; tokenizer режет по embedding level; emit реордерит визуально (L2) + реверс code points для RTL; авто-детект базы; RTL → right-align default |
| #75 | completed | PDF/A-1b — `options.pdfA`: PDF 1.4 header + детерминированный `/ID` (FNV-хэш тела) + OutputIntent с встроенным sRGB ICC v2 (`src/pdf/icc-profile.ts`, сгенерирован с нуля) + XMP `/Metadata` с pdfaid (`src/pdf/xmp.ts`) + subset-tag в FontName/BaseFont + flatten альфы PNG на белый (SMask запрещён). CLI `--pdfa`. 12 тестов |
| #76 | completed | M5.1: Vector-graphics emit layer — `src/pdf/vector-graphics.ts` (новый `'shape'` DrawCommand; PathSegment/VectorPath/StrokeStyle/VectorShape; `emitVectorShape` → `q cm … (B/f/S/n) Q`, dash/cap/join). Shape-проход в `emitPageContent` ПЕРЕД текстом (заливка под текстом, текст фигуры — поверх) |
| #77 | completed | M5.2: `src/pdf/arc-to-bezier.ts` — эллиптическая дуга → ≤90° κ-bezier (общий для ellipse/roundRect/custGeom arcTo); y-up знак сосредоточен здесь; guard нулевых радиусов |
| #78 | completed | M5.3: Shapes end-to-end — модель ShapeBlock + `BodyElement` `'shape'`; `drawing-parser.ts` (MCE `resolveMc`, wps:wsp, spPr/xfrm/prstGeom/solidFill/ln); preset-geometry rect/roundRect/ellipse; `layoutShapeBlock` + paginate + 4 visitor-ветки; `resolveColor` проброшен через parseDocument/parseHeaderFooter/parseBodyElements (дефолт — Office-палитра) |
| #79 | completed | M5.4: MCE robustness (Choice по Requires / Fallback / неизвестный ns→no-shape, вложенный AltContent) + `a:ln` dash (prstDash→pt-массив) / cap (rnd/sq/flat→J) / width |
| #80 | completed | M5.5: Preset breadth — triangle/rtTriangle/diamond/parallelogram/trapezoid/pentagon/hexagon/line/straightConnector1/4×arrow; неизвестный preset → bounding rect fallback |
| #81 | completed | M5.6: custGeom — `a:pathLst/a:path` moveTo/lnTo/cubicBezTo/quadBezTo(→cubic)/arcTo/close; `customPaths` scale path-space→pt + один y-flip |
| #82 | completed | M5.7: Text in shape — wps:txbx/w:txbxContent + bodyPr (insets 91440/45720, anchor t/ctr/b); layoutParagraphBlock в inset-rect, вертикальный якорь, эмит как `'line'` поверх заливки; шрифты текста фигуры в subset (`embedUsedFonts`) |
| #83 | completed | M5.8: Theme colours — `src/ooxml/drawingml/theme-parser.ts` (a:clrScheme, srgbClr/sysClr@lastClr); `src/ooxml/drawingml/colors.ts` (RawColor/ColorResolver, `makeColorResolver`, tx1/bg1→dk1/lt1 алиасы, дефолтная Office-2013-палитра). Конвертер грузит word/theme/theme1.xml через relationships |
| #84 | completed | M5.9: Polish — wp:anchor (размер из extent, трактуется как блок), oversized-clamp (фигура масштабируется под страницу, без зацикливания), shape в mixed-run отбрасывается (текст сохраняется), exhaustiveness-аудит всех `el.kind` |
| #85 | completed | M5.10: Charts model + parser + wiring — `Chart`/`ChartSeries`/`ChartBlock` + `BodyElement` `'chart'`; `src/ooxml/drawingml/chart-parser.ts` (chartSpace→plotArea→bar/line/pie `c:ser` с **кэшем** numCache/strCache, title, legend, barDir, grouping, цвет серии из fill или a:ln); drawing-parser детектит chart graphicData uri → `c:chart r:id`; конвертер `loadCharts` (resolve rel→chart1.xml) → `options.charts`; renderer резервирует бокс |
| #86 | completed | M5.11: Column/bar render — `src/pdf/chart-geometry.ts` (Heckbert `niceScale`, общий `buildFrame`: оси/гридлайны/тики/категории/легенда/заголовок) + кластерные столбцы (col) и горизонтальные (bar); рендер через vector-слой, подписи через text-pass; `embedUsedFonts` собирает глифы подписей |
| #87 | completed | M5.12: Line render — `buildLineScene` (полилиния на серию через общий frame); цвет серии из a:ln |
| #88 | completed | M5.13: Pie render — `buildPieScene` (секторы через arc-to-bezier, старт 12 ч по часовой, цвета слайсов dPt/accent-cycle, %-подписи, легенда по категориям) |
| #89 | completed | M5.14: Polish + demo — заголовок/легенда (r/b)/`formatTick`/accent-cycle/graceful unknown-fallback; `charts-demo.pdf`; handoff |
| #90 | completed | M5.15: OfficeMathML — модель `MathNode` (рекурсивная) + `Run.math`; `src/ooxml/math/omml-parser.ts` (m:oMath/oMathPara, m:r/m:t стили, m:f); `src/pdf/math-layout.ts` — box-model движок с нуля (MathBox width/ascent/descent + позиционированные glyph/rule/path items, TeX-подобные константы, авто-курсив переменных); fractions; интеграция: MathToken в строке (line-height учитывает ascent/descent), emit glyph+rule, embedUsedFonts собирает мат-глифы по variant |
| #91 | completed | M5.16: Scripts (m:sSup/sSub/sSubSup/sPre, 0.7× размер, raise/lower) + radicals (m:rad, нарисованный surd-путь с винкулумом, degree) |
| #92 | completed | M5.17: n-ary (m:nary, ∑∏∫ нарисованы векторными путями; limLoc undOvr/subSup; chr/subHide/supHide) + m:func + m:limLow/limUpp. Фикс `poAttr`: добавлен `@_m:` namespace (m:val для m:chr/m:type/m:sty/...) |
| #93 | completed | M5.18: delimiters (m:d, растяжимые скобки ()[]{}\|⟨⟩ нарисованы по высоте контента, begChr/endChr/sepChr) + matrices (m:m, выровненная сетка) + accents (m:acc — hat/bar/vec/dot/tilde нарисованы) + m:bar + m:groupChr (фигурная скоба) |
| #94 | completed | M5.19: display (m:oMathPara → центрирование параграфа, m:jc) + inter-atom spacing (medium space вокруг бинарных операторов/отношений) + `math-demo.pdf`; handoff |
| #95 | completed | M6.1: Tagged PDF фундамент — `src/pdf/struct-tree.ts` (`StructNode`/`StructTreeBuilder.emit`: StructTreeRoot + StructElem + ParentTree, pre-order DFS для детерминизма). `options.tagged`; `pdfA` расширен `'PDF/A-1a'` (→ tagged). `structId` на DrawCommand; параграф → `Document → P`. `emitPageContent` tagging-aware (per-line MCID, `/P <</MCID n>> BDC…EMC`, артефакт-обёртка fill/image/border/shape). Catalog `/MarkInfo`+`/StructTreeRoot`+`/Lang`; page `/StructParents`+`/Tabs /S`. XMP conformance `A`. **tagged=false ⇒ байт-в-байт без изменений** |
| #96 | completed | M6.2: Заголовки — `w:outlineLvl` (§17.3.1.20) в `paragraph-properties.ts` → `ParagraphProperties.outlineLevel`; резолв `outlineLevel`/`styleId` в `ResolvedParagraphProperties`; маппинг 0–8 → H1–H6 (clamp) + эвристика styleId `Heading[1-9]`/`Title`→H#. Marked-content тег = тип структуры (`tagFor`), не хардкод `/P` |
| #97 | completed | M6.3: Артефакты — явный флаг `artifact: 'pagination'` на командах header/footer (`markPagination`); line-проход типизирует их `/Artifact <</Type /Pagination>> BDC` (надёжный дискриминатор, не «нет structId») |
| #98 | completed | M6.4: Таблицы — `Table → TR (на строку) → TD (на логическую ячейку, пропуская vMerge continuation) → P`; per-cell `structId` в `emitRowChunk`; узлы TD/P переиспользуются между чанками строки (page split); colSpan → один TD |
| #99 | completed | M6.5: Списки — `L → LI → LBody → P` со вложением по `ilvl` (стек `ListFrame`; вложенный L внутри LBody родителя); `numbering` пробрasывается на `ParagraphBlock`; маркер внутри P (без Lbl); стек сбрасывается на не-list блоке |
| #100 | completed | M6.6: Figures + alt — `wp:docPr @descr/@title` в `drawing-parser.ts` → `altText` на ImageBlock/ShapeBlock/ChartBlock; пробрasывается на laid-out блоки; `createFigure` (Figure + непустой `/Alt`, fallback Image/Shape/Chart-title); `emitTaggedRuns` (per-run figure-маркировка image/shape, coalescing смежных команд одной фигуры — chart-формы под одним MCID) |
| #101 | completed | M6.7: `/Lang` — `options.language` → catalog `/Lang` (default en-US); конвертер авто-детектит `w:lang` (дешёвый regex по styles/document, `detectDocxLanguage`); `/ViewerPreferences <</DisplayDocTitle true>>` при наличии title |
| #102 | completed | M6.8: Conformance gate — `tests/tagged-pdf.test.ts` (21 тест): инвариант «нет painting-операторов вне marked content» на богатом документе (заголовок+параграф+таблица с заливкой/границами+вложенный список+картинка+header/footer), каждый Figure имеет `/Alt`, XMP conformance A; `tagged-demo.pdf` |
| #103 | completed | M6.9: **veraPDF 1.30 формальная валидация** — установлен (izpack console под Java 17), -1a и -1b проходят. Найдены+исправлены 2 реальных бага: (1) §6.3.5 `/CIDSet` отсутствовал в дескрипторе субсета (`glyphClosure`→`buildCidSet`); (2) §6.3.8 ToUnicode — `embedUsedFonts` теперь собирает **шейпленные** глифы (лигатуры) как emit (раньше лиг-глиф рендерился, но вырезался из субсета), ToUnicode маппит лиг-глифы → составляющие code points + скан с U+0009 (TAB-глиф маркеров). Opt-in gate `tests/verapdf.test.ts` + структурный `/CIDSet`-тест |
| #104 | completed | M6.10: **PDF/A-2** (a/b/u) — `parsePdfAProfile` (часть/уровень/tagged/версия); `options.pdfA` расширен до 8 значений; XMP `pdfaid:part` 1/2/3 + conformance A/B/U; PDF 1.7 для part 2/3. **Прозрачность**: part≥2 не флэттит альфу (SMask сохраняется), страницы получают transparency group `/Group <</S/Transparency /CS [/ICCBased sRGB]>>` (§6.2.4.3). `/CIDSet` гейтится на part 1 (для -2/-3 опционален и требует точного совпадения — проще не писать). veraPDF 2b/2u/2a PASS |
| #105 | completed | M6.11: **PDF/A-3 + associated files** — `src/pdf/embedded-file.ts` (`/EmbeddedFile` stream + `/Filespec` с `/AFRelationship`); `options.attachments` → catalog `/AF` + `/Names /EmbeddedFiles` (только part 3 или обычный PDF); конвертеры `embedSource` встраивают исходный docx/xlsx как `/Source`. Публичные типы `AttachedFile`/`AFRelationship`/`PdfALevel`. veraPDF 3b/3u/3a PASS с встроенным источником; `pdfa2-demo.pdf`/`pdfa3-demo.pdf` |
| #106 | completed | C1: **Защита от zip-bomb** в `OpcPackage.open` — слоистые капы (размер архива / размер entry / суммарный распакованный / число entry) через fflate-фильтр, который отсекает запись ДО распаковки; щедрые дефолты (`OpcOpenOptions`), переопределяемы. 5 тестов (`tests/opc-package.test.ts`) |
| #107 | completed | C2: **Песочница LibreOffice (Docker)** для golden-render untrusted-документов — `scripts/corpus/sandbox/Dockerfile` (послойный debian-slim + headless LO, non-root, образ `docgen-losandbox` 565 МБ); `sofficeToPdfSandboxed`/`referenceToPdf` (`docker run` с `--network none`, `--cap-drop ALL`, `--security-opt no-new-privileges`, `--read-only`+tmpfs uid 1000, лимиты cpu/mem/pids, вход ro). **Наш парсер тоже изолируется**: `convert-one.ts` запускается child-процессом с wall-clock-таймаутом и heap-капом (`CORPUS_SANDBOX=docker`). Прогнано end-to-end на POI-docx |
| #108 | completed | C3: **Fetch-корпус** — `scripts/corpus/fetch-corpus.ts` (`npm run corpus:fetch`) тянет lic-clean docx/xlsx из Apache POI test-data (Apache-2.0, скачано **127 docx + 349 xlsx** OOXML; .doc/.xls/.docm/.xlsm пропускаются) в `corpus/external/` (gitignored) + манифест провенанса. Раннер принимает `CORPUS_DIR` |
| #109 | completed | **Робастность изображений** (найдено POI-корпусом) — `embedUsedImages` оборачивает `embedImage` в try/catch, при ошибке картинка пропускается; image-проход рендерит только команды с валидным `imageResourceName`, `emitImageToken` гейтит на него же. Неподдерживаемая/битая картинка (Palette PNG, bit depth 4, interlaced, GIF, битый поток) больше **не валит весь документ** — пропускается без висячей `/Im… Do`-ссылки, остальной контент рендерится. 9 ранее падавших POI-docx теперь конвертируются (8 с ошибкой формата + `issue_51265_3`, дававший битый PDF из-за висячей ссылки). Регрессия в `tests/image.test.ts` (describe «image robustness») |
| #110 | completed | **SpreadsheetML prefix-agnostic** (найдено POI-корпусом) — некоторые продьюсеры (Haansoft HCell и др.) пишут xlsx с ЯВНЫМ namespace-префиксом `x:` (`<x:workbook>…<x:sheet r:id=…>`), а парсеры workbook/worksheet/sharedStrings/styles матчили теги по дефолтному (беспрефиксному) namespace → 0 листов → `xlsx has no sheets` (6 файлов корпуса). Фикс: `removeNSPrefix: true` в конфиге всех 4 spreadsheet-`XMLParser` (срезает префикс с тегов и атрибутов; `r:id`→`id` через фолбэк; беспрефиксный кейс не меняется; ECMA-376 допускает любой префикс). Регрессия в `tests/xlsx.test.ts` (`x:`-префикс на workbook/sst/worksheet) |
| #111 | completed | **Tracked changes (accept-all)** (найдено POI-корпусом) — `collectRuns` собирал только `w:r`, поэтому вставленные runs внутри `w:ins` (§17.13.5.18) выпадали (на `delins.docx` терялось 1368/1842 символов). Фикс: `w:ins` и `w:moveTo` добавлены в `RUN_CONTAINER_TAGS` → вставленный/перемещённый-в текст рендерится; `w:del`/`w:moveFrom` намеренно вне набора → удалённый/перемещённый-из опускается (текст в `w:delText`, не `w:t`) = семантика «принять все изменения» / финальный документ. `delins.docx` 474→1326 символов. Регрессия в `tests/document-parser.test.ts` |
| #112 | completed | **Forced page break `<w:br w:type="page">`** (найдено POI-корпусом, #37) — `parseRun` трактовал ЛЮБОЙ `<w:br>` как `\n`, поэтому жёсткий разрыв страницы (§17.3.3.1) рендерился как мягкий перенос → POI header/footer-сюита давала 1 страницу вместо 2. Фикс: `Run.pageBreak` (только `type="page"`; прочие типы остаются `\n`) → `ParagraphBlock.pageBreakAfter` → цикл рендера флашит страницу перед следующим блоком (переиспользует `flushPage`). `ThreeColHead/Foot/HeadFoot`, `PageSpecificHeadFoot`, `DiffFirstPageHeadFoot` и др. → 2 страницы. Регрессии в `tests/section.test.ts` (e2e) + `tests/document-parser.test.ts` (parser) |
| #113 | completed | **Header/footer-only документ** (найдено POI-корпусом, #36) — при пустом теле `paginateSections` не давал ни одной страницы (`flushPage` бейлит на пустом `current`), а `renderStyledPdf` форсил ПУСТУЮ fallback-страницу без колонтитулов → весь текст файла (живущий только в header/footer) терялся. Фикс: `paginateSections` форсит финальную страницу через header-путь (`flushPage(true)`), когда тело ничего не дало → колонтитулы рендерятся. `headerFooter.docx` 0→54 символа (=эталон), `EmptyDocumentWithHeaderFooter` 0→12 (=эталон), `Headers` 0→9. Нормальные документы (тело непустое) не затронуты. Регрессия в `tests/section.test.ts` |
| #114 | completed | **xlsx OOM guard** (тех-долг #41) — `CVLKRA-KYC` (204 КБ) рвал 512 МБ heap: `sheet15` объявляет `A1:XFD23`, 49 194 ячейки (48 со значениями, остальные пустые стилизованные до колонки XFD/16384) → `worksheetToBody` строил плотную сетку `maxColumn+1`. Фикс (`xlsx-to-pdf.ts`): сетка ограничена «used range» (содержательные ячейки + merges), пустые вне него отбрасываются; + бэкстоп-кап `MAX_GRID_COLS=1024`/`MAX_GRID_ROWS=50000` и клэмп merge-расширения для untrusted. Конвертится (74 стр.). Регрессия в `tests/xlsx.test.ts` (raw-xlsx стрэй-XFD → байт-в-байт с без-стрэя) |
| #115 | completed | **xlsx sharedStrings DoS guard** (тех-долг #42) — `poc-shared-strings.xlsx`: 1 общая строка 1 048 583 символа × 12 000 ячеек → шейпинг ~12 ГБ → хэнг >60с. Два библиотечных self-limit'а: (1) парсеры капят строку ячейки до Excel-лимита 32 767 (shared-strings + worksheet inlineStr); (2) конвертер ограничивает суммарный текст листа `MAX_SHEET_TEXT_CHARS=1 000 000` (`textBudget`). Конвертится ~2с (76 стр.). Регрессия в `tests/xlsx.test.ts` |
| #116 | completed | **Вложенные таблицы (table-in-cell)** (найдено POI-корпусом, #44) — модель и парсер уже рекурсили (`TableCell.content: BodyElement[]`, `parseTableCell`→`parseBodyElements`), но рендерер пропускал не-параграфы в ячейке (`if (el.kind !== 'paragraph') continue`). Фикс (`styled-page-renderer.ts`): `layoutTableCell` раскладывает вложенный `w:tbl` через `layoutTableBlock` по inner-width ячейки (`CellLayout.nestedTables`); `emitRowChunk` рендерит их под lines рекурсивным вызовом самого себя (инсет в content-box ячейки, под structId родителя при tagged); `TableBlock` получил `colCount`; `splitRowIntoChunks` не сплитит строку с вложенной таблицей. `60329.docx` 0→4607 символов (эталон 4812 ≈96%). Ячейки без вложенных таблиц — байт-в-байт. Регрессия в `tests/styled-render.test.ts`. Ограничения: параграф после вложенной таблицы в той же ячейке рендерится после неё (lines-then-tables); вложенные таблицы не пагинируются между страницами |
| #117 | completed | **Per-run резолв шрифтов** (мульти-семейный) — раньше весь документ рендерился ОДНИМ семейством (`options.registry`); теперь каждый текстовый ран резолвит своё семейство по `w:ascii` (`resolveFamilyKey`: sans→roboto / serif→tinos / mono→cousine). Новое `StyledRenderOptions.registriesByFamily` (`Map<FamilyKey, FontRegistry>`); ключ font-resource стал композитным `family:variant` (helper `runFontKeyAndParsed` + устойчивый `lookupFont`). Async `convertDocxToPdf` детектит distinct-семейства (`detectDocxFamilyKeys`) и тянет набор на каждое; pure-sans доки идут прежним one-family путём. **Single-family — байт-в-байт** (поведение гейтится наличием `registriesByFamily`; 334 прежних теста без изменений). Проверено: mixed sans+serif+mono → встроены Roboto+Tinos+Cousine, каждый ран в своём. Math/chart/fallback — на базовом `registry`. Регрессия в `tests/styled-render.test.ts` |
| #118 | completed | **Встроенные шрифты (`w:embed`)** — документ может зашивать свои шрифты (`word/fonts/*.odttf`, обфусцированные); теперь они де-обфусцируются (§17.8.1: XOR первых 32 байт reversed-GUID-ключом) и используются **напрямую** → glyph-exact, без подмены вообще (и мы, и LibreOffice берут ОДИН реальный шрифт документа). Новый модуль `src/ooxml/wordproc/font-table.ts` (`parseFontTable` + `deobfuscateEmbeddedFont` + `loadEmbeddedFonts`→`Map<name, FontRegistry>`); `StyledRenderOptions.embeddedFonts` (высший приоритет в `runFontKeyAndParsed` по `w:ascii`); sync-конвертер грузит из `pkg` (без сети — шрифты в файле). Проверено: `saut_page.docx` рендерится встроенным **ArialMT**, не Roboto. Гейтится наличием встроенных шрифтов (нет → байт-в-байт). Тесты в `tests/font-table.test.ts` (де-обфускация self-inverse + восстановление sfnt + parse) |
| #45a | completed | **xlsx print-model — слой 1 (gridlines + print area + парсинг модели)** — главный рычаг по 303 ⚠️ (TextSim ≥95% был лишь у 4/349). (1) **Подавление gridlines** (§18.3.1.70): Excel/Calc по умолчанию НЕ печатают сетку — раньше мы всегда клали полный тонкий grid (`insideH/insideV` + рамка) на КАЖДЫЙ лист → системное расхождение с golden. Теперь синтетический grid рисуется **только** при `<printOptions gridLines="1">`; по умолчанию рисуются лишь рамки из стилей ячеек. (2) **Print area** (§18.2.5): `_xlnm.Print_Area` из `<definedName localSheetId>` клиппит рендер в заданный диапазон (со стартовым смещением — не обязан начинаться с A1); пересекается с used-range (защита от over-declare). (3) **Парсинг полной print-модели**: `pageSetup` scale/fitToWidth/fitToHeight, `<sheetPr><pageSetUpPr fitToPage>`, `<printOptions>`, `<rowBreaks>/<colBreaks>`, workbook `<definedNames>`. Новый `src/ooxml/spreadsheet/defined-name-ref.ts` (`parseAreaRef`: `Sheet1!$A$1:$D$20` / кавычки / multi-area → bounding box). `worksheetToBody` рефакторён на оконные индексы `rowStart/colStart`. **Проверено на локальном POI-корпусе:** 333/351 конвертируются (18 «провалов» = фаззер/zip-bomb/битый zip, отвергаются в OPC-слое — НЕ регрессии); reach: 323 файла теперь без синтетической сетки (300 без printOptions + 23 с `gridLines="0"`, которые мы раньше игнорировали), 10 с сеткой, 8 print area, 38 scale, 7 fitToPage. +9 тестов в `tests/xlsx.test.ts`. Гейтится: нет print area / `gridLines` не задан → grid просто не рисуется (раньше рисовался) |
| #45b | completed | **xlsx print-model — fit-to-page масштаб.** `<pageSetup scale="N">` и `<pageSetUpPr fitToPage>`+`fitToWidth` → единый shrink-фактор на шрифты + высоты строк (в `auto`-layout: уменьшенный текст авто-fit пакует без агрессивного переноса, который был на полном кегле). **Shrink-only** (увеличение ломает авто-fit и Excel fit-to-page не увеличивает), флор 10% (минимум Excel). `computePrintScale` (fitToPage→`min(1, contentW·fitW/Σcol)`, иначе `scale/100`) + `sheetContentWidthTwips` + `scaleRunFont` (дефолт 11pt=22 half-pt); дефолт ширины колонки 1440→**960** твипс (реальный Excel default 8.43 chars). Гейтится `scaled = printScale < 0.999` → нескейленные листы **байт-в-байт** (тест `scale=100` = no-op). Проверено на POI: 0 регрессий, 15 листов с реальным scale/fitToPage конвертятся. +3 теста. fitToHeight (height-fit) отложен — нужен точный суммарный height (у нас content-driven) |
| #45c | completed | **xlsx print-model — print titles / повтор header-строк (+ docx `w:tblHeader` бонусом).** Рендерер раньше игнорировал `RowProperties.isHeader` — теперь **повторяет лидирующие header-строки** наверху каждой страницы, на которую переносится таблица. В `RowLayout` добавлен `isHeader` (из `TableRow.properties.isHeader`); в пагинационном цикле таблицы при mid-table flush'е заново эмитятся header-строки (как **artifacts**, без structIds → не дублируют дерево тегов), с гардом «строка влезает под повторённой шапкой». Повторяется только **максимальный лидирующий префикс** header-строк. Для xlsx: `_xlnm.Print_Titles` → `parseTitleRowRange` (только row-диапазон `$1:$2`; колоночные титулы — для гориз. пагинации, которой у нас нет) → строки в диапазоне помечаются `isHeader`. **Cross-cutting фикс:** docx-таблицы с `w:tblHeader` (парсер уже ставил `isHeader`, но рендерер не повторял) теперь тоже повторяют шапку. Гейтится: нет header-строк ИЛИ таблица влезает на страницу → **байт-в-байт** (контрольный тест). Проверено POI: 7 xlsx с Print_Titles, 0 регрессий (docx 110 OK / 17 fail — все OPC/parse/security, не рендерер). +6 тестов (`parseTitleRowRange`, повтор/неповтор) |
| #137 | completed | **docx таблицы — 2 бага на реальном договоре (`Шаблон_договора_подряда`).** (1) **Cell shading не парсился:** `table-parser.parseCellProperties` читал tcW/gridSpan/vMerge/borders/margins, но НЕ `w:shd` → `CellProperties.shading` никогда не ставился для docx → cyan-шапка (`<w:shd w:fill="00FFFF">`) рендерилась белой (рендерер-то fill умел с #57). Фикс: парс `w:shd@w:fill` (direct hex; `auto`/theme — без заливки). (2) **Длинный текст в узкой ячейке наезжал на соседний столбец:** в fixed-layout таблице (gridCol 1790/4911/…) vMerge-ячейка «На устройство котлована, разработка грунта (7 чел)» в колонке 90pt рендерилась ОДНОЙ 180pt-строкой поверх col2. Причина — **Knuth-Plass**: в узкой колонке 2-словная строка требует stretch-ratio ~8 > `TOLERANCE_RATIO` → все брейки «слишком рыхлые», пропускались, нода умирала overfull → emergency-фолбэк лепил одну overfull-строку. Фикс: too-loose брейк (`r > tolerance`) теперь **кандидат последней надежды** (badness ∝ r³, никогда не выигрывает при наличии feasible-брейка), а не отбрасывается → узкая колонка переносит рыхло, а не overfull. Overfull (`r < -1`) по-прежнему infeasible. Гейт: для нормального текста (есть feasible-брейк) — **байт-в-байт** (405 тестов целы, veraPDF целы); left-align не растягивается (`computeJustifyExtra` гейтит на `both`). Проверено визуально (rasterize стр.26: cyan-шапка + col1 переносится в свою колонку). +2 теста (KP narrow-wrap; docx `w:shd` parse). **(3) Пропали внутренние gridlines таблицы:** у таблицы НЕТ `<w:tblBorders>`, а каждая ячейка задаёт разделители на своих `right`/`bottom` (напр. col2: `right sz=2`, но БЕЗ `left`). Наш dedup (#28) рисует внутренние вертикали с `left`-стороны правой ячейки и `right` только у последней колонки → разделитель, заданный на `right`, не рисовал НИКТО. Фикс (`resolveCellBorders`): общее ребро резолвится через **§17.4 border-conflict resolution** (`heavierBorder`) — не `??`, а **побеждает более тяжёлая граница** (по `sizeEighthPt`, ties по style-rank double>thick>single>dashed>dotted): `left = isFirstCol ? own.left : heavierBorder(own.left??insideV, leftNeighbor.right??insideV)`, аналогично `top`/aboveNeighbor.bottom. Проброс соседей через `layoutTableCell`/`layoutTableRow` (`aboveBordersByCol` по колонкам). Гейт: таблицы с `tblBorders`/одно-сторонними бордюрами — без изменений (409 тестов целы). +2 теста (сепаратор на `right` рисуется; heavier-wins 2pt vs 0.5pt). Визуально: полная сетка как в Word. **Остаётся gap:** table styles (`tblStyle` + conditional formatting первой строки/банд) — отдельная фича, не conflict-резолв |
| #136 | completed | **Инфра — CI workflow (#66).** Новый `.github/workflows/ci.yml`: (1) **gates** job (на каждый push/PR): `npm ci` → typecheck → lint → format:check → build → test (node 20, ubuntu, npm-кэш). (2) **corpus** job (opt-in, `workflow_dispatch`): ставит LibreOffice + mupdf-tools, строит синтетический корпус (`corpus:build`), гоняет харнесс (`corpus`) vs LO-golden, аплоадит `corpus/report.md` артефактом. **Security:** CI-прогон только по ДОВЕРЕННЫМ синтетическим фикстурам; fetched POI-корпус (untrusted, содержит бомбы) — через Docker-sandbox (`CORPUS_SANDBOX=docker`), задокументировано в комментарии job'а (для self-hosted/hardened runner). **Отложено:** GovDocs1 scaling (большой публичный корпус — отдельный fetch + хранилище), font-agnostic visual-метрика (TextSim уже font-agnostic; добавочная geometry-метрика — research) |
| #135 | completed | **Типографика — BiDi N0 paired-bracket resolution (#65).** UAX #9 rule N0 (ранее пропускалось): парные скобки `()[]{}…` резолвятся по СОДЕРЖИМОМУ, а не только по соседям (N1), чтобы зеркалиться корректно в RTL. `resolvePairedBrackets` (BD16 стек-пэйринг, глубина ≤63, канон. эквивалентность 2329/232A≡3008/3009): strong e внутри пары → e; иначе strong o внутри + предыдущий контекст o → o; иначе нейтрально (N1/N2). Таблица канонических Bidi_Paired_Bracket (ASCII + репрезентативный Unicode-набор). Вызывается перед N1/N2; codePoints проброшены в `resolveNeutralTypes`. +2 теста («a(bא)» RTL → скобки R/level 1 по содержимому vs «a(b)» → L/level 2). Гейт: нет скобок → без изменений (17 прежних BiDi-тестов целы). 405 тестов. **Отложено:** NSM-after-bracket уточнение (нужны original-types до W1), bundled RTL-шрифт (упаковочное решение против minimal-deps — RTL покрывается встроенными шрифтами #118 / шрифтом вызывающего) |
| #134 | completed | **Типографика — Arabic cursive joining (#64).** Арабские буквы меняют форму по позиции (isolated/initial/medial/final). Новый `src/font/arabic-joining.ts`: `arabicJoiningType(cp)` (R/L/D/C/U/T по Unicode §9.2; основной блок U+0600–06FF + tatweel/ZWJ/ZWNJ) + `assignArabicForms(cps)` (алгоритм: ближайшие non-transparent соседи → joinsPrev/joinsNext → форма). GSUB **Type 1 single-substitution** парсер (`parseGsubArabicForms` → `init`/`medi`/`fina` gid→gid карты; Format 1 delta + Format 2 array). `shapeText` принял `joiningForms` (опц.): между char→gid и лигатурами подменяет глиф по форме. Проброшено через `ParsedTtf.joiningForms` + 3 call-site (cid-font ×2, renderer). **Работает с любым арабским шрифтом** (встроенным #118 или базовым, если есть init/medi/fina). Гейт: пустые карты (не-арабский шрифт) → output без изменений (Latin/BiDi тесты не затронуты). +7 тестов. 403 теста. **Отложено:** mark positioning (GPOS Type 4 mark-to-base/mark), chaining contextual (GSUB/GPOS Type 5/6), L-type leftjoining |
| #133 | completed | **PDF — JPEG 2000 (`/JPXDecode` pass-through) (#63).** JP2-картинки теперь встраиваются как `/JPXDecode` XObject (как JPEG через `/DCTDecode` — PDF-ридер сам декодирует wavelet-codestream, мы только читаем размеры; БЕЗ wavelet-декодера). `detectImageFormat` распознаёт JP2 box-сигнатуру (`00 00 00 0C 6A 50 20 20 0D 0A 87 0A`) и raw codestream (`FF 4F FF 51`). `readJpeg2000Info`: размеры из `jp2h`→`ihdr` (HEIGHT/WIDTH u32) или из SIZ-маркера (Xsiz/Ysiz − offsets) для raw/fallback. Раньше JP2 → unsupported → пропуск (#109). +2 теста (детект + embed 128×64→/JPXDecode). 396 тестов. **NB:** `/JPXDecode` разрешён в PDF/A-2/3, но НЕ PDF/A-1. **Отложено:** object streams / compressed xref (инвазивный рефактор writer'а; запрещены в PDF/A-1; риск байт-в-байт гейтов — низкая отдача) |
| #132 | completed | **Подписи — ECDSA (#62).** Подпись поддерживает ECDSA (P-256/384/521 + SHA-256) рядом с RSA PKCS#1 v1.5. `CmsParams.signatureAlgorithm: 'rsa'|'ecdsa'` → SignerInfo signatureAlgorithm = `ecdsa-with-SHA256` (OID 1.2.840.10045.4.3.2, ABSENT params) вместо rsaEncryption; digest остаётся SHA-256. `SignerCredentials.algorithm:'ecdsa'` → `signPdf` подписывает через WebCrypto ECDSA, конвертит raw r‖s (P1363) → DER Ecdsa-Sig-Value `SEQ{INTEGER r, INTEGER s}` (`ecdsaRawToDer`, переиспользует `der.integer` minimal-encoding). +1 тест (P-256 round-trip: CMS → извлечь sig-value → raw → WebCrypto verify против pubkey = OK; ecdsa OID присутствует). Гейт: дефолт 'rsa' → байт-в-байт. 394 теста. **Отложено:** PAdES (signing-certificate-v2/ESS + SubFilter ETSI.CAdES.detached), RFC-3161 timestamp (нужен TSA = сеть, вне scope no-network-либы), подпись внешних PDF через incremental-update (нужен парсинг+аппенд xref внешнего PDF) |
| #131 | completed | **xlsx print-model — fitToHeight (#61).** `computePrintScale` теперь учитывает высоту: при `fitToPage` + явном `fitToHeight≥1` shrink-фактор = min(width-fit, height-fit), где height-fit = `contentHeightTwips·fitH / Σ(row heights)`. Σrow = сумма rendered-row высот (custom `<row ht>`, иначе Excel-дефолт ~15pt=300 twips) — для xlsx высоты строк известны заранее (в отличие от docx content-driven), поэтому оценка tractable. `sheetContentHeightTwips` (page H − top/bottom margins). **Гейт:** height ограничивается ТОЛЬКО при явном `fitToHeight≥1` → width-only «fit all columns» (fitToHeight=0/absent) и не-fitToPage — байт-в-байт. +1 тест (120-строчный лист → меньше страниц). 393 теста. **Отложено:** `verticalCentered` (нужна суммарная rendered-высота ДО пагинации + офсет — неудобно с top-down потоком; редкий P4) |
| #130 | completed | **Tagged PDF — per-element `/Lang` (#60).** §14.9.2: абзац на языке ≠ документного дефолта теперь получает свой `/Lang` на StructElem (AT переключает произношение). `RunProperties.lang` (парс `w:lang@w:val`, §17.3.2.20) → `ResolvedRunProperties.lang` (run-каскад) → токены несут язык; `dominantParagraphLang` (взвешен по символам) на блок; в `paginateSections` P/H-узел получает `.lang`, если доминирующий язык ≠ дефолт (`options.language`/en-US). Инфра `StructNode.lang`→`/Lang` уже была. Гейт: non-tagged ИЛИ язык=дефолт → байт-в-байт (англ. тесты не затронуты). +1 тест (ru-RU абзац → `/Lang (ru-RU)`, англ. — нет). veraPDF Level A проходит. 392 теста. **Отложено:** `Lbl` для маркеров списков — нужен intra-line MCID-split (маркер `${marker}\t` в одной строке с текстом; structId — per-DrawCommand, не per-token; токен-уровневый marked-content = инвазивный рефактор emit), per-run Span `/Lang` (та же причина) |
| #129 | completed | **Shapes — colour transforms (lumMod/lumOff/shade/tint) + gradFill (#59).** (1) **Colour modulation** (§20.1.2.3): `RawColor` расширен `mods?: ColorMod[]`; резолвер применяет `applyColorMods` — shade·val (к чёрному), tint·val+(1−val) (к белому), lumMod/lumOff на HSL-яркости (RGB↔HSL хелперы). Раньше игнорировалось → темы «Accent N, Lighter X%» (lumMod 60000 + lumOff 40000) рендерились ПОЛНЫМ тёмным акцентом вместо светлого. Парсеры (drawing + chart) читают transform-детей `a:srgbClr`/`a:schemeClr`. (2) **gradFill** → solid-аппроксимация средним цветом стопов (`a:gsLst/a:gs`); PDF axial/radial shadings не эмитятся. +5 тестов. 391 тест. **Отложено:** group shapes (`wpg`/`a:grpSp` chOff/chExt — рекурсия трансформов), text rotation (нужна ротация text-emit — та же инфра-проблема, что у chart axis titles) |
| #128 | completed | **Math — equation arrays `m:eqArr` (#58).** §22 OfficeMathML: парсер не знал `m:eqArr` (выровненные системы уравнений). Новый `MathEqArray` (`type:'eqArr'`, rows) в модели+barrel; `parseEqArray` (каждый `m:e`→строка-уравнение); `layoutEqArray` стекует строки вертикально flush-left, блок центрируется на math-оси (как 1-колоночная матрица, но прижата влево, rowGap 0.5·size); кейсы в `layoutMath` dispatch + `mathGlyphSegments` (сбор глифов для субсета). +2 теста. 386 тестов. **Отложено (milestone-размер):** матшрифт STIX/Cambria — требует OpenType **MATH table** (glyph variants / stretchy construction / italic correction) + bundling OFL-шрифта; структурные символы (∑∏∫√, скобки) уже рисуются вектором, обычные — из base-шрифта. text/display sizing — частично (display центрируется, n-ary limLoc subSup/undOvr) |
| #127 | completed | **Charts — polish: data labels, axis titles, value-axis auto-min (#57).** (1) **Data labels** (`c:dLbls/c:showVal` — group- или series-level): `Chart.showValues` → значения печатаются на барах (центр сегмента/над столбцом), линиях (над точкой); `fmtDataLabel` (int как есть, иначе ≤2 знака). (2) **Axis titles** (`c:catAx/c:title`, `c:valAx/c:title`): `Chart.catAxisTitle`/`valAxisTitle` → cat-заголовок по центру снизу, val-заголовок горизонтально над плотом (без ротации — chart-текст не вращается; осознанное упрощение). `collectAT` хелпер для a:t. (3) **Value-axis auto-min** (`niceScale`): line/scatter больше НЕ форсят 0 (данные далеко от нуля → ось стартует у min); bar/area по-прежнему включают 0 (нужна осмысленная база). chart-title text уже эмитился (label→text pass). +4 теста. 384 теста. **colorN/styleN** (colors1.xml/style1.xml цветовые карты) — отложено (отдельные парты; series-цвета из `spPr` уже работают — низкая отдача) |
| #126 | completed | **Charts — новые типы: stacked/percentStacked, area, scatter, doughnut (#56).** Парсер уже маппил area/scatter/doughnut-теги в типы; работа была в геометрии/рендере. (1) **Stacked/percentStacked bar/col:** `buildBarScene` детектит `grouping` — stacked складывает сегменты по категории (cumulative pos/neg), percentStacked нормирует на сумму категории и пинит value-ось 0..100%; value-ось рескейлится на суммы (`stackedTotals`/`groupingFrameOpts`); `buildFrame` принял `FrameOpts` (dataRange + formatValue для %). (2) **Area** (`buildAreaScene`): залитые полосы от baseline (standard, back-to-front) или стопкой (stacked/percent); новый `ChartPolygon`+`polygons` в сцене, `polygonPrim` в рендере (нижний слой z-order). (3) **Scatter** (`buildScatterScene`): парсер читает `c:xVal`/`c:yVal`→`ChartSeries.xValues`; собственный фрейм с 2 числовыми осями (X+Y niceScale, гридлайны, маркеры). (4) **Doughnut:** `Chart.doughnut` (из `c:doughnutChart`); `buildPieScene` пробивает белый центральный диск (full-circle wedge r·0.5) + подписи на кольцо. Dispatcher: area→area, scatter→scatter. +16 тестов. 380 тестов. Прочие типы (radar/3D/bubble) — по-прежнему 'unknown' (рамка-плейсхолдер) |
| #125 | completed | **xlsx — number-format parity: `;@` text-section leak + scientific notation (#55).** (1) **Date text-section leak:** форматы `mmm-yy;@` / `m/d/yyyy;@` / `[$-409]mmmm\ d\,\ yyyy;@` — это даты с `;`-секцией для текста (`@`). `isDateFormat` верно детектил дату, но `formatExcelDate` токенизировал ВЕСЬ код включая `;@` → литералы `;@` подтекали после каждой даты. Фикс: `formatExcelDate` берёт только первую секцию (`splitSections(format)[0]`). **15/349 файлов** затронуто; на `54288` `;@` →0. (2) **Scientific notation** (numFmtId 11 `0.00E+00` + custom `##0.0E+0` engineering / `0.0e-0`): `applyNumericSection` раньше ронял `E`/знак (мусор) — теперь `formatScientific`: нормализация мантиссы на `intDigits` значащих (1 для `0.00E+00`, 3 для engineering → экспонента кратна 3), zero-pad экспоненты, знак (`E+` всегда / `E-` при <0), регистр `e`/`E` сохранён, renormalize при carry (999→`1.0e+03`). **Замер-вывод:** `47813` (`0.000`) уже был **char-identical** Calc — explicit-decimal форматтер корректен; xlsx TextSim-tail оказался НЕ number-форматы (54764 = XML entity-бомба, не репрезентативна; реальный остаток = column-major reading-order, не «неверный текст»). Дроби `?/?` (0 файлов) / thousands-comma (3, fiddly) — отложено. +8 тестов. Не-`;@`/`E` числа и не-даты — байт-в-байт |
| #124 | completed | **xlsx — клиппинг текста ячейки по cell-overflow модели (TextSim).** Excel/Calc: текст non-wrap ячейки **переливается** в ПУСТЫЕ соседние ячейки справа (left/general align), но **обрезается** там, где сосед занят (рендерится только влезающее, остальное отбрасывается). Мы рендерили полностью. Фикс (`xlsx-to-pdf.ts`): для string-ячеек (left/general, не merge) считаем доступную ширину (своя колонка + последовательные пустые справа); если переливание блокирует занятая ячейка — обрезаем текст до `availTwips/TWIPS_PER_EXCEL_CHAR` символов (`cellHasContent` helper). **`xlsx-jdbc` 66%→95.0% TextSim** (на пороге ✅). Достаёт genuine-overflow подмножество; прочие «ours>>ref» (54764 75/13) — иная причина (number-форматы, → #55). +1 тест (clip при занятом соседе / overflow при пустом). Числа (####), right/center, merge — не трогаются |
| #123 | completed | **xlsx — implicit `r=` позиции (§18.3.1.4), correctness-баг.** `r=` на `<row>`/`<c>` опционален — позиция тогда по порядку (след. строка после предыдущей; след. колонка после предыдущей ячейки). Парсер ТРЕБОВАЛ `r=` (`parseCell` возвращал null без него) → файлы с implicit-позициями рендерились **полностью пустыми** (worst failure mode). Фикс (`worksheet-parser.ts`): running `currentRow` (explicit `r=` сбрасывает, иначе +1) + running `prevCol` в строке (explicit `r=` сбрасывает, иначе +1); `parseCell(c, fallbackRow, fallbackCol)`. **`56278.xlsx`: 0 → 65 ячеек** (был 0/5457 chars). +2 теста (full-implicit + mixed explicit/implicit). Найдено корпус-диагностикой остатка ⚠️ |
| #122 | completed | **xlsx — дроп синтетического заголовка листа (#1 TextSim-фикс).** Диагностика корпуса (stext our vs Calc на мелких ⚠️-файлах) вскрыла: главная причина низкого TextSim — наш bold-заголовок с именем листа, которого Calc/Excel `--convert-to pdf` НЕ печатают НИГДЕ (прежняя гипотеза «Calc печатает имя в шапке» — неверна). Убрали заголовок; листы >1 ломают страницу **пустым no-text параграфом** (`{runs:[]}` + pageBreakBefore → 0 глифов). **Корпус: ✅ 19→163 · ⚠️ 309→165** (145 файлов ⚠️→✅, 1 регрессия; xlsx 5%→47% clean). Тест multi-sheet обновлён (проверяет данные листов + ОТСУТСТВИЕ имён). 364 теста |
| #121 | completed | **Tagged PDF — `/ColSpan` + `/RowSpan` для span-ячеек (§14.8.5.2).** Span-ячейки таблицы теперь несут `/A <</O /Table /ColSpan N /RowSpan N>>`: ColSpan из `CellLayout.colSpan` (gridSpan), RowSpan вычисляется проходом вниз по vMerge-продолжениям (`tableCellRowSpan` — считает middle/end до конца группы). `StructNode.colSpan/rowSpan` (new) + обобщённый `/A`-emit (Scope+ColSpan+RowSpan в одном attr-объекте). Закрыл ограничение (d). +2 теста (gridSpan→/ColSpan 2; vMerge→/RowSpan 2). veraPDF Level A проходит. Не-span ячейки без `/A` |
| #120 | completed | **Tagged PDF — `TH` для заголовочных строк таблиц (§14.8.5.2).** Раньше КАЖДАЯ ячейка была `TD`; теперь ячейки header-строки (`RowProperties.isHeader` — детект уже был с #47) эмитятся как `TH` с атрибутом `/A <</O /Table /Scope /Column>>` (AT привязывает их к данным под ними). `StructNode.scope` (new) → emit. Закрыл tagged-PDF ограничение (b). +1 тест (`tblHeader` → 2×TH+/Scope, 2×TD). **veraPDF 1a/2a/3a (Level A) по-прежнему проходят** — атрибут конформный. Не-header таблицы байт-в-байт |
| #119 | completed | **Цифровые подписи (ISO 32000 §12.8) — последняя фича M6.** PKCS#7/CMS detached (`adbe.pkcs7.detached`) с нуля, без крипто-библиотек (только WebCrypto — платформенный API). Новый слой `src/crypto/`: `asn1.ts` (DER-энкодер: seq/set/oid/integer/octetString/cmsTime/setOfBody + минимальный reader `readTlv`/`children`/`certIssuerAndSerial`), `cms.ts` (`buildPkcs7Detached`: SignedData v1, signedAttrs=contentType/messageDigest/signingTime, SHA-256, sha256WithRSAEncryption, issuerAndSerialNumber из cert). PDF-слой: `PdfRawToken` (verbatim-токен для fixed-width `/ByteRange`), `src/pdf/signature.ts` — `addSignaturePlaceholder` (невидимое `/Sig`-поле + `/AcroForm SigFlags 3` + sig-dict с placeholder ByteRange/Contents) и async `signPdf` (находит дыру, считает ByteRange, SHA-256 покрытых байт, строит CMS, вписывает в `/Contents` — остальные байты не сдвигаются). Рендерер: `StyledRenderOptions.signaturePlaceholder` (мутирует `/Annots` первой страницы + catalog `/AcroForm` после сборки). Конвертеры: async `convert*` принимают `options.signature` (= placeholder + credentials) → emit placeholder + `signPdf`. **tsconfig lib += WebWorker** (типы Crypto/SubtleCrypto/CryptoKey). **Проверено внешне:** `openssl cms -verify` → **«CMS Verification successful»**; +6 тестов (CMS round-trip через WebCrypto verify, e2e: ByteRange покрывает ровно нужные байты = messageDigest, детерминизм RSA, reserve-overflow throw, convertDocxToPdf e2e). Гейтится: нет `signaturePlaceholder` → байт-в-байт. Фикстура `tests/fixtures/sign/` (self-signed RSA-2048). **Не реализовано:** ECDSA/PAdES/timestamp(RFC 3161), подпись внешних PDF (incremental update) — у нас integrated-placeholder |
| #45d | completed | **xlsx print-model — ручные page breaks + centered.** (1) **Manual `<rowBreaks>`:** `RowProperties.pageBreakBefore` (new) → `RowLayout.breakBefore` → в пагинации таблицы форсит flushPage перед строкой (даже если влезает), затем повторяет header-строки. brk `id` трактуется как 0-based первая строка новой страницы → break перед `absR===id`. (2) **`<printOptions horizontalCentered>`:** `TableProperties.alignment` (new: left/center/right) → `TableBlock.xOffsetPt = tableXOffset(contentW − tableW)` → таблица эмитится со сдвигом `marginLeft+xOffset` (узкая таблица центрируется/прижимается). Gated: нет breakBefore И alignment=left/absent → `xOffsetPt=0` → **байт-в-байт**. Проверено POI: 1 rowBreaks + 13 horizontalCentered файла, 0 регрессий (docx 110 OK/17 fail — все OPC/parse). +2 теста (page-count при break, сдвиг fill-rect при center). **fitToHeight** отложен — нужен точный суммарный rendered height (content-driven, layout-level; редкий кейс, fitToHeight=0 — норма) |

### Состояние репозитория

```
src/
  index.ts                   Публичный API barrel пакета (convert*, renderStyledPdf,
                             FontRegistry/parseTtf, hyphenation + типы)
  opc/                       ECMA-376 Part 2 — package + getPartRelationships + resolveRelatedPart
  ooxml/spreadsheet/         §18 — SpreadsheetML парсеры
    cell-reference.ts        A1 ↔ {row, col} (0-indexed внутри, 1-indexed снаружи)
    defined-name-ref.ts      §18.2.5 — parseAreaRef (Print_Area "Sheet!$A$1:$D$20"
                              → bounding box) + parseTitleRowRange (Print_Titles
                              row-диапазон "$1:$2"); strip qualifier + $ + кавычки
    workbook-parser.ts       xl/workbook.xml → ParsedWorkbook { sheets, date1904,
                              definedNames[] (name/localSheetId/value — Print_Area/Titles) }
    shared-strings-parser.ts xl/sharedStrings.xml → string[]
    worksheet-parser.ts      xl/worksheets/sheet*.xml → cells (с styleIndex)
                              + ColumnWidth[] (<col>) + MergedRange[] (<mergeCell>)
                              + RowHeight[] (<row ht customHeight>, pt-precision)
                              + XlsxPageMargins (inches) + XlsxPageSetup (paperSize/
                              orientation/scale/fitToWidth/fitToHeight)
                              + fitToPage (<pageSetUpPr>) + XlsxPrintOptions (gridLines/
                              horizontal·verticalCentered) + rowBreaks/colBreaks (<brk>)
    styles-parser.ts         xl/styles.xml → fonts, fills, borders, cellXfs, numFmts
    number-format.ts         applyNumberFormat: built-in numbers (0-49) + custom code
                              parser + Excel date formats (14-22, 45-47, custom
                              date tokenizer с m=month/minute disambiguation)
  ooxml/wordproc/            §17 — все парсеры
    po-helpers.ts            preserveOrder XML walker (w:/r:/xml: namespaces)
    po-to-flat.ts            адаптер PO → non-PO для парсеров свойств
    document-parser.ts       BodyElement[] + parseSection + parseHeaderFooter
    table-parser.ts          tblPr/tblGrid/tr/tc parser
    numbering-parser.ts      §17.9 — abstractNum/num parser
    styles-parser.ts         §17.7 — styles.xml parser
    paragraph-properties.ts  §17.3.1 — pPr с numPr и hanging indent
    run-properties.ts        §17.4.1 — rPr
    drawing-parser.ts        §20 — w:drawing → picture | wps:wsp shape; MCE
                             resolveMc/expandMcChildren (Choice>Fallback); spPr
                             (xfrm/prstGeom/custGeom/solidFill/ln), txbx/bodyPr
  ooxml/drawingml/           §20/§21 DrawingML общее
    colors.ts                RawColor/ColorResolver, makeColorResolver, tx/bg
                             алиасы, дефолтная Office-2013-палитра
    theme-parser.ts          a:clrScheme → name→hex (srgbClr/sysClr@lastClr)
    chart-parser.ts          §21.2 — chart1.xml: chartSpace/plotArea/bar|line|
                             pieChart, c:ser cat/val КЭШ (num/strCache), title,
                             legend, barDir/grouping, цвет серии (fill или a:ln)
  ooxml/math/                §22 OfficeMathML
    omml-parser.ts           m:oMath/oMathPara → MathNode; r/f/sSup/sSub/
                             sSubSup/sPre/rad/nary/func/limLow/limUpp/d/m/acc/
                             bar/groupChr; m:rPr стили; m:val через m: namespace
  document-model/            Полная типизированная модель: BodyElement, Paragraph,
                             Run, Table/Row/Cell, Border, CellMargins, Style, StyleSheet,
                             Numbering, AbstractNumbering, NumberingLevel, NumberingReference,
                             SectionProperties, PageSize, PageMargins, HeaderFooterReference,
                             ShapeBlock/ShapeGeometry/ShapeFill/ShapeLine/ShapeTransform/
                             ShapeTextBody/CustomGeometry (DrawingML shapes),
                             Chart/ChartSeries/ChartBlock (charts),
                             MathNode (рекурсивная: run/fraction/script/radical/
                             nary/func/limit/delimiter/matrix/accent/bar/groupChr)
                             + Run.math (inline OfficeMath)
  style-cascade/             Resolver: defaults → para style → char style → direct
                             (+ bidi/rtl флаги в resolved properties)
  numbering/                 NumberingState (per-numId counters) + marker formatter
  line-breaker/              Knuth-Plass total-fit (box/glue/penalty, active-list cap)
  hyphenation/               Liang trie + en-US/ru patterns (lazy load)
  bidi/                      UAX #9 — char-types (Latin/Hebrew/Arabic ranges) +
                             algorithm (P/X/W/N/I/L с isolates) + reorderVisual (L2)
  font/                      TTF parser, subsetter, FontRegistry, opentype-layout (GSUB/GPOS)
  crypto/                    Цифровые подписи (§12.8) с нуля, только WebCrypto:
                             asn1.ts (DER-энкодер seq/set/oid/integer/octetString/
                             cmsTime/setOfBody + reader readTlv/children/
                             certIssuerAndSerial), cms.ts (buildPkcs7Detached →
                             PKCS#7 SignedData detached, SHA-256 + RSA)
  pdf/                       Objects (+ PdfHexString/PdfRawToken/unicodeString), serialize, writer
                             (+ /Info trailer, configurable version, /ID hash),
                             cid-font (Type0+CIDFontType2, subset-tag prefix),
                             image-xobject (JPEG /DCTDecode, PNG decoder с фильтрами
                             None/Sub/Up/Average/Paeth + /FlateDecode + /SMask для alpha,
                             flattenAlpha для PDF/A), icc-profile (sRGB ICC v2 с нуля),
                             xmp (XMP packet с pdfaid, conformance A/B),
                             struct-tree (§14.7-14.8 tagged PDF: StructNode/
                             StructTreeBuilder → StructTreeRoot + StructElem +
                             ParentTree, MCR-ссылки, pre-order DFS для
                             детерминизма; типы Document/Sect/H1–H6/P/L/LI/Lbl/
                             LBody/Table/TR/TH/TD/Figure без RoleMap),
                             embedded-file (§7.11 / PDF/A-3 §6.8 associated files:
                             /EmbeddedFile stream + /Filespec с /AFRelationship —
                             для встраивания исходного docx/xlsx),
                             signature (§12.8 цифр. подпись: addSignaturePlaceholder
                             — /Sig-поле + /AcroForm + placeholder ByteRange/Contents;
                             async signPdf — ByteRange/SHA-256/CMS/splice),
                             vector-graphics (path/fill/stroke/dash/transform emit,
                             emitVectorShape), arc-to-bezier (κ-метод, ellipse/
                             roundRect), preset-geometry (15 presets + custGeom →
                             VectorPath, y-up), chart-geometry (Chart → scene:
                             niceScale + buildFrame + bar/line/pie, чистая, через
                             measure-fn), math-layout (§22 матнабор с нуля:
                             layoutMath → MathBox с glyph/rule/path items;
                             структурные элементы — дроби/радикалы/большие
                             операторы/скобки/акценты — рисуются вектором, без
                             матшрифта),
                             styled-page-renderer:
                               - layout para + table + image блоков, multi-section
                                 page size/margins из sectPr
                               - Knuth-Plass + Liang hyphenation, justify, gridSpan,
                                 vMerge с suppression границ
                               - BiDi: tokenize по embedding level, визуальный
                                 reorder (L2) + реверс RTL code points на эмите
                               - dedup границ (per-cell top+left, последний bottom/right)
                               - default/first/even headers/footers, row split
                               - **повтор header-строк** (RowProperties.isHeader →
                                 RowLayout.isHeader): лидирующий префикс header-строк
                                 переэмитится наверху каждой continuation-страницы
                                 как artifact (w:tblHeader / xlsx Print_Titles)
                               - **table alignment** (TableProperties.alignment →
                                 TableBlock.xOffsetPt): center/right-сдвиг узкой
                                 таблицы; **manual page break** (RowProperties.
                                 pageBreakBefore → RowLayout.breakBefore): force-flush
                                 перед строкой (xlsx rowBreaks)
                               - image XObject embed + page /XObject resource + cm/Do
                               - shapes (DrawingML): layoutShapeBlock (EMU→pt,
                                 width/height clamp), buildShapeTransform (поворот
                                 вокруг центра + flip), атомарная пагинация,
                                 shape-проход ПЕРЕД текстом; текст фигуры в inset-
                                 rect с вертикальным якорем — эмитится как line
                               - charts: layoutChartBlock → buildChartLayout
                                 (scene из chart-geometry → shape/line примитивы
                                 в локальном y-up, транслируются на страницу);
                                 unknown-тип → placeholder-рамка
                               - math: MathToken (kind в Token union) — атомарный
                                 box со своим ascent/descent (line-height растёт);
                                 layoutMath via measure+fontFor(variant); emit
                                 glyph (Tj) + rule (re/f) + path (vector); display
                                 (oMathPara) центрируется через jc
                               - /Info dict из docProps/core.xml или options.info
                               - PDF/A-1b (options.pdfA): OutputIntent+ICC, XMP
                                 /Metadata, PDF 1.4, /ID, subset prefix, flatten alpha
                               - tagged PDF (options.tagged / pdfA='PDF/A-1a'):
                                 paginate строит StructTree (P/H1–H6 по
                                 outlineLvl, L/LI/LBody по ilvl со стеком,
                                 Table/TR/TD, Figure+/Alt); structId/artifact на
                                 DrawCommand; emit оборачивает каждую строку
                                 `/<тип> <</MCID n>> BDC…EMC`, фигуры/картинки —
                                 emitTaggedRuns, декорации/header-footer —
                                 /Artifact; Catalog /MarkInfo+/StructTreeRoot+
                                 /Lang+/ViewerPreferences; **tagged=false ⇒
                                 байт-в-байт как раньше**
  converter/                 convertDocxToPdf — fonts, styles.xml, numbering.xml,
                             sectPr, header/footer parts, media parts через relationships,
                             word/theme/theme1.xml → ColorResolver (schemeClr→hex),
                             chart parts (loadCharts) → options.charts.
                             convertXlsxToPdf — workbook → ВСЕ sheets → grid с
                             cell styles (font/alignment/numFmt/shading/borders) +
                             dates (serial → calendar, 1900/1904 epoch) + merged cells +
                             column widths + row heights (atLeast rule) + sheet title +
                             pageBreakBefore на 2-м+ + SectionProperties из первого
                             листа (pageSize по paperSize+orientation, pageMargins из
                             inches).
                             **Print-model:** gridline-подавление (синтет. grid только
                             при printOptions gridLines="1"); print area (_xlnm.Print_Area
                             → клип окна rowStart/colStart); fit-to-page (computePrintScale
                             → shrink шрифтов+высот по scale/fitToWidth); print titles
                             (_xlnm.Print_Titles → isHeader → повтор); manual rowBreaks
                             (→ pageBreakBefore); horizontalCentered (→ alignment center)
scripts/corpus/              Валидация vs LibreOffice: build-corpus (синтетика),
                             fetch-corpus (реальные POI docx/xlsx), lib (PPM/
                             stext парсеры, diff-метрики, sofficeToPdf +
                             sofficeToPdfSandboxed), run (оркестратор; CORPUS_DIR/
                             CORPUS_SANDBOX), convert-one (изолированный конвертер
                             для sandbox-режима), sandbox/Dockerfile (headless
                             LibreOffice для untrusted-входа)
tests/
  fixtures/fonts/            Roboto Regular/Bold/Italic/BoldItalic (Apache 2.0) —
                             тест-фикстуры (офлайн); НЕ в пакете, НЕ продукт
  fixtures/build-docx.ts     buildDocx, buildDocxFromBody (numberingXml/headerXml/
                             footerXml/settingsXml/themeXml/images/charts), buildRichDocx
  fixtures/build-png.ts      buildTinyPng — программный генератор PNG для тестов
                             (zlibSync + CRC32, RGBA 8-bit)
  fixtures/build-xlsx.ts     buildXlsx — генератор xlsx с per-cell styleIndex,
                             merged cells, column widths, multi-sheet, row heights,
                             pageMargins/pageSetup (+scale/fit), date1904, printOptions,
                             fitToPage, rowBreaks, definedNames (Print_Area/Titles)
  *.test.ts                  367 тестов: rPr (+rtl), pPr (+bidi +outlineLvl), styles, cascade,
                             document-parser, table-parser, numbering, section
                             (+titlePg/even/multi-section), image (+inline mixed + robustness),
                             number-format (built-in + custom + dates +
                             disambiguation + 1904 epoch), xlsx (cell-ref + parsers
                             + e2e: bold font, #,##0, custom format, merges, multi-sheet,
                             cell fill, thick per-cell borders, row heights,
                             pageSetup paperSize/landscape, date1904
                             + **print-model**: parseAreaRef/parseTitleRowRange, print
                             model parsing, gridline on/off, print-area clip (+offset),
                             fit-to-page/scale shrink, Print_Titles header-repeat,
                             manual rowBreaks page-count, horizontalCentered shift),
                             styled-render
                             (mixed fonts, color, size, alignment, justify, gridSpan,
                             vMerge, borders, headers/footers, MediaBox, row split),
                             knuth-plass, opentype-layout (GSUB/GPOS), info-metadata,
                             bidi (P/X/W/N/I/L + reorder), bidi-render (e2e),
                             pdfa (ICC + XMP + e2e), remote-fonts (substitution +
                             injected-fetch авто-загрузка),
                             vector-graphics (emit B/f/S, dash/cap/join),
                             arc-to-bezier (κ + ellipse/roundRect), preset-geometry
                             (15 presets + custGeom scale/y-flip), shape (parse +
                             e2e: presets, fill/line/dash, MCE, custGeom, текст в
                             фигуре, anchor/oversized/mixed-run, theme-цвета),
                             theme (clrScheme + resolver alias),
                             chart-geometry (niceScale/formatTick + bar/line/pie
                             scene), chart (parse + e2e: column bars/colors/axes,
                             line polyline, pie wedges/legend, MCE chart drawing),
                             math (parse + engine + e2e: fractions, sub/sup,
                             radical surd, n-ary ∑, delimiters, matrix, accents,
                             display centering),
                             tagged-pdf (M6 — StructTreeRoot → Document/H1–H6/P/
                             L/LI/LBody/Table/TR/TD/Figure, per-line MCID +
                             ParentTree resolution, pagination-артефакты, alt-
                             текст фигур, /Lang, PDF/A-1a conformance gate:
                             «нет painting-операторов вне marked content»),
                             verapdf (opt-in формальная валидация через veraPDF
                             CLI — профили 1b/1a/2b/2a/3b/3a: текст/лигатуры,
                             rich tagged, прозрачность, встроенный источник;
                             пропуск если veraPDF не установлен, env VERAPDF), smoke
tests/output/
  table-demo.pdf             5×3 таблица с insideH/insideV + auto-fit + цвет
  merge-demo.pdf             7-строчная таблица с gridSpan=3 title, gridSpan=2 итог, vMerge
  justify-demo.pdf           Lorem ipsum: left vs both
  list-demo.pdf              Нумерованный многоуровневый + bullet списки
  header-footer-demo.pdf     US Letter + центрированный header + footer на 36pt от краёв
  image-demo.pdf             Заголовок + два PNG (красный 4″×2″, синий 5″×1.5″) с описанием
  shapes-demo.pdf            DrawingML: roundRect с центрированным текстом (Callout),
                             ellipse, rightArrow, hexagon (schemeClr accent5),
                             dashed-outline rect, rect повёрнут на 20°
  charts-demo.pdf            DrawingML charts: clustered column (2 серии, легенда),
                             line (2 серии), pie (4 слайса, %-подписи, легенда)
  math-demo.pdf              OfficeMathML (display): квадратное уравнение, ∑k²,
                             гауссов интеграл ∫e^(-x²), предел e, матрица+детерминант
  payroll.pdf                xlsx: 8×6 таблица зарплат, plain styling
  sales-report.pdf           xlsx со styles: merged title row A1:E1, bold headers,
                             #,##0 числа справа, красные жирные итоги, col widths
  multi-sheet.pdf            xlsx с 3 листами — Доходы / Расходы / Прибыль —
                             каждый на своей странице PDF через pageBreakBefore
  sales-styled.pdf           xlsx с синим header (fill+white text), zebra rows,
                             жёлтой итоговой строкой с красными medium-thick
                             per-cell borders
  schedule.pdf               xlsx-расписание Q1 с датами в форматах m/d/yyyy,
                             dd.mm.yyyy, d mmmm yyyy, dddd hh:mm
  row-heights-demo.pdf       xlsx с custom row heights: header 28pt, body rows
                             18/50/80pt — видно разные высоты строк
  landscape-letter.pdf       xlsx с pageSetup paperSize=1 + orientation=landscape
                             → Letter landscape 792×612pt MediaBox
  date1904-demo.pdf          xlsx с workbookPr date1904="1" — serial 0 рендерится
                             как 1 January 1904 (а не 30 December 1899)
  bidi-demo.pdf              UAX #9: LTR vs RTL (w:bidi) выравнивание + LTR-числа
                             в RTL-абзаце (на Roboto без ивритских глифов —
                             видно только выравнивание/числа, реверс глифов
                             покрыт юнит-тестами bidi.test.ts)
  pdfa-demo.pdf              PDF/A-1b: PDF 1.4 + /ID + OutputIntent(sRGB ICC) +
                             XMP pdfaid + subset-шрифты (структурно валиден)
  tagged-demo.pdf            PDF/A-1a (tagged): H1/H2 заголовки, параграф,
                             таблица с заливкой+границами, вложенный список,
                             картинка с alt-текстом — полное дерево структуры
                             (Document/H1/H2/P/Table/TR/TD/L/LI/LBody/Figure),
                             в Acrobat виден Tags/Reading Order
  pdfa2-demo.pdf             PDF/A-2a (tagged, PDF 1.7): полупрозрачная картинка
                             сохраняет SMask, страница несёт transparency group;
                             veraPDF --flavour 2a PASS
  pdfa3-demo.pdf             PDF/A-3a (tagged): исходный docx встроен как
                             associated file (/AFRelationship /Source, в панели
                             Attachments — source.docx); veraPDF --flavour 3a PASS
  styled.pdf                 mixed fonts/colors/sizes/alignment
конфиги (корень)
  package.json               name=reamkit, MIT, exports, scripts (build, test,
                             test:eslint/types/lib/build, lint, format, docs:api,
                             corpus), prepack/prepublishOnly=build, publishConfig
  tsconfig.json              module ESNext + moduleResolution Bundler, noEmit,
                             paths @/*→src/*, strict
  tsconfig.eslint.json       extends tsconfig; src+tests+scripts+configs
                             (граф для type-checked ESLint)
  vite.config.ts             tanstackViteConfig (2 entry: index + document-model)
                             + dts-хук (расширения, dir→/index.js) + sourcemap off
  vitest.config.ts           resolve.alias @→src; include tests/**/*.test.ts
  eslint.config.js           @tanstack/config/eslint (tanstackConfig) —
                             type-checked, import-гигиена; формат за prettier
  .prettierrc.json/.ignore   singleQuote/semi/trailingComma all/printWidth 100
  .github/workflows/         ci.yml (гейты + opt-in corpus), release.yml (tag→npm)
```

**409/409 тестов зелёные** (включая live-прогон veraPDF всех 8 PDF/A-профилей, если установлен). Strict TypeScript (exactOptionalPropertyTypes + noUncheckedIndexedAccess) — чисто.

## Пакет (npm)

Имя пакета — **`reamkit`** (бренд Ream). Сборка через **Vite + `@tanstack/config`**
(`tanstackViteConfig`). `npm pack` проверен install-смоком: импорт (root +
subpath) резолвится у NodeNext-потребителя, async/sync конвертеры и PDF/A
работают.
- **Импорты без расширений и без `/index`:** исходники пишутся `from '@/font'`
  (не `'@/font/index.js'`). tsconfig: `module: ESNext` + `moduleResolution:
  Bundler`, `noEmit: true` (typecheck only). Vite/rollup проставляют пути в
  сгенерированном `.js`; хук `beforeWriteDeclarationFile` восстанавливает
  расширения в `.d.ts` для NodeNext-потребителей — **директория → `/index.js`,
  файл → `.js`** (резолвит по дереву исходников). Проверено: `tsc` с
  `moduleResolution: NodeNext` у потребителя против опубликованных `.d.ts`
  проходит.
- **build:** `npm run build` = `vite build`. Вывод в **`dist/esm/`** (ESM,
  cjs:false), per-file модули + `.d.ts`, без source maps. `prepack` собирает.
  fflate/fast-xml-parser остаются external (vite-plugin-externalize-deps).
- **Тулинг (под `@tanstack/config`):** ESLint `tanstackConfig`
  (`@tanstack/config/eslint`) — type-checked (project=`tsconfig.eslint.json`),
  import-гигиена, `import type`; форматирование оставлено Prettier
  (eslint-stylistic у TanStack делает только `spaced-comment`, конфликта нет;
  `no-non-null-assertion` off). **publint + are-the-types-wrong** валидируют форму
  пакета (`npm run test:build`, attw `--profile esm-only` — node10/CJS игнорятся,
  мы ESM-only). **typedoc** → Markdown API-доки (`npm run docs:api` →
  `docs/reference/`, gitignored, регенерируемо). Prettier `.prettierrc.json`
  (singleQuote/semi/trailingComma all/printWidth 100; `npm run format`/`:check`).
- **Алиасы:** `@/* → src/*` (tsconfig `paths` + `baseUrl`). Все импорты в src/
  tests/scripts через `@/`, без хвостового `/index`. Резолв: tsc (paths),
  vite build (tanstack tsconfigPaths plugin), vitest (`resolve.alias` в
  vitest.config). В опубликованном dist алиасов нет — переписаны в относительные
  пути с `.js`.
- **Шрифты скачиваются автоматически** (`src/fonts/remote-fonts.ts`). Главный
  API — **async**: `await convertDocxToPdf(docx)` без `fonts` сам тянет открытый
  шрифт-замену с CDN (`@expo-google-fonts` через jsDelivr, полные TTF). Подмена
  как в LibreOffice: sans→Roboto, serif→Tinos, mono→Cousine (детект по `w:ascii`
  в документе). Кэш по URL, инъекция `fontFetch` для тестов/офлайна. Синхронные
  `convertDocxToPdfSync`/`convertXlsxToPdfSync` требуют `fonts` (без сети).
- **Публичный API:** `src/index.ts` → `dist/esm/index.js`. Экспортирует
  `convertDocxToPdf`/`convertXlsxToPdf` (async, авто-шрифт) +
  `convertDocxToPdfSync`/`convertXlsxToPdfSync` (+ опции), `fetchFontSet`/
  `resolveFamilyKey`, `renderStyledPdf`/`StyledRenderOptions`/`DocumentInfo`,
  `FontRegistry`/`parseTtf`/`subsetTtf` (+ типы), hyphenation (+ типы).
- **exports map:** `.` → `dist/esm/index.js`, `./document-model` →
  `dist/esm/document-model/index.js`, `./package.json`. `main`/`types` тоже на `dist/esm`.
- **Шрифты НЕ бандлятся** — качаются в рантайме (см. секцию выше). В репозитории
  шрифтов нет каталога `assets/`; Roboto-фикстуры для офлайн-тестов лежат в
  `tests/fixtures/fonts/`.
- **Браузер-чистая.** В коде библиотеки **ноль** импортов `node:*`/`Buffer`/
  `process`/`__dirname` — путь конвертации использует только нативное
  (`fetch`, `TextEncoder/Decoder`, `Uint8Array`, fflate, fast-xml-parser).
  Работает в браузере, Node, serverless, edge без изменений. CLI/bin удалён —
  потребитель сам решает, как подать `Uint8Array` (File/fetch/fs).
- **files:** `dist`, `README.md`. tests/scripts/corpus НЕ публикуются.
- **Опубликовано:** `reamkit@0.1.0-alpha.0` в npm — MIT, **provenance** (SLSA v1
  attestation + sigstore-подпись из GitHub Actions), 162 файла / 731 КБ unpacked.
  dist-tag `alpha` (`npm i reamkit@alpha`); `latest` тоже на альфе — штатно для
  первого publish, переедет на первую стабильную. Репо: `alex-krassavin/reamkit`.
- **Релиз — tag-triggered** (`.github/workflows/release.yml`): бамп `version` →
  `git tag vX.Y.Z && git push --tags` → CI гоняет гейты, сверяет тег с
  package.json, `npm publish --provenance`, пререлизы (`-alpha`) уходят в
  non-default dist-tag, создаёт GitHub Release. Нужен secret `NPM_TOKEN` типа
  **Automation** (обходит 2FA в CI).
- **Грабли публикации (выучено):** (1) unscoped `reamjs` зарезан npm
  similarity-guard'ом (нормализуется в `ream.js`) → переименовали в `reamkit`;
  бренд «Ream» в доках остаётся. (2) `EOTP` на publish = токен не Automation →
  пересоздать как Automation. (3) `latest` навешивается на первый publish даже с
  `--tag alpha`. (4) шаг «Create GitHub Release» под `if: event==push` —
  через `workflow_dispatch` пропускается (первый релиз ушёл dispatch'ем после
  транзиента, поэтому Release-объекта на GitHub нет — можно создать вручную).
- **Подстраховать на будущее:** retry на шаг `npm publish` (транзиент уронил
  тег-релиз, помог ручной dispatch); опционально — переезд на tokenless
  **Trusted Publishing** (теперь пакет существует, настраивается в один заход).

## Известные ограничения

- **Knuth-Plass** — реализован, tolerance ratio 4 для feasibility. Active list capped at 200 узлов, отбрасывает хвост с худшими demerits.
- **Hyphenation** — Liang algorithm + en-US/ru patterns (lazy load через `getHyphenator(lang)`). Hyphenator передаётся через `options.hyphenator` в `convertDocxToPdf`. Для других языков нужно подгрузить свои patterns через `splitPatternBundle` + `createHyphenator`.
- **GSUB/GPOS** — поддерживаются только Lookup Type 4 (ligatures) и Type 2 Format 1/2 (pair kerning). Contextual substitutions (Type 5/6), mark positioning (GPOS Type 4-6), и язык-специфичные features (script/lang filter) не применяются.
- **BiDi (RTL)** — UAX #9 реализован (P/X/W/N/I/L rules с isolates). Парсятся `<w:bidi/>` (RTL база абзаца) и `<w:rtl/>` (RTL run, оборачивается в RLE…PDF). Авто-детект базы по первому strong-символу. Токены режутся по embedding level, на эмите реордерятся визуально (L2) + code points реверсятся для odd-level. Ограничения: (a) N0 paired-bracket rule не реализован; (b) арабское cursive joining (init/medi/fina формы) не применяется — глифы в изолированной форме, но в правильном визуальном порядке; (c) нет RTL-шрифта в bundle (Roboto без иврита/арабского) — рендер глифов требует подключения своего шрифта.
- **Image alignment** — пока всегда left.
- **PNG ограничения** — поддерживаются color types 0/2/4/6, bit depth 8, без interlace. Palette (type 3), interlaced и bit depth ≠ 8 не поддерживаются. **Graceful degradation (#109):** неподдерживаемая/битая картинка (включая неизвестные форматы и GIF) **не валит документ** — пропускается без висячей `/Im… Do`-ссылки, остальной контент рендерится (найдено POI-корпусом).
- **Footnotes/endnotes** — нет.
- **Float / anchored objects** — `<wp:anchor>` text wrap не реализован (anchored картинка трактуется как inline).
- **Header/footer numPr** — нумерация в headers/footers использует независимый счётчик.
- **OfficeMathML (M5)** — реализован from-scratch box-model матнабор: runs (авто-курсив переменных, m:sty/m:nor), дроби (m:f, noBar), индексы/степени (m:sSup/sSub/sSubSup/sPre), радикалы (m:rad с нарисованным surd + degree), n-арные (m:nary: ∑∏∫ нарисованы вектором + limLoc undOvr/subSup), функции (m:func), пределы (m:limLow/limUpp), разделители (m:d: растяжимые ()[]{}|⟨⟩), матрицы (m:m), акценты (m:acc: hat/bar/vec/dot/tilde нарисованы), m:bar, m:groupChr; inline + display (m:oMathPara центрируется); inter-atom spacing вокруг операторов/отношений. **Структурные элементы рисуются вектором → матшрифт не нужен.** **Ограничения:** (a) обычные символы берутся из bundled-шрифта (Roboto), поэтому символы вне его (→ U+2192, многие математические операторы) рендерятся как .notdef; полноценно нужен STIX/Cambria Math; (b) фигурная скоба { } нарисована приближённо (выглядит угловато); (c) нет text-style vs display-style различия размеров (display-метрики всегда); (d) нет italic-correction, точной MATH-таблицы, N0/contextual spacing; (e) повёрнутый/вертикальный матнабор, equation arrays (m:eqArr), m:box/m:borderBox/m:phant — нет.
- **DrawingML charts (M5)** — реализованы: **column/bar** (clustered), **line**, **pie**. Парсятся из chart1.xml по **кэшу** (numCache/strCache, формула не вычисляется); оси с Heckbert nice-scale + гридлайны + тики, подписи категорий, легенда (справа/снизу), заголовок, цвет серии (fill или a:ln) с accent-cycle fallback, pie %-подписи и dPt-цвета слайсов. **Не реализовано:** (a) stacked/percentStacked рендерятся как clustered; (b) area/scatter/doughnut/radar/3D → `unknown` (placeholder-рамка, не падает); (c) value-axis всегда от 0 (нет авто-минимума); (d) нет осевых заголовков, data labels (кроме pie %), combo (несколько chart-групп — берётся первая), secondary axis; (e) цвет серии не читается из chart `colorN.xml`/`styleN.xml` (берётся из серии или accent-cycle, который совпадает с дефолтной темой, но не с кастомной); (f) charts в headers/footers/ячейках таблиц не рендерятся (общее ограничение non-paragraph контента).
- **DrawingML shapes (M5)** — реализованы: `wps:wsp` через `w:drawing`, preset-геометрии (rect/roundRect/ellipse/triangle/rtTriangle/diamond/parallelogram/trapezoid/pentagon/hexagon/line/4×arrow; неизвестные → bounding rect), custGeom (moveTo/lnTo/cubicBezTo/quadBezTo/arcTo/close), solidFill/noFill, a:ln (width/color/dash/cap), поворот (`xfrm @rot`) + flipH/flipV, текст в фигуре (`wps:txbx`/`bodyPr` insets + вертикальный якорь), MCE (`mc:AlternateContent` → wps Choice, VML Fallback игнорируется), theme-цвета (`schemeClr` через theme1.xml + дефолтная Office-палитра). **Не реализовано:** (a) фигуры — только block-level (standalone-параграф); inline-фигура в run с текстом отбрасывается (текст сохраняется); (b) текст в повёрнутой фигуре рендерится без поворота (вертикально), заливка/обводка поворачиваются; (c) z-order глобальный по типу команды (заливка фигур всегда под текстом страницы); (d) gradFill/pattFill/blipFill (только solid/none), group shapes (`wpg`/`a:grpSp`), `lumMod`/`lumOff`/`shade`/`tint` модификаторы цвета, несколько subpath в custGeom (берётся первый `a:path`) — следующие итерации.
- **xlsx row heights** — `<row ht customHeight>` маппится в `heightRule='atLeast'`. Excel "exact" поведение (truncation) не повторяется: строки расширяются для длинного контента вместо его обрезки.
- **xlsx formulas** — берётся cached value `<v>`, формула не вычисляется.
- **xlsx page setup** — используется только первый sheet (multiple sections не поддержаны). Из paperSize реализованы Letter/Tabloid/Legal/A3/A4/A5/A6; неизвестные значения → fallback на A4.
- **TTC** (TrueType Collection) — поддерживается только одиночный sfnt.
- **PDF/A** — реализованы **PDF/A-1b** (visual) и **PDF/A-1a** (tagged/accessible) через `options.pdfA`. -1b: PDF 1.4 header, `/ID`, OutputIntent с встроенным sRGB ICC, XMP `/Metadata` с pdfaid, subset-tag в именах шрифтов, flatten альфы. Дескриптор субсета несёт **`/CIDSet`** (§6.3.5); **`/ToUnicode`** покрывает лигатуры (→ составляющие code points) и скан с U+0009 (TAB-глиф маркеров списков), §6.3.8. -1a добавляет логическую структуру (см. ниже). **Реализованы все 8 профилей** через `options.pdfA`: `PDF/A-1b/1a/2b/2u/2a/3b/3u/3a`. **PDF/A-2** (ISO 19005-2, PDF 1.7): уровень `u` (Unicode), **прозрачность сохраняется** (image `/SMask` + page transparency group с device-independent ICCBased sRGB, §6.2.4.3), `/CIDSet` опционален (не пишется для -2/-3). **PDF/A-3** (ISO 19005-3): встраивание любых **associated files** (`/AF` + `/AFRelationship`, `src/pdf/embedded-file.ts`); `options.attachments` + конвертерный `embedSource` встраивает исходный docx/xlsx как `/Source`. **Все 8 формально проходят veraPDF 1.30** (`--flavour 1b/1a/2b/2u/2a/3b/3u/3a`) — проверено на text/ligatures/lists/tables/figures/math/shapes/xlsx/transparency/embedded-source; opt-in gate `tests/verapdf.test.ts`. **Не реализовано:** object streams / compressed xref (классический xref — валиден и для -2/-3); JPEG2000.
- **Цифровые подписи (§12.8) ✅** — `signPdf()` / `options.signaturePlaceholder` / конвертерный `options.signature`. PKCS#7/CMS detached (`adbe.pkcs7.detached`, SHA-256 + RSA PKCS#1 v1.5) собирается с нуля (`src/crypto/` ASN.1+CMS), подписывается через WebCrypto, вписывается в `/Contents` поверх placeholder без сдвига байт (ByteRange). Невидимое `/Sig`-поле + `/AcroForm SigFlags 3`. **Внешне проверено** `openssl cms -verify` («CMS Verification successful»). **Не реализовано:** ECDSA/PAdES/timestamp, подпись внешних PDF (incremental update) — у нас integrated-placeholder; подпись + PDF/A одновременно не гарантирует строгий PAdES.
- **Tagged PDF / PDF/A-1a (M6)** — `options.tagged` или `pdfA: 'PDF/A-1a'`. Строится дерево логической структуры (`/StructTreeRoot` → `Document` → H1–H6 / P / L→LI→LBody / Table→TR→TD / Figure), marked content (`BDC/EMC` + MCID на каждой строке текста), `/ParentTree` (page `/StructParents`), `/Tabs /S`, `/MarkInfo`, `/Lang` (авто из `w:lang`), `/ViewerPreferences /DisplayDocTitle`, alt-текст фигур (`/Alt` из `wp:docPr @descr/@title`, fallback Image/Shape/Chart). Заголовки — из `w:outlineLvl` (0–8→H1–H6) или эвристики styleId; списки — вложение по `w:ilvl`; декорации (заливки/границы) и колонтитулы — `/Artifact` (header/footer типизированы `/Pagination`). Marked-content тег = тип структуры. `tagged=false` ⇒ вывод байт-в-байт как до M6. **Ограничения:** (a) `Lbl` для маркеров списка не выделяется (маркер внутри P — нужны intra-line MCID-диапазоны); (b) ✅ заголовочные строки таблиц (`<w:tblHeader/>`) теперь `TH` + `/Scope /Column` (#120); (c) фигуры с текстом — Figure + соседние P (не вложенные); (d) ✅ `/ColSpan`/`/RowSpan` атрибуты пишутся для span-ячеек (#121); (e) per-element `/Lang` (смена языка внутри документа) не пишется. **Формально проходит veraPDF `--flavour 1a` (Level A).**

## Технический долг — security/robustness (✅ РЕШЕНО)

Найдено при триаже POI-корпуса (продолжение линии C1). Оба **исправлены**
библиотечными self-limit'ами, не зависящими от внешних harness-капов:

- **✅ #41 (#114) — xlsx OOM.** `CVLKRA-KYC` (204 КБ) рвал 512 МБ heap: `sheet15`
  объявляет `A1:XFD23` — 49 194 ячейки, из них лишь 48 со значениями, остальные
  пустые стилизованные (стайл на всю строку до колонки XFD/16384) → плотная
  сетка `maxColumn+1`. Фикс: сетка ограничена «used range» (содержательные
  ячейки + merges); + бэкстоп-кап `MAX_GRID_COLS=1024`/`MAX_GRID_ROWS=50000` и
  клэмп merge-расширения для untrusted-входа. Конвертится (74 стр.).
- **✅ #42 (#115) — sharedStrings parse-DoS.** `poc-shared-strings.xlsx`: 1 общая
  строка 1 048 583 символа × 12 000 ссылающихся ячеек = шейпинг ~12 ГБ → хэнг.
  Фикс: (1) парсеры капят строку ячейки до Excel-лимита 32 767 символов (shared +
  inlineStr); (2) конвертер ограничивает суммарный текст листа
  `MAX_SHEET_TEXT_CHARS=1 000 000`. Конвертится за ~2с (76 стр.).

## Опционально осталось (M0–M6 завершены — всё ниже НЕОБЯЗАТЕЛЬНО)

**Настоящих фич из роадмапа НЕ осталось.** M0–M4 (docx/xlsx→layout→PDF, шрифты,
текст, BiDi), M5 (shapes + charts + OfficeMathML), M6 (Tagged PDF + PDF/A-1/2/3 +
цифровые подписи) — **закрыты**. Библиотека готова к использованию as-is. Ниже —
только полировка/расширения по приоритету реальных документов.

### Backlog (приоритизированный план — задачи #54–#66 в трекере)

Каждая строка = коммитируемый юнит. Приоритет P1 (следующий рычаг) → P4 (нишевое).

| # | P | Пункт |
|---|---|---|
| ~~#54~~ | ✅ | **xlsx: клиппинг текста ячейки** — СДЕЛАНО (#124): cell-overflow модель (перелив в пустые / обрезка у занятых). `xlsx-jdbc` 66→95%. Достаёт genuine-overflow подмножество |
| ~~#55~~ | ✅ | **xlsx number-format** — СДЕЛАНО (#125): `;@` date text-section leak (15 файлов) + scientific notation (numFmtId 11 был сломан + engineering). Замер: 47813 уже char-identical Calc — explicit-decimal корректен, tail = reading-order, не форматы |
| ~~#56~~ | ✅ | **Charts типы** — СДЕЛАНО (#126): stacked/percentStacked bar+area, area (заливка), scatter (xVal/yVal, 2 числовые оси), doughnut (центр-дыра). +16 тестов |
| ~~#57~~ | 🟡 | **Charts polish** — СДЕЛАНО частично (#127): data labels (showVal), axis titles (горизонт.), value-axis auto-min (line/scatter). **Отложено:** colorN/styleN цветовые карты (отдельные парты, низкая отдача) |
| ~~#58~~ | 🟡 | **Math** — СДЕЛАНО частично (#128): equation arrays `m:eqArr`. **Отложено:** матшрифт STIX/Cambria (MATH table — milestone-размер), полный text/display sizing |
| ~~#59~~ | 🟡 | **Shapes** — СДЕЛАНО частично (#129): lumMod/lumOff/shade/tint colour transforms + gradFill→average. **Отложено:** group shapes (`wpg`/`grpSp`), text rotation |
| ~~#60~~ | 🟡 | **Tagged PDF** — СДЕЛАНО частично (#130): per-element `/Lang` (закрывает (e)). **Отложено:** `Lbl` маркеров списков (intra-line MCID-split — закрывает (a)) |
| ~~#61~~ | 🟡 | **xlsx print-model хвост** — СДЕЛАНО частично (#131): fitToHeight (через Σ row-heights). **Отложено:** verticalCentered |
| ~~#62~~ | 🟡 | **Подписи** — СДЕЛАНО частично (#132): ECDSA (P-256/384/521). **Отложено:** PAdES, RFC-3161 timestamp (нужен TSA/сеть), подпись внешних PDF (incremental) |
| ~~#63~~ | 🟡 | **PDF** — СДЕЛАНО частично (#133): JPEG 2000 (`/JPXDecode` pass-through). **Отложено:** object streams / compressed xref (PDF/A-1-forbidden, инвазивно) |
| ~~#64~~ | 🟡 | **Типографика** — СДЕЛАНО частично (#134): Arabic cursive joining (joining-алгоритм + GSUB Type 1 init/medi/fina). **Отложено:** GSUB/GPOS Type 5/6 chaining, mark positioning (GPOS Type 4) |
| ~~#65~~ | 🟡 | **Типографика** — СДЕЛАНО частично (#135): BiDi N0 paired-bracket resolution. **Отложено:** bundled RTL-шрифт (упаковочное решение; покрывается embedded #118) |
| ~~#66~~ | 🟡 | **Инфра** — СДЕЛАНО частично (#136): CI workflow (gates + opt-in corpus job). **Отложено:** GovDocs1 scaling, font-agnostic visual-метрика |

Детали по категориям ниже.

**1. Fidelity форматов (M5 polish):**
- **Charts:** stacked/percentStacked (сейчас рисуются как clustered),
  area/scatter/doughnut/radar/3D (→ placeholder-рамка), авто-минимум value-оси,
  осевые заголовки + data labels, цвета из `colorN.xml`/`styleN.xml`,
  combo/secondary axis.
- **Math:** матшрифт (STIX/Cambria) для символов вне Roboto, text/display style,
  equation arrays, лучшая фигурная скоба.
- **Shapes:** `gradFill`, group shapes (`wpg`/`a:grpSp` chOff/chExt),
  `lumMod`/`tint`/`shade`, поворот текста, inline-фигуры в run.

**2. Tagged PDF / PDF-UA полировка:**
- ✅ **`<w:tblHeader>` → `TH` + `/Scope /Column`** (#120) и **`/ColSpan`/`/RowSpan`
  для span-ячеек** (#121) — СДЕЛАНО, veraPDF-clean (Level A).
- Осталось: `Lbl` для маркеров списков (intra-line MCID-диапазоны), per-element
  `/Lang` (смена языка внутри документа), вложенные Figure+текст.

**3. Хвост xlsx print-model (#45d deferred):** `fitToHeight` (нужен точный
суммарный rendered-height — content-driven, layout-level), `verticalCentered`.
Редкие кейсы.

**4. Расширения подписей (#119 deferred):** ECDSA (сейчас RSA PKCS#1 v1.5),
**PAdES** + RFC-3161 timestamp, подпись **внешних** PDF через incremental-update
(сейчас integrated-placeholder — только наши PDF); подпись+PDF/A одновременно не
гарантирует строгий PAdES.

**5. PDF-расширения:** object streams / compressed xref (сейчас классический
xref — валиден и для -2/-3), JPEG2000.

**6. Типографика (RTL / сложные скрипты):**
- Arabic cursive joining (GSUB `init`/`medi`/`fina`/`isol`), GSUB/GPOS Type 5/6
  (contextual) + mark positioning, BiDi N0 (paired-bracket resolution), RTL-шрифт
  в bundle (или загрузка пользовательского) для рендера иврита/арабского.

**7. Инфраструктура / валидация (не фичи продукта):**
- ✅ **3 прогона корпуса СДЕЛАНЫ** (349 xlsx, sandbox): print-model (#52) улучшил visual
  на 73 файла; диагностика → дроп заголовка листа (#53/#122) → **✅ 19 → 163 · ⚠️ 309 →
  165** (xlsx 5%→47% clean). См. «✅✅ ОБНОВЛЕНО» в xlsx-секции + `report-xlsx-titledrop.md`.
- ✅ **3-й прогон (текущий код после #54–#57, Docker-sandbox)** → `corpus/report.md`:
  **✅ 163→164 · ⚠️ 165→164 · ❌ 21** (net **+1** vs title-drop). Diff: **+4** (genuine-overflow
  файлы, которые #54-клиппинг подрезал к Calc: `VLookupFullColumn` 4650→**4036=ref ровно**,
  `CustomXMLMappings`, `CustomXmlMappings-inverse-order`, `59021`) / **−3** (over-clip:
  `56511`/`56822`/`test_conditional_formatting` — мы РЕЖЕМ текст, который Calc показывает).
  **Находка про #54-клиппинг:** он фундаментально ограничен **font/column-model mismatch с
  Calc** — мы рендерим Roboto + наша модель ширины колонки, Calc рендерит Liberation (уже)
  → `56511` «dlgkdflgdfjkl»×4 в 8.43-char колонке: у Calc влезает (узкий шрифт), у нас режется.
  **Width-based клиппинг (по реальным advance-ширинам) ИССЛЕДОВАН и ОТВЕРГНУТ** — A/B показал,
  что он хуже: `VLookupFullColumn` ✅→⚠️ (4036→3729), `CustomXMLMappings` ✅→⚠️; только `56822`
  лучше. Оставлен shipped char-count клиппинг (#54, net +1). Остаток xlsx ⚠️ = этот font-mismatch
  + column-major reading-order, не исправимо без точных метрик шрифта Calc. **XXE/бомбы**: весь
  349-прогон через sandbox безопасен (`xxe_in_schema` ✅, exit 0, без хэнгов).
- ⭐ **Остаток xlsx TextSim → задачи #54 (клиппинг ячеек — 2-й виновник), #55 (number-
  форматы), #57 (chart-title)**: матчить текстовый вывод Calc.
- CI-интеграция (прогон корпуса + таблица регрессий), масштаб, GovDocs1 scaling,
  font-agnostic visual + drift-по-зонам метрики.

**Приоритет если продолжать:** **xlsx TextSim** (новый binding-constraint, вскрыт
прогоном) — или **charts** (stacked/area/scatter, весомо). Остальное — по требованию.
*(TH/ColSpan/RowSpan #120/#121, прогон корпуса — ✅ сделано.)*

Валидация (следующий крупный блок — см. «Открытые вопросы»):
- **Корпус из ~300 реальных docx + ~100 xlsx** + визуальный diff vs LibreOffice
- **veraPDF** ✅ установлен (1.30.1, `~/verapdf/verapdf`) и интегрирован как opt-in
  gate (`tests/verapdf.test.ts`, env `VERAPDF`); -1a и -1b проходят формально

## Прогресс на момент handoff

**123 задачи закрыто. M0–M6 завершены** (M5 = shapes + charts + OfficeMathML; M6 = Tagged PDF + PDF/A-1/2/3 + **цифровые подписи** — все 8 профилей формально проходят veraPDF, подпись openssl-verified; **xlsx print-model завершён** — gridlines/print-area/fit-to-page/print-titles/breaks/centered). Core docx + xlsx → PDF работает с правильной семантикой:
- DOCX: styled runs (font/color/size/alignment), таблицы (gridSpan/vMerge/borders,
  row split при больших строках), списки (decimal/bullet/roman/letter с counter
  state), multi-section с per-section pgSz/pgMar/headers/footers,
  default/first/even header bands (titlePg + evenAndOddHeaders), inline images
  смешанные с текстом
- XLSX: workbook → все sheets с per-cell fonts/alignment/numFmt/shading/borders,
  merged cells, column widths, row heights, Excel дат (serial → calendar для 14-22,
  45-47 + custom format с m=month/minute disambiguation, 1900/1904 epoch),
  pageSetup (paperSize + orientation) + pageMargins из первого листа.
  **Print-model Excel:** подавление gridlines (по умолчанию, как Calc print),
  print area (_xlnm.Print_Area), fit-to-page/scale (shrink под fitToWidth),
  print titles (_xlnm.Print_Titles — повтор шапок), ручные rowBreaks,
  horizontalCentered — вместо «grid-as-table» теперь реальная print-модель
- Текст (M4): Knuth-Plass total-fit для абзацной вёрстки; опциональный Liang
  hyphenator (en-US / ru) для красивых justify-абзацев; OpenType GSUB
  ligatures (fi/fl/ffi) и GPOS pair kerning из bundled-шрифтов; BiDi (UAX #9)
  для смешанного RTL/LTR (иврит/арабский) — embedding levels, визуальный
  reorder, RTL right-align
- PDF: /Info метаданные (Title/Author/…) из docProps/core.xml; **PDF/A-1b**
  (архивный профиль) через `options.pdfA`
- Tagged PDF / PDF/A-1a (M6): логическая структура для accessibility —
  `/StructTreeRoot` → Document → H1–H6 / P / L→LI→LBody / Table→TR→TD / Figure;
  marked content (MCID) + `/ParentTree`, артефакты декораций и колонтитулов,
  alt-текст фигур, `/Lang`, `/MarkInfo`. `options.tagged` или `pdfA:'PDF/A-1a'`;
  при выключенном tagged вывод байт-в-байт прежний
- DrawingML shapes (M5): `wps:wsp` фигуры — 15 preset-геометрий + custGeom,
  solidFill / a:ln (dash/cap/width), поворот+flip, текст в фигуре (insets +
  вертикальный якорь), MCE (wps Choice > VML Fallback), theme-цвета (schemeClr
  → theme1.xml / Office-палитра); общий слой векторной графики (path/fill/
  stroke/bezier)
- DrawingML charts (M5): column/bar (clustered), line, pie из кэша chart1.xml;
  Heckbert nice-scale, оси/гридлайны/тики, подписи категорий, легенда (r/b),
  заголовок, цвета серий (fill/a:ln + accent-cycle), pie %-подписи; рендерятся
  целиком через общий векторный слой + text-pass
- OfficeMathML (M5): from-scratch box-model матнабор — дроби, степени/индексы,
  радикалы (нарисованный surd), n-арные (∑∏∫ вектором), функции, пределы,
  растяжимые разделители, матрицы, акценты; inline + display (центрирование);
  структурные элементы рисуются вектором, обычные символы — из шрифта (без
  матшрифта). Квадратное уравнение, ∑k², гауссов интеграл — рендерятся корректно

## Корпус-валидация (vs LibreOffice)

Харнесс реализован — `scripts/corpus/`: `npm run corpus:build` (синтетика),
`npm run corpus:fetch` (реальные docx/xlsx), `npm run corpus` (прогон).
Без npm-зависимостей: парсим PPM (P6) и stext XML от `mutool` сами.
- **Эталон:** `soffice --convert-to pdf`. **Растеризация/текст:** `mutool draw`
  (`-F pnm -c rgb` для пикселей, `-F stext` для координат текста).
- **Метрики на документ:** TextSim (LCS similarity текста, font-agnostic),
  Drift (медианный сдвиг baseline-y), Visual (доля несовпавших пикселей),
  Pages. Отчёт → `corpus/report.md`. Раннер принимает `CORPUS_DIR` (какой набор
  валидировать; по умолчанию `corpus/inputs`).
- **Setup:** LibreOffice + mupdf через brew; bundled Roboto скопирован в
  `~/Library/Fonts`, чтобы LO подставлял тот же шрифт (иначе визуальный diff
  зашумлён формой глифов).
- **Безопасность (untrusted-документы):** `CORPUS_SANDBOX=docker` включает
  изоляцию для непроверенного входа — (1) эталонный рендер LibreOffice идёт в
  залоченном контейнере (`docgen-losandbox`, образ собирается `npm run
  corpus:sandbox:build`): `--network none`, `--cap-drop ALL`,
  `--security-opt no-new-privileges`, `--read-only`+tmpfs, лимиты cpu/mem/pids,
  вход смонтирован read-only; (2) наш собственный парсер тоже запускается
  child-процессом (`convert-one.ts`) с wall-clock-таймаутом и heap-капом; (3) на
  уровне библиотеки `OpcPackage.open` капит распаковку (anti-zip-bomb). По
  умолчанию (для наших доверенных фикстур) — быстрый in-process путь.
- **Корпус:** `corpus/inputs/` — 8 синтетических документов (text/styled/
  justified/table/list/header-footer/2×xlsx, A4 + Roboto). `corpus/external/`
  (gitignored) — реальные файлы из `corpus:fetch` (Apache POI test-data,
  Apache-2.0; скачано 127 docx + 349 xlsx OOXML — .doc/.xls/.docm/.xlsm
  пропущены) + `manifest.json` (провенанс).
  `corpus/.work/` (промежуточные PDF/PPM) в .gitignore.

**Первый прогон (DPI 100):** core docx — отлично:
- text-basic / text-justified / text-styled / table-basic: TextSim **100%**,
  visual 0.1–2.9%, drift 1–5pt — рендер очень близок к LibreOffice.
- header-footer: drift 720pt — артефакт метрики (reading-order header↔body в
  stext), не баг рендера; TextSim 81.8% по той же причине.
- list-basic: 40 vs 31 символов (наши маркеры "1." + tab vs LO) — небольшое
  расхождение, стоит разобраться.
- xlsx: ожидаемое расхождение (мы рендерим grid-as-table + sheet title, LO Calc
  печатает по своей print-модели). **✅ print-model реализован (#45a–d) + прогон
  сделан: visual улучшился (73 файла), но ⚠️ гейтит TextSim, не visual — см. ниже.**

### Первый прогон реального POI-корпуса — docx (127, sandbox, DPI 100)

`CORPUS_DIR=corpus/external/poi-docx CORPUS_SANDBOX=docker CORPUS_ISOLATE_OURS=1
npm run corpus`. Отчёт → `corpus/report-docx.md`.

**127 docs — первый прогон ✅ 57 · ⚠️ 53 · ❌ 17 → после фиксов #110–#113
✅ 66 · ⚠️ 44 · ❌ 17** (+9 в ✅, 0 регрессий: page-break #112 +7, header/footer-only
#113 +2; tracked-changes #111 поднял TextSim). Корпус окупился
сразу: нашёл реальный баг робастности (неподдерживаемая/битая картинка валила
весь документ → исправлено, задача #109; 9 файлов ушли из ❌: 26→17).

**❌ 17 — все интенциональные safe-reject'ы враждебного/битого входа, НИ одного
реального бага конвертации:**
- 9× `invalid zip data` — зашифрованные (`bug53475-password-is-*`) и
  обрезанные/фуззерные архивы;
- 5× `OPC missing _rels/.rels` — битые OPC-пакеты (POI fuzzer);
- 1× `OPC archive rejected (zip-bomb guard)` — **сработал C1** на фуззер-бомбе;
- 1× `External entities are not supported` — **заблокирован XXE** (`ExternalEntityInText`);
- 1× `Maximum nested tags exceeded` — **депт-гард** (`deep-table-cell`).

То есть защитные меры (C1 zip-bomb, XXE-блок, кап вложенности) демонстрируемо
срабатывают на реальном adversarial-корпусе. (Под sandbox-изоляцией отчёт
показывает родовое «Command failed: npx tsx…» — настоящие сообщения видны в
in-process прогоне без `CORPUS_ISOLATE_OURS`.)

**⚠️ 53 — карта разобрана:**
1. **Page size/orientation (4) — РАЗОБРАНО, не баг парсинга.** `pgSz` при наличии
   доезжает до MediaBox: `IllustrativeCases` 828↔827 — ±1px округление, корректно.
   `Bug60341`/`MultipleBodyBug` вообще без `<w:pgSz>` (почти пустые 316–499 Б
   тест-стабы) → мы дефолтим A4, LO — Letter (US-локаль): доброкачественный
   дефолт-мисматч, A4-дефолт оставлен (глобально корректнее). `bug65649` —
   мультисекция → **#43 РАЗОБРАНО:** ориентация по секциям работает (репро
   portrait+landscape via `w:orient` корректен, залочен тестом); расхождение dims
   у `bug65649` — следствие page-count-дрейфа на 166-страничном/15-секционном
   монстре, не баг ориентации.
2. **Пустой/низкий текст.** ✅ `60329` (0/**4812**!) — **корень был: вложенные
   таблицы** (58 `w:tbl` table-in-cell; ячейка не рендерила вложенную таблицу) →
   **ИСПРАВЛЕНО (#116):** TextSim **0→96.3%** (0→4607 символов; docx-корпус
   перепрогнан — headline 66/44/17 без регрессий). ✅ **header/footer-only —
   ИСПРАВЛЕНО (#113):** `headerFooter` 0→54,
   `EmptyDocumentWithHeaderFooter` 0→12, `Headers` 0→9 (пустое тело теперь несёт
   колонтитулы). `Tika-792`/`61470` остались 0 — отдельная причина, не
   header/footer-only.
3. **Расхождение числа страниц (23).** ✅ **POI header/footer-сюита ИСПРАВЛЕНА (#112)**
   (`ThreeColHead`/`Foot`/`HeadFoot`, `SimpleHeadThreeColFoot`, `PageSpecificHeadFoot`,
   `FancyFoot`, `DiffFirstPageHeadFoot`, `HeaderFooterUnicode` — была стабильно 1 vs 2):
   причина не в колонтитулах, а в забытом `<w:br w:type="page">` (трактовался как `\n`)
   → теперь 2 страницы. Остальное (`Bug51170` 4≠7, `bug57031` 17≠14, `bib-chernigovka`
   12≠15, `drawing` 16≠20) — накопленный layout-drift на больших документах vs другой
   движок (ожидаемо).
4. **✅ Tracked changes — ИСПРАВЛЕНО (#111).** `delins.docx` **474 → 1326 символов**,
   TextSim **40%→84%** (эталон 1842; docx-корпус перепрогнан — headline 57/53/17 без
   регрессий). `w:ins`/`w:moveTo` добавлены в `RUN_CONTAINER_TAGS` (вставленный/
   перемещённый-в текст показывается); `w:del`/`w:moveFrom` опускаются =
   accept-all / финальный документ. Остаток до 1842 — удалённый текст, который LO
   рисует зачёркнутым (семантический выбор, не баг).

Прочее (TextSim 20–90% на коротких документах, мелкий drift) — расхождения
вёрстки/метрик, не блокеры. xlsx — см. следующую секцию.

### Первый прогон реального POI-корпуса — xlsx (349, sandbox, DPI 100)

`CORPUS_DIR=corpus/external/poi-xlsx …`. Отчёт → `corpus/report-xlsx.md`.

**349 docs — ✅ 18 · ⚠️ 303 · ❌ 28** (первый прогон) → **после фикса #110:
✅ 19 · ⚠️ 308 · ❌ 22** (6 `xlsx has no sheets` ушли в ✅/⚠️ с хорошей точностью
текста — `65016` 313990/289574, `craftonhills` 670/693, `style-alternate-content`
318/330; **0 регрессий**). Низкий ✅ ожидаем:
мы рендерим grid-as-table + заголовок листа, LO Calc печатает по своей
print-модели → почти всё «divergent» по визуалу/пейджингу, но **конвертируется**.

В отличие от docx, среди ❌ нашлись **реальные баги парсера** (корпус снова
окупился). Разбор 28 ❌ in-process (sandbox прячет причину в «Command failed»):
- **19 safe-reject'ов:** 15× fuzzer/crash (POIFuzzer/POIXSSFFuzzer/XLSX2CSVFuzzer,
  `crash-*`); `xlsx-corrupted`/`58616`/`deep-data`/`protected_passtika` →
  invalid zip / encrypted; `49609` → OPC missing `_rels/.rels`.
- **✅ 6× `xlsx has no sheets` — ИСПРАВЛЕНО (#110):** `58760`, `59021`,
  `59746_NoRowNums`, `65016`, `style-alternate-content`, `craftonhills.edu`.
  **Был корень:** файлы используют ЯВНЫЙ namespace-префикс `x:`
  (`<x:workbook>…<x:sheets><x:sheet r:id=…/>`), а наши SpreadsheetML-парсеры
  матчили теги по дефолтному (беспрефиксному) namespace → 0 листов. **Фикс:**
  `removeNSPrefix: true` во всех 4 spreadsheet-парсерах
  (workbook/worksheet/sharedStrings/styles) — срезает любой префикс с тегов и
  атрибутов (`r:id`→`id` подхвачен фолбэком); беспрефиксный общий случай не
  меняется. Все 6 теперь дают валидный `%PDF-1.7`. Регрессия в `tests/xlsx.test.ts`.
- **✅ 1× OOM — ИСПРАВЛЕНО (#114):** `CVLKRA-KYC` (204 КБ) рвал 512 МБ heap
  (`sheet15` = 49 194 пустых стилизованных ячейки до XFD). Сетка ограничена used
  range → конвертится (74 стр.).
- **✅ 1× parse-DoS — ИСПРАВЛЕНО (#115):** `poc-shared-strings` висел >60с (строка
  1 МБ × 12 000 ячеек). Кап строки 32 767 + бюджет текста листа → ~2с (76 стр.).
- 1× `xlsx-corrupted` — наш конвертер **проходит**; ❌ был от downstream
  soffice/mutool (не наш баг).

**⚠️ 303 divergent** — системное «grid-as-table vs Calc print-model» (другой
пейджинг/масштаб), не отдельные баги; точечный разбор — после фикса
namespace-префикса.

> **✅ ОБНОВЛЕНО — print-model реализован (#45a–d) + ПОВТОРНЫЙ ПРОГОН СДЕЛАН**
> (sandbox, 349 xlsx, `corpus/report-xlsx-printmodel.md`):
> **✅ 19 · ⚠️ 309 · ❌ 21** — headline почти не сдвинулся, НО это вскрыло, что
> **метрику гейтит TextSim, а НЕ visual** (важная коррекция прежней гипотезы):
> - **Visual улучшился на 73 файла, ухудшился на 0** (медиана 0.9%→0.6%); топ-выигрыши
>   драматичны: `49156` 32%→**8%** (пересёк visual-гейт!), `58896` 55%→32%,
>   `FormulaEvalTestData` 32%→22%, `45540_*` 20%→11%. Множество файлов с visual ≥10%
>   ушло ниже порога. **Print-model сработал по своей цели (визуальная структура).**
> - Но **264 из ⚠️/❌ теперь visual-clean (<10%) и валятся ТОЛЬКО по TextSim-гейту**
>   (медиана TextSim 69%); по visual-гейту валятся лишь 46. Т.е. для xlsx
>   **substitute-font НЕ floor'ит visual** (медиана <1%) — реальный binding-constraint
>   это **TextSim** (совпадение текста с Calc: number-форматы, заголовок/футер
>   «Page N» у Calc, reading-order, набор печатаемых ячеек).
> - **Следующий рычаг для xlsx ⚠️→✅** — не визуал (уже хорош), а **TextSim**: матчить
>   текстовый вывод Calc.
>
> **✅✅ ОБНОВЛЕНО (3-й прогон, #53 — дроп заголовка листа):** диагностика вскрыла, что
> **#1 TextSim-виновник — наш синтетический заголовок листа** (`bug65306`/`49872`/
> `57798`: лишний «Sheet1/2/3»; `xlsx-jdbc`: «Sheet_number_1»). Calc/Excel
> `--convert-to pdf` НЕ печатают имя листа НИГДЕ (даже в шапке — прежняя гипотеза была
> неверна). Дропнули заголовок (листы >1 ломают страницу пустым no-text параграфом).
> **Результат на полном корпусе: ✅ 19 → 163 · ⚠️ 309 → 165 · ❌ 21** — **145 файлов
> ⚠️→✅** (1 регрессия: `MalformedSSTCount` 95→89%, граничный malformed-SST). xlsx-корпус
> **5% → 47% clean**. (`corpus/report-xlsx-titledrop.md`.) Остаток TextSim → задачи
> #54 (клиппинг ячеек — 2-й виновник), #55 (number-форматы), #57 (chart-title).

### Шрифты и visual-метрика (важно для интерпретации)

Visual-mismatch у docx (~85% <10%, 65% <2%) **не равен** «качеству вёрстки» —
он завязан на ШРИФТ. Библиотека сама решает шрифт: async `convertDocxToPdf(docx)`
без `fonts` детектит семейство документа (`detectDocxFontFamily`) и тянет
substitute (sans→Roboto / serif→Tinos / mono→Cousine). Sync-путь с `fonts` —
опционален (офлайн/детерминизм).

**Находка (прогон `CORPUS_AUTOFONT=1`, флаг в `run.ts`):** переключение корпуса с
фикс-Roboto на реальную авто-подкачку **почти не двигает агрегат** (visual <10%:
93→94 из 110; `heading123` 10.1%→9.4% ⚠️→✅, но пара доков чуть хуже из-за
одно-семейного детекта). **Причина:** Docker-LibreOffice САМ подменяет шрифты
(только Liberation/DejaVu, без Cambria/Calibri документа). Значит метрика меряет
«наш substitute vs substitute LO», и остаточная разница (Tinos ≠ Liberation
Serif) **не закрывается выбором семейства**. Вывод: «пиксель-точно как LO» —
мираж (LO не истина, а другой подменщик); layout-движок при этом совпадает.

**Реальные рычаги fidelity (не «догнать LO»):** (а) ✅ **встраивание `w:embed`-шрифта
документа — РЕАЛИЗОВАНО (#118):** `word/fonts/*.odttf` де-обфусцируются и
используются напрямую → glyph-exact, без подмены (и мы, и LO берут ОДИН реальный
шрифт документа); (б) ✅ **per-run резолв шрифта — РЕАЛИЗОВАНО (#117):**
мульти-семейный registry, каждый текстовый ран резолвит своё семейство по `w:ascii`
(sans→roboto/serif→tinos/mono→cousine) → корректный ИНТЕНТ смешанных доков (раньше
было одно-семейно на документ); async `convertDocxToPdf` тянет набор на каждое
detected-семейство; single-family — байт-в-байт. Метрику стоит дополнить
font-agnostic визуалом (сравнивать формы/позиции, нормализовав шрифт).

### Следующие шаги по валидации
1. **Весь POI-корпус прогнан — docx (127) и xlsx (349)** (см. две секции выше).
   Осталось: закрывать находки по приоритету. docx: page size/orientation (топ-1),
   пустой/низкий текст (топ-2). xlsx: ✅ namespace-префикс `x:` исправлен (#110),
   остаются OOM (`CVLKRA-KYC`, #41) и parse-DoS (`poc-shared-strings`, #42).
   Инфраструктура проверена end-to-end (`corpus:fetch` + `CORPUS_SANDBOX=docker`).
   Реальный untrusted GovDocs1 — через тот же sandbox-путь (скачивание не
   автоматизировано).
2. **xlsx — Excel print-model (главный рычаг по 303 ⚠️).**
   ✅ **Слой 1 СДЕЛАН (#45a):** (а) **подавление gridlines** — синтетическая сетка
   больше не рисуется по умолчанию (только по `<printOptions gridLines="1">`); это
   универсальный сдвиг к golden — затронуты 323/333 листа (раньше grid лепился на
   все); (б) **print area** `_xlnm.Print_Area` клиппит рендер в диапазон (с офсетом
   старта); (в) распарсена вся остальная модель (scale/fitToWidth/fitToHeight,
   fitToPage, printOptions, row/colBreaks, definedNames). Проверено на 351 локальном
   POI xlsx (0 регрессий). См. #45a в Tasks.
   ✅ **fit-to-page масштаб СДЕЛАН (#45b):** `<pageSetup scale>` и fitToPage+fitToWidth
   → единый shrink-фактор на шрифты+высоты строк (auto-fit пакует уменьшенный текст
   без агрессивного переноса). Shrink-only, флор 10%, gated (нескейленные — байт-в-байт).
   ✅ **print titles / повтор header-строк СДЕЛАН (#45c):** рендерер повторяет лидирующие
   `isHeader`-строки наверху каждой страницы таблицы (`_xlnm.Print_Titles` → `parseTitleRowRange`
   → `isHeader`); **бонусом** docx `w:tblHeader` тоже теперь повторяется. Gated (нет header
   ИЛИ таблица влезает → байт-в-байт).
   ✅ **ручные page breaks + centered СДЕЛАНЫ (#45d):** `<rowBreaks>` → force-break-at-row
   (`RowProperties.pageBreakBefore` → `RowLayout.breakBefore`); `<printOptions horizontalCentered>`
   → `TableProperties.alignment` → `TableBlock.xOffsetPt` (сдвиг таблицы). Gated → байт-в-байт.
   ⏳ **Остаток (низкий приоритет):** **fitToHeight** (нужен точный суммарный rendered height —
   content-driven/layout-level; редкий кейс), vertical centered. Решение по служебному заголовку
   листа: **оставлен** (LO Calc по дефолтной page-style печатает имя листа в шапке — наш
   body-заголовок близок).

   **Итог print-model:** L1 (gridlines+print area+парсинг) + fit-to-page + print-titles/header-repeat
   + manual breaks + centered — **сделаны и протестированы** (40 xlsx-тестов, 0 регрессий на POI).
   Это системно сводит «grid-as-table vs Calc print-model» к настоящей print-модели Excel.
3. Улучшить метрику drift для headers/footers (сопоставлять по зонам, не по
   глобальному порядку строк).
4. CI-интеграция: прогон корпуса + публикация таблицы регрессий, трекинг
   покрытия спеки.
5. **veraPDF** — ✅ сделано. Установлен 1.30.1 (izpack-инсталлятор в console-
   режиме под Java 17; cask по-прежнему отсутствует). PDF/A-1a **и** -1b проходят
   формально. По пути найдены и исправлены 2 реальных бага conformance:
   §6.3.5 (отсутствовал `/CIDSet` в субсете) и §6.3.8 (ToUnicode не покрывал
   лигатурные глифы и TAB-глиф маркеров списков; заодно сбор глифов для субсета
   теперь шейпит текст как emit — раньше лигатурный глиф рендерился, но
   вырезался из субсета). Прогон встроен в `tests/verapdf.test.ts`.
