# Эпики — дорожная карта после 1.3.0

Три крупных направления из бэклога. Они РАЗНЫЕ по природе, и это главное при
планировании:

```
              читает              пишет
  docx  ──────────────►  FlowDoc  ──────────────►  PDF / SVG / HTML   (есть)
  xlsx  ──────────────►     │
                            ▼
  E-DOCX:   FlowDoc ──────────────────► .docx     (новый writer, обратный путь)
  E-PDF:    .pdf ──► реконструкция ──► FlowDoc     (новый reader, чужой IR)
  E-SHEET:  табличный IR-узел рядом с FlowDoc      (внутренний рефактор)
```

- **E-DOCX** — пятый writer на УЖЕ готовом IR (как svg/html). Самый «наш».
- **E-PDF** — reader из формата, который ничем не похож на OOXML. Самый дорогой.
- **E-SHEET** — смена внутреннего представления, не новый формат. Архитектурный долг.

Четвёртый эпик добавлен позже (после 1.8.0): **E-PPTX** — reader `.pptx` →
FlowDoc, слайд = страница; ложится на готовый IR (Route A). Секция и декомпозиция
PX0–PX6 — в конце документа, перед сводкой.

**Позиционирование Ream: в первую очередь docx + excel → pdf.** Excel —
ПЕРВОКЛАССНАЯ цель, не second-class. Это поднимает E-SHEET из «долга по запросу»
в стратегический эпик: текущая проекция «лист = документ с таблицей» — это ПОТОЛОК
для Excel-фич, а раз Excel первоклассен, потолок надо снимать.

**Рекомендуемый порядок: E-DOCX → E-SHEET → E-PDF.** docx-writer первым — лучшее
выгода/риск, замыкает bytes-in/bytes-out и даёт roundtrip-гейт, усиливающий ВСЮ
разработку. E-SHEET вторым как инвестиция в первоклассное Excel-направление
(разблокирует условное форматирование, спарклайны, фильтры, xlsx-writer). E-PDF —
отдельный крупный заход за «универсальностью», наименее связан с core-миссией.

Точки привязки в коде (проверено на 1.3.0):
- `Ream.convert(to, opts)` диспетчит по строке таргета (`ream.ts`, `ReamTarget`).
- `DocumentWriter<TDoc>` уже несёт `consumes: 'flow' | 'page'` (`ir/adapters.ts`).
- `FlowDoc.numbering`/`styles`/`headersFooters` подписаны «Raw definitions
  (round-trip material)» (`ir/flow.ts`) — IR проектировался под обратную запись.
- OPC-сборка (zip+content-types+rels) уже живёт в `tests/fixtures/build-docx.ts` —
  материал для подъёма в `src/`.

---

## E-DOCX — docx-writer (FlowDoc → .docx)

**Цель.** `Ream.parse(bytes).convert('docx')` → редактирование / нормализация /
санитайз docx в браузере, и docx→docx roundtrip-гейт.

**Усилие: среднее.** Писать OOXML легче, чем читать: мы сами выбираем подмножество
разметки, нет чужого многообразия, нет MCE-ветвлений.

### Главное архитектурное решение (принять ДО кода)

FlowDoc **лосси**: `body` несёт РЕЗОЛВЛЕННЫЕ пропсы (каскад схлопнут — решение
S6-3). Наивная запись body даст «плоский» валидный docx без стиле-ссылок —
тяжёлый, непохожий на вход, но корректный. Три варианта:

- **A (рекомендую для v1): денормализованный writer.** Пишем резолвленные пропсы
  как direct formatting на каждом ране/параграфе; `styles.xml` минимальный
  (docDefaults + Normal). Честно лосси, но валидно и просто. Roundtrip != байты,
  но IR-после-roundtrip эквивалентен (то, что реально проверяемо).
- **B: хранить direct-пропсы в IR отдельно от резолвленных.** Честный roundtrip,
  но трогает S6-стадию и раздувает IR. Большой эпик сам по себе — НЕ в v1.
- **C: round-trip-карман.** Reader складывает сырой `styles.xml`/`numbering.xml` в
  side-channel IR, writer переиздаёт его как есть. Уже частично есть (`numbering`
  raw). Прагматично для docx→docx, но не работает для xlsx→docx.

→ **v1 = вариант A.** Гейт roundtrip сравнивает IR(parse(write(IR))) ≈ IR, не байты.

**Статус: D1–D7 + T1–T4 ✅ (writer покрывает реальный документ; гейт практически чист).**
Roundtrip-гейт: 1100 корпус-доков, **ноль writer-падений**, **1099/1100 полная IR-идентичность**
(POI 110/110, LO 989/990). T1 — форматы картинок по магик-байтам (raster + EMF/WMF + встроенный
PDF). T2 — ридер парсит легаси VML-картинки (`<w:pict>/<w:object>` → `<v:imagedata r:id>`), 45
доков рендерят ранее невидимые VML/ActiveX/OLE-превью. T3 — writer сериализует DrawingML-шейпы
(preset/custom-геометрия, fill, line, текст-боди) → inline. T4 (хвосты) закрыл краевые ±1: collapse
картинки сквозь tracked-change/SDT-обёртки, round-trip разрыва страницы, sectPr на pPr блок-картинки,
сохранение пустой гиперссылки, VML-картинка только при наличии imagedata. Единственный остаток —
tdf115883: байты картинки физически вырезаны из пакета (dangling rel), переносить нечего. Остаточные
v1-лоссы: сноски/чарты/math не пишутся; шейп round-trip как inline (floating-позиция теряется).

### Декомпозиция (по нарастанию)

1. **D1 — OPC-writer в core.** Поднять zip-сборку из `build-docx.ts` в
   `src/core/opc/opc-writer.ts`: части → `[Content_Types].xml` + `_rels` + zip
   (fflate `zipSync`). Детерминизм (порядок частей, без дат в zip-заголовках).
   Гейт: собранный пакет читается обратно своим же `OpcPackage.open`.
2. **D2 — каркас writer'а + таргет.** `DocumentWriter<FlowDoc>` с `consumes:'flow'`,
   `ReamTarget += 'docx'`, ветка в `convert()`. Пустой документ (один параграф) →
   валидный минимальный docx. Гейт: roundtrip пустышки.
3. **D3 — body → w:p / w:r.** Параграфы, раны, резолвленные rPr/pPr (вар. A),
   разрывы, табы. Списки — маркеры уже материализованы в body, но для honest docx
   нужны w:numPr (numbering raw уже в IR → переиздать numbering.xml + ссылки).
4. **D4 — таблицы.** w:tbl/w:tr/w:tc, grid, спаны (gridSpan/vMerge), границы,
   шейдинг из резолвленных cell-пропсов.
5. **D5 — секции, HF, картинки.** sectPr (размер/поля/колонки), header/footer-парты
   + rels, w:drawing inline (картинки из ResourceStore → media-парты + rels).
   Сноски/гиперссылки/закладки — по образцу их reader-парсеров наоборот.
6. **D6 — roundtrip-гейт + корпус.** docx→IR→docx→IR, сравнение IR (нормализованно).
   Прогон по POI+LO корпусу: записать всё, что прочитали, без падений; диф IR.
7. **D7 — многосекционность.** Per-section sectPr: mid-doc встраивается в pPr
   закрывающего параграфа (body[endIndex-1]), финальная секция — body-child;
   каждая секция несёт свои HF (дедуп общих партов).

### Риски
- Лосси-IR (см. выше) — закрыто выбором варианта A + честной формулировкой гейта.
- Детерминизм zip (для байт-гейта своих фикстур) — fflate даёт, проверить mtime=0.
- Картинки/шрифты НЕ перевкладываем — пишем ссылки на ResourceStore-байты как есть.

---

## E-PDF — PDF-reader (PDF → FlowDoc)

**Цель.** `Ream.parse(pdf).convert('html'|'docx')` — извлечение текста, PDF→Office.
Замыкает Ream в универсальный документ-движок.

**Усилие: очень большое — фактически новая подсистема (½–1× от всего остального).**

### Почему трудно (честно)
PDF — это НЕ семантика, а инструкции рисования: «покажи глиф G в точке (x,y)».
В нём нет параграфов/таблиц/списков. Реконструкция FlowDoc из PDF — задача
OCR-уровня: глифы→слова (по зазорам), слова→строки (по baseline), строки→
параграфы (по leading/отступам), угадывание таблиц (по выравниванию колонок).

### Декомпозиция (крупными штрихами)
1. **P1 — COS-парсер.** Чтение объектов PDF: xref (классический + xref-streams),
   trailer, /Root, потоки + фильтры (FlateDecode есть в writer; нужны DCTDecode/
   и пр. на чтение). У нас есть ЗАПИСЬ объектов, но не ЧТЕНИЕ — это с нуля.
2. **P2 — content-stream интерпретатор.** Текстовые операторы (BT/ET/Tj/TJ/Tm/Td),
   графическое состояние (CTM), позиции глифов. ToUnicode/CMap → текст.
3. **P3 — Tagged-fast-path.** Если PDF тегирован (а МЫ такие пишем!) — структура
   берётся из /StructTreeRoot почти даром. Чужие PDF в массе НЕ тегированы → нужен
   и эвристический путь.
4. **P4 — эвристическая реконструкция.** Глифы→слова→строки→параграфы→(таблицы).
   Самый исследовательский кусок; качество = метрика, не бинарь.
5. **P5 — проекция в FlowDoc** + losses. Шрифты/картинки в ResourceStore.

### Зачем вообще
Tagged-PDF, который мы пишем, — идеальный вход для P3: roundtrip pdf(tagged)→FlowDoc
проверяем на СВОИХ же выходах. Это и есть честная отправная точка эпика, а чужие
нетегированные PDF — отдельная (большая) фаза.

### Прогресс
- **EP1 ✓ — COS-парсер.** Новая подсистема `src/pdf-reader/` (чистое дополнение, байт-в-ноль для
  всех существующих выходов). Переиспользует объектную модель писателя (`src/pdf/objects.ts`) — парс =
  инверсия сериализации. `lexer.ts`: токенайзер (числа/имена/строки литеральные+hex/`<<`/`>>`/`[`/`]`/
  keyword'ы) + `readStreamBody`. `parser.ts`: `parseObject` (грамматика значений с lookahead на `N G R`)
  + `parseIndirectObject`. `document.ts` `PdfFile`: классический `xref`+`trailer` (с цепочкой /Prev),
  `resolve(ref)` с кэшем, дерево страниц с наследованием MediaBox/Resources, декод FlateDecode (fflate
  `unzlibSync`), brute-force скан как recovery (битый xref / xref-stream). Подтверждено: наш писатель
  пишет КЛАССИЧЕСКИЙ xref (нет ObjStm/XRefStm) → reader читает свой же выход. Тесты: `pdf-reader-cos`
  (11, юниты грамматики), `pdf-reader-document` (4, roundtrip писатель→reader + реальный docx→pdf→read).
- **EP2 ✓ — content-stream интерпретатор + извлечение текста.** `content.ts` `interpretContent`: КА по
  операторам content-stream'а (q/Q/cm — CTM; BT/ET; Tf/Td/TD/Tm/T*/TL/Tc/Tw/Tz/Ts; Tj/TJ/'/"), трекинг
  text+line матриц, эмит ОДНОГО позиционированного run'а на show-оператор (TJ-скобка склеивается в один
  run в стартовом origin); координаты в page-space (Tm·CTM), эффективный кегль = Tfs·scale(матрицы);
  inline-image (BI…ID…EI) скипается. `cmap.ts`: парсер `/ToUnicode` CMap (codespacerange→1/2 байта,
  bfchar, bfrange hex+array). `font.ts`: `ContentFont` из /Font-словаря — decode через ToUnicode,
  advance из /Widths (simple) или /W+/DW (Type0/CID). `text.ts` `extractPageText`: строит font-map из
  /Resources/Font и гоняет интерпретатор. Honest e2e: текст ЧИТАЕТСЯ обратно из реального docx→pdf
  (в порядке чтения — первая строка выше второй). Тесты: `pdf-reader-content` (7), `pdf-reader-text` (4).
- **EP3 ✓ — tagged fast-path (StructTreeRoot → FlowDoc).** EP3a: интерпретатор трекает marked-content
  (BDC /MCID → push, /Artifact → none, EMC → pop; каждый run помечается mcid'ом — связь структуры с
  текстом). `struct-tree.ts` `readStructTree`: обход /StructTreeRoot → дерево /StructElem (резолв /S, /Pg,
  /K → вложенные элементы vs MCR/MCID + /A /ColSpan/RowSpan). `tagged.ts` `reconstructTaggedPdf`: склейка
  дерева с per-page MCID→текст → FlowDoc-body: H1–H6 → outlineLevel, P → параграф. EP3b: Table→TR→TH/TD →
  настоящий FlowDoc-Table (equal-width grid, colSpan, header-строки), LI → параграф (Lbl-маркер + LBody).
  Honest e2e: docx→tagged pdf→FlowDoc восстанавливает заголовки/параграфы/порядок чтения/уровни + таблицы
  (2×2 с текстом ячеек) + списки. undefined для нетегированного PDF. Тесты: `pdf-reader-tagged` (5).
- **EP4 ✓ — эвристика для нетегированных PDF.** `layout.ts` `reconstructByLayout`: ран'ы по общему baseline
  → строки (сорт по x, пробелы по зазорам с оценкой ширины half-em), строки → параграфы по вертикальному
  зазору (> 1.5× кегля), заголовки по кеглю выше медианы. `flow-build.ts` — общие `paragraphBlock`/
  `buildFlowDoc` (tagged.ts перевели на них). Качество = метрика, не бинарь. Тест: нетегированный docx→pdf
  → параграфы в порядке чтения + крупная строка как заголовок.
- **EP5 ✓ — pdfReader в фасаде.** `reader.ts` `pdfReader: DocumentReader<FlowDoc>`: sniff `%PDF-` (с
  допуском мусора в начале), read = PdfFile.parse → tagged ?? heuristic → FlowDoc + losses (degraded для
  нетегированного, dropped для картинок). Добавлен в `DEFAULT_READERS` (docx/xlsx/**pdf**), `SOURCE_MIME.pdf`.
  Теперь `Ream.parse(pdf).convert('html'|'docx')` РАБОТАЕТ. Байт-в-ноль (PDF сниффится отдельно от PK-ZIP,
  write-путь не тронут). Тесты `pdf-reader-facade`: sniff, format='pdf', text в interlayer, →html, →docx
  (валидный, репарсится в тот же текст), losses. **E-PDF замыкает Ream в универсальный документ-движок.**

- **EP6 ✓ — извлечение растровых картинок.** Интерпретатор (content.ts) теперь ловит оператор `Do` с
  текущим CTM и mcid → `ImagePlacement`. `images.ts` `collectPageImages` резолвит имя в `/Resources
  /XObject`, рекурсивно входит в Form-XObject'ы (композит `/Matrix`, ограничение глубины) и декодирует
  через `image-decode.ts`: JPEG (`/DCTDecode`) и JPEG2000 (`/JPXDecode`) — passthrough; всё остальное →
  raw-сэмплы → PNG (`png-encode.ts`, новый минимальный энкодер RFC 2083 + CRC32). Цветовые пространства
  DeviceGray/RGB/CMYK, CalGray/RGB, ICCBased (по `/N`), Indexed (по палитре); фильтры Flate/RunLength/
  ASCII85/ASCIIHex + PNG/TIFF-предикторы; bpc 1/2/4/8/16; `/SMask` → альфа PNG. Неподдержанное
  (ImageMask-стенсил, Separation/DeviceN/Lab, CCITT/JBIG2, LZW) → типизированный loss. Эмиссия: tagged-путь
  — кейс `Figure` (картинка по mcid + `/Alt`), плюс «осиротевшие» картинки в конец по позиции; heuristic-путь
  — интерливинг картинок с абзацами по верхней кромке (`y`). `reconstruct*` теперь возвращают `{doc,
  losses}`; reader сужает blanket-loss до «vector graphics not reconstructed». Байт-в-ноль (только чтение).
  Тесты `pdf-reader-images`: PNG-энкодер round-trip, Flate-RGB/DCT/Indexed/ImageMask декод, honest e2e
  docx(картинка)→pdf→parse→FlowDoc→html (tagged и untagged).

**Итог E-PDF.** Новая подсистема `src/pdf-reader/` (lexer/parser/document/content/cmap/font/text/
struct-tree/tagged/layout/flow-build/reader/images/image-decode/png-encode) — чистое дополнение, байт-в-ноль
для всех существующих выходов. PDF → FlowDoc: объекты (EP1) → текст (EP2) → структура (tagged EP3 /
эвристика EP4) → фасад (EP5) → растровые картинки (EP6). ~60 тестов, honest e2e на СВОИХ же выходах.

---

## 1.7.0 — хвосты (грайнд «доделаем всё»)

Десять опц. хвостов, каждый отдельным коммитом с полным гейтом, байт-в-ноль для писателей.

**PDF-чтение (E-PDF продолжение):**
- **EP7 ✓ — xref-streams + object-streams.** `document.ts`: `XrefEntry` (uncompressed/compressed); чтение
  `/Type/XRef` (W-поля/Index/типы 0/1/2) и `/Type/ObjStm` (resolve type-2 декодит ObjStm + парсит член).
  Гибридный `/XRefStm`, `/Prev`-цепочка. Brute-force тоже индексирует ObjStm + достаёт `/Catalog` изнутри.
  Предиктор-математика вынесена в `predictor.ts` (шарится с image-decode). Раньше объекты в ObjStm были НЕ
  читаемы → сжатые PDF теряли контент.
- **EP8 ✓ — `/Link`-аннотации → hrefs.** `extractPageText` собирает `/Link` (`/A /URI`), тегает раны по
  `/Rect`; `paragraphFromRuns` коалесцирует по href. Оба пути (tagged/heuristic) несут ссылки → html `<a>`/docx.
- **EP9 ✓ — шифрованные PDF (Standard handler, пустой user-пароль).** `crypto.ts` синх-примитивы (MD5/RC4/
  SHA-256/384/512/AES-CBC enc+dec, все по FIPS/RFC-векторам); `decrypt.ts` вывод ключа (Алг. 2 / 2.A / 2.B
  R6) + по-объектная дешифровка (RC4/AESV2/AESV3). honest e2e: docx→AES-256-pdf→читается обратно.
- **EP10 ✓ — залитый вектор (огранич.).** Интерпретатор ловит path (m/l/c/re/h) + fill-paint (f/B/…) + цвет
  (rg/g/k) → залитые пути; `vector.ts` фильтрует мусор (волоски, белое, фон), `shapeBlock` → custom-geom
  shape; heuristic-путь интерливит. Штрихи/градиенты/клипы — документированный loss.

**Полнота писателей:**
- **WT1 ✓ — xlsx embedded charts.** Шаренный `chart-serializer.ts` (инверсия chart-parser) → drawingN.xml +
  chartN.xml. Последний кусок сетки, который не round-trip'ился.
- **WT2 ✓ — docx сноски/endnotes.** runXml эмитит `w:footnoteReference`/`w:footnoteRef`; `emitNotes` пишет
  footnotes.xml/endnotes.xml (+ separator-стабы).
- **WT3 ✓ — docx чарты + OfficeMath.** Чарты через тот же chart-serializer (инлайн-drawing). `omml-serializer.ts`
  — инверсия omml-parser для ВСЕХ MathNode (дроби/скрипты/радикалы/nary/функции/пределы/делимитеры/матрицы/
  акценты/бары/groupChr/eqArr).

**xlsx-печать:**
- **SE-T ✓ — fitToWidth=N + бандинг.** fitToWidth=N>1 масштабирует колонки и бандит scaled-ширины на N страниц
  (раньше — одна переразмеренная таблица). («rowSpan через банды» оказался не-баг: вертикальный merge —
  фиксированная колонка, банды режут по колонкам.)

**Итог 1.7.0.** PDF-чтение стало промышленным (сжатые/шифрованные/ссылки/картинки/залитый вектор); docx/xlsx
round-trip — полный (сноски/чарты/math/embedded-charts). Новые модули `src/pdf-reader/{predictor,crypto,
decrypt,vector}.ts`, `src/core/drawingml/chart-serializer.ts`, `src/word/omml-serializer.ts`. 745 тестов.

## 1.8.0 — хвосты (грайнд «доделай все хвосты»)

Восемь хвостов, каждый отдельным коммитом с полным гейтом, байт-в-ноль для писателей.

- **EP13 ✓ — текст в Form XObject.** `extractPageText` рекурсивно входит в формы (`/Name Do`), компонуя
  `/Matrix` и собственные шрифты формы; текст из form-wrapped тел восстанавливается на обоих путях.
- **EP14 ✓ — шифрованный PDF с НЕпустым user-паролем.** `Ream.parse(bytes, { password })` → `decrypt.ts`
  падит пароль (R2-4) / хэширует UTF-8 (R6) в деривацию ключа; пустая строка по-прежнему открывает
  permissions-only.
- **EP11 ✓ — штриховой вектор → линии.** Интерпретатор ловит RG/G/K + w + S/s/B/b; путь несёт
  `strokeHex`/`lineWidth`; `shapeBlock` → `ShapeLine` (fill:none).
- **EP12 ✓ — LZW-картинки.** PDF/TIFF LZW (9→12 бит, clear/EOD, KwKwK, `/EarlyChange`) в `image-decode.ts`,
  рядом с Flate; предиктор после.
- **EP15 ✓ — CCITT-факс.** Новый `src/pdf-reader/ccitt.ts` — T.4/T.6 с нуля (run-таблицы + 2D-моды),
  Group 4 + Group 3 1-D → DeviceGray.
- **EP16 ✓ — градиенты первокласснее.** Модель `ShapeGradient` (core/vector); docx-reader парсит `a:gradFill`
  в стопы (не усреднение); SVG/HTML/PDF (axial/radial shading pattern, `src/pdf/shading.ts`) рендерят
  достоверно; docx round-trip; чтение PDF лифтит shading-pattern обратно в градиент-заливку
  (`src/pdf-reader/shading.ts`, /Function 2/3/0). PDF/A держит solid-fallback.
- **EP17 ✓ — двухколоночная реконструкция.** Untagged-путь делит страницу по центральному жёлобу и читает
  колонка-за-колонкой; консервативно (полноширинная строка пересекает центр → нет деления), без регрессий.

**Итог 1.8.0.** PDF-чтение почти полное: form-текст, реальные пароли, штрихи, LZW/CCITT-картинки, градиенты
(emit+read), двухколоночный layout. Новые модули `src/pdf/shading.ts`, `src/pdf-reader/{ccitt,shading}.ts`.
**Всё ещё отложено (deep/низкая ценность):** JBIG2, голый `sh`-оператор (нужен clip-трекинг), tables-by-
alignment-эвристика. `.doc`/`.xls` — OFF-LIMITS (подтверждено). 771 тест.

---

## E-SHEET — табличный IR-узел (стратегический: Excel первоклассен)

**Цель.** SheetDoc рядом с FlowDoc: ячейки с типами/формулами, замороженные панели,
диапазоны печати — БЕЗ притворства «лист = документ с таблицей».

**Стратегический эпик, не просто долг.** Ream позиционируется как docx + excel →
pdf; Excel — первоклассная цель. Текущая проекция «лист = таблица в документе» —
это потолок для Excel-фич, и снятие его инвестирует прямо в core-миссию.

**Усилие: большое, byteRisk высокий.** Сама по себе user-фичи не даёт, но
РАЗБЛОКИРУЕТ ветку первоклассных Excel-фич (условное форматирование, спарклайны,
фильтры, осмысленный xlsx-writer), которые на текущей проекции делать всё тяжелее.

### Сейчас
xlsx-reader проецирует лист в `BodyElement[]` (таблица + чарты), переиспользуя
docx-pipeline (print-model.ts). Прагматично и работает, но лист — не текст.

### Что меняется
- Новый IR-узел `SheetDoc` (ячейки, столбцы/строки, merge, print setup, drawings).
- xlsx-reader → SheetDoc напрямую, без проекции в BodyElement.
- Layout: либо SheetDoc→PageDoc свой путь, либо SheetDoc→FlowDoc-адаптер у границы
  PDF (сохранить текущий рендер, но за чистой границей).

### Предусловие
Полный корпус-контроль (POI 349 + LO 293 xlsx) на КАЖДОМ шаге — byteRisk высокий.
Подход как с A9 PageAssembler: дословный перенос текущего поведения за чистую
границу СНАЧАЛА (байт-нулевой), и только потом новые Excel-фичи поверх.

### Точки привязки в коде (проверено на 1.4.0)
Текущий пайплайн (`src/excel/`):
```
xlsx → parseWorkbook (sheets, date1904, definedNames)
     → parseSharedStrings + parseXlsxStyles (XlsxStyles: cellXfs/fonts/fills/borders/numFmts)
     → parseWorksheet → ParsedWorksheet   (грид-модель: cells/columns/merges/rowHeights/
                                            pageSetup/printOptions/breaks/drawingRelId)
     → worksheetToBody(ws, sharedStrings, styles, …) → BodyElement[] (Table)   ← ПРОЕКЦИЯ
     → parseSheetDrawing → chart-блоки
     → FlowDoc { body, section, charts, info } → renderStyledPdf (общий docx-путь)
```
- `xlsx-reader.ts` САМ пишет в шапке: «SheetDoc deliberately deferred tech debt». Это закрытие долга.
- `ParsedWorksheet` (`worksheet-parser.ts`) — уже почти грид-узел; не хватает workbook-обёртки
  (несколько листов + общие sharedStrings/styles/definedNames/theme) и резолва значений.
- `worksheetToBody` (`print-model.ts`, ~300 строк) — вся Excel-семантика: `resolveCellText`
  (sharedStrings+numFmt), `runPropsFromXf`/`shadingFromXf`/`bordersFromXf`, merges, print-scale.
  **Это и есть будущий SheetDoc→FlowDoc адаптер** — переносим как есть.
- Гейт xlsx — sandbox LO TextSim/geometry (`scripts/corpus/run.ts`), similarity, НЕ байт-identity.

### Архитектурное решение (принять ДО кода)
Где живёт SheetDoc и как рендерится:
- **A: SheetDoc → PageDoc, свой грид-layout** — замороженные панели, истинная постраничная
  нарезка колонок, грид-пагинация. Снимает потолок рендера, но это БОЛЬШОЙ отдельный движок.
- **B (рекомендую для v1): SheetDoc → FlowDoc-адаптер у границы.** `worksheetToBody` уже ЕСТЬ
  этот адаптер — переносим за чистую границу без изменений, рендер байт-в-ноль. SheetDoc-узел
  разблокирует фичи; путь рендера прежний. Вариант A — поздняя отдельная инвестиция.

→ **v1 = вариант B.** Узел важнее рендера: фичи цепляются к SheetDoc, проекция их отрисовывает.

### Декомпозиция (волнами; байт-гейт на каждом шаге)

**Волна A — герметичная граница SheetDoc (байт-в-ноль; рефактор, ради которого всё).**
- **SA0 — xlsx-байт-гейт.** Снапшоты PDF-байт (или FlowDoc) на N репрезентативных xlsx-фикстурах
  (расширить OOP-0) — сеть безопасности рефактора. Без неё «байт-в-ноль» недоказуем.
- **SA1 — типы SheetDoc в core.** `src/core/ir/sheet.ts`: `SheetDoc(kind:'sheet') = { sheets,
  styles, definedNames, resources, charts?, info?, section? }`; `Sheet = { name, cells, columns,
  rowHeights, merges, dims, pageSetup, printOptions, breaks, drawings }`; `Cell = { row, col,
  type, rawValue, styleIndex, formula? }`. Решение: ячейка несёт СЫРОЕ значение + styleIndex
  (резолв остаётся в проекции → байт-в-ноль тривиален). Только типы, без проводки.
- **SA2 — ридер строит SheetDoc; проекция = SheetDoc→FlowDoc-адаптер.** Перенести `worksheetToBody`
  + оркестровку (page-break между листами, чарты, section) в `sheet-to-flow.ts`, потребляющий
  SheetDoc. `readXlsx` строит SheetDoc, затем `projectSheetDoc(sheet) → FlowDoc`; `xlsxReader`
  пока `produces:'flow'` (адаптер внутри). **Критический шаг: байт-identity PDF на всём xlsx-корпусе.**

**Волна B — SheetDoc первоклассен у границы (байт-в-ноль).**
- **SB1 — `xlsxReader produces:'sheet'` + проекция у фасада.** `convert('pdf'|'svg'|'html')` гонит
  `projectSheetDoc` → flow-путь. `Ream.parse(xlsx).sheet` отдаёт workbook для инспекции (как `.flow`).
  SheetDoc становится реальным выходом ридера, FlowDoc — производный вид. (Можно слить с SA2.)

**Волна C — первые Excel-фичи на SheetDoc (выплата за рефактор; были «в потолок»).**

- **SC1 ✓ — условное форматирование (cellIs).** Готово (коммит `48d0224`). `<conditionalFormatting>`/
  `<cfRule type="cellIs">` → dxf-оверрайд (fill/font) per-cell по SheetDoc. `conditional-format.ts`
  (`buildConditionalFormatter`), `print-model.ts` cell-loop hook, `styles-parser.parseDxfs`.
- **SC1b ✓ — colorScale.** Готово (коммит `6f52bb8`). 2/3-стоповый градиент: порог из экстента
  значений диапазона (`collectRangeValues` + `resolveCfvo`: min/max/num/percent/percentile),
  интерполяция в RGB. Ровно «cross-cell min/max», который грид-узел даёт, а таблица-проекция — нет.

**Общий фундамент остатка волны C — «cell-decoration»: per-cell векторный/бар-оверлей.**
dataBar, iconSet, sparkline и autofilter-кнопка — все суть «нарисуй графику в финальном rect ячейки».
Этот rect известен в `emitRowChunk` (`styled-layout.ts:3534`, где сейчас рождается `FillItem` заливки:
`cellX`/`cellWidth`/`rowHeight`). Туда добавляется (а) узкий `FillItem` для бара и (б) `ShapeItem`
(`VectorShape`) для иконок/спарклайнов — рисуются тем же `emitVectorShape` (`pdf/vector-graphics.ts`),
что и чарты. Слой вводит **SC1c** (первый потребитель — dataBar), переиспользуют **SC2** и **SC3**.
Байт-в-ноль везде: поля опциональны, ячейка без декорации идёт прежним путём → снапшоты не двигаются.

- **SC1c ✓ — dataBar + iconSet** (`0eaf8f6`, `a464a6d`). Два cfRule-типа, рисующие графику, а не заливку.
  - *Парс* (`worksheet-parser.parseCfRule` — диспетчер по `@type`): `dataBar` (`<dataBar><cfvo min/>
    <cfvo max/><color/></dataBar>` + minLength/maxLength/showValue/gradient), `iconSet`
    (`<iconSet iconSet="3TrafficLights1"><cfvo/>×N`, N порогов). Модель `CfRuleDataBar`/`CfRuleIconSet`
    в spreadsheet-model; union `CfRule` ширится.
  - *Эвалюатор* (`conditional-format.ts`): `CfOverride` += `dataBar?{fraction,colorHex}` и
    `icon?{set,index}`. fraction = (value−min)/(max−min) по той же `collectRangeValues`, что у colorScale;
    icon-index — по порогам cfvo.
  - *Рендер dataBar* (Strategy A): `CellProperties.dataBar?` → `CellLayout` → `emitRowChunk` кладёт
    `FillItem` шириной `cellWidth×fraction` ПОВЕРХ заливки, ПОД текстом. Width FillItem уже честен
    (`styled-page-emitter.ts:568`) → PDF без правок. HTML — linear-gradient. SVG заливку не красит → скип.
  - *Рендер iconSet*: встроенная вектор-геометрия иконок (кружки светофора / стрелки-треугольники /
    флаги) через `PathBuilder` → `VectorShape` у левого края ячейки + левый text-inset (cell margin),
    чтобы текст не наезжал.
  - *Тесты*: fraction→ширина бара; value→индекс иконки; приоритет; экстент. Опц. байт-гейт-фикстуры.
  - *Хвост*: dataBar с осью (отриц. значения), gradient-вариант, `w:dataBar` в docx-writer — отложить.

- **SC2 ✓ — спарклайны** (`6af8ba8`). Мини line/column/winLoss в ячейке. Строится НА фундаменте SC1c.
  - *Парс*: `<x14:sparklineGroups>` в worksheet `extLst` (`removeNSPrefix:true` снимает `x14:`/`xm:` —
    `worksheet-parser.ts:43`). `parseSparklines(wsObj)` рядом с `conditionalFormats` (`:73`) →
    `ParsedWorksheet.sparklines`; тип `ParsedSparkline{type,dataRange,sqref,colorHex?,lineWeight?}`.
  - *Резолв значений*: data-range `<xm:f>` может быть на ДРУГОМ листе (`Sheet1!$C$1:$C$10`) → резолв в
    `sheet-to-flow.ts` (там SheetDoc со ВСЕМИ листами), не в per-sheet `print-model`. Значения ЖИВЫЕ из
    грида (не cache).
  - *Геометрия*: новый `src/core/drawingml/sparkline-geometry.ts` — `buildSparkline(type,values,wPt,hPt)
    → VectorShape[]` через `PathBuilder` (polyline для line; rects для column; +/− rects для winLoss).
    Без осей/легенды/лейблов — этим лёгкий относительно `ChartScene`.
  - *Размещение*: заполняет ОДНУ ячейку → cell-decoration-слой SC1c. Размер ячейки известен
    (`columnWidths`/`rowHeightMap` — `print-model.ts:341/293`).
  - *Тесты*: геометрия (N точек → N−1 сегментов; знак winLoss); размещение; cross-sheet-резолв диапазона.

- **SC3 ✓ — Excel-таблицы** (`071a1af`; банды + header-шейдинг, autofilter-кнопка в хвосте). `<tableParts>` → `xl/tables/tableN.xml`.
  - *Парс/проводка*: worksheet-parser снимает `tablePart` rId'ы (как `drawingRelId`, `:68`); xlsx-reader
    РЕЗОЛВИТ rel'ы (pkg+rels уже есть, паттерн drawing — `xlsx-reader.ts:75`) → новый
    `src/excel/table-parser.ts` `parseTablePart` → на `ParsedWorksheet/SheetDoc`. Тип
    `ExcelTablePart{ref,name,styleName,columns,autoFilter?,showRowStripes,showFirstColumn,...}`.
  - *Стили*: built-in `TableStyleMedium2`&co НЕ в файле → хардкод-рецепт name→(header=accentN,
    band2=accentN·tint0.8, border=accentN) поверх темы воркбука. Резолвер темы уже есть
    (`buildXlsxColorResolver` — `xlsx-reader.ts:120`), переиспользовать.
  - *Применение*: в `worksheetToBody` оверлей `CellShading` на ячейки в `table.ref` (паритет строки →
    band1/band2; header-row → header-цвет; first/last-col-флаги) — паттерн precomputed per-cell lookup,
    как `cfFormatter`. Бэндинг-движок docx (`style-cascade/table.ts`) концептуально переиспользуем,
    но проще inline в cell-loop.
  - *autoFilter*: дропдаун-кнопка в header-ячейке = маленький треугольник → iconSet-вектор-слой SC1c.
    Опц./хвост.
  - *Тесты*: парс table.xml; бэндинг (паритет→цвет); header-шейдинг; styleName→accent-резолв по теме.

**Порядок волны C: SC1c → SC2 → SC3.** SC1c вводит cell-decoration + бар-слой; SC2 (спарклайн заполняет
ячейку) и SC3 (autofilter-кнопка) переиспользуют его. Каждый шаг — байт-в-ноль (фичечит только файлы,
использующие фичу); снапшоты не двигаются, как в SC1/SC1b.

**Волна D — xlsx-writer (симметрия, аналог E-DOCX; «осмысленный writer» возможен ТОЛЬКО на SheetDoc).**
- **SD1 ✓ — xlsx-writer** (`691e1ca`). `writeXlsx(SheetDoc)` → workbook.xml + sheetN.xml + sharedStrings
  + styles через core OPC-writer. `DocumentWriter.consumes:'sheet'`; `convert('xlsx')` у Ream и фасада
  (отвергает flow-вход). Пишет ядро грида: ячейки (t/s/v), строки/колонки/высоты, merges, dimension,
  стили (numFmts/fonts/fills/borders/cellXfs/dxfs).
- **SD2 ✓ — roundtrip-гейт** (`f8ab727`). xlsx→SheetDoc→xlsx→SheetDoc: IR-идентичность written-surface +
  байт-стабильность (`b2===b1`, детерминированный fixpoint). 25 фикстур (после SD3).
- **SD3 ✓ — паритет writer'а.** SD3a (`86c82ac`): page setup / print options / breaks. SD3b
  (`ab0dc50`+`adc0c24`): условное форматирование, спарклайны (extLst), table-парты (новые tableN.xml +
  rels). Весь грид-surface теперь round-trip'ит (IR-identity + byte-stable); не пишутся только embedded
  charts (reported as loss). SD3c (`781a866`): корпусный roundtrip-гейт (`corpus:roundtrip:xlsx`) —
  **331 readable → 331 identical, 0 divergent, 0 writer-failed** на poi-xlsx; нашёл и починил reader-баг
  (prefixed `.rels` `<ns0:Relationship>` → 0 листов; removeNSPrefix в parseRelationships).

**Хвосты render-полировки волны C (видимое качество Excel→PDF).** TC1 (`0189ac8`): per-style accent-цвета
таблиц + белый header-текст. TC2 (`d6347e7`): верные icon-глифы (3Signs diamond/triangle/circle, *Gray
монохром). TC3 (`04aae07`): cross-sheet спарклайны + gap для пустых ячеек. TC4 (`06073a0`): dataBar с осью
для отрицательных. TC5: верные глифы оставшихся семейств — symbols (check/exclamation/cross на
рампе), ratings (монохромный bars-метр, `filled = bucket+1` из `count`) и quarters (часовой pie,
`filled` из `count-1` секторов). Новые `CellIconShape` (check/cross/exclamation/bars/pie) + опц.
`CellIcon.fill{filled,levels}`; рендерятся в PDF (мульти-прим в `buildCellIconShape`) и HTML (inline-SVG).
Автофильтр-кнопка — осознанный non-goal: Excel прячет её в print/PDF, рисовать нечего.

**Волна E — грид-пагинация через проекцию (прагматичный «вариант A»).** Не отдельный SheetDoc→PageDoc
движок (нулевой видимый выигрыш сверх проекции, высокий риск регрессий), а расширение проекции с
переиспользованием проверенного layout-движка.
- **SE1 ✓ — column-band пагинация широких листов.** Раньше `computeColumnWidths` (styled-layout) РАВНОМЕРНО
  сжимал лист шире страницы в одну страничную ширину; теперь, если лист не масштабируется (не fit-to-page и
  не явный `<scale>`) и шире печатной области ИЛИ несёт ручной `<colBreaks>`, проекция режет его на
  колоночные банды (`src/excel/column-bands.ts` `computeColumnBands` — жадная упаковка + границы по
  colBreaks) и эмитит ПО ОДНОЙ таблице на банду; банды 2+ получают `pageBreakBefore` на первой строке →
  «down, then over». Горизонтальный мёрж через границу банды клиппится к стартовой банде (дальше — пусто).
  Print-titles строки повторяются на каждой банде (через `isHeader`). Байт-в-ноль для всех текущих фикстур
  (узкие / fit-to-page не триггерят); новые байт-гейт кейсы `column-bands`, `column-breaks`. Тесты:
  `tests/column-bands.test.ts` — мат-юнит + e2e (3 банды → 3 страницы; colBreak → 2; fit-to-page → 1 банда).
- **SE2 ✓ — фрозен-панели в IR.** `<sheetView><pane state="frozen|frozenSplit">` →
  `ParsedWorksheet.pane{frozenRows=ySplit, frozenCols=xSplit}` (plain "split" игнорится); xlsx-writer
  пишет `<sheetViews><pane>` обратно (roundtrip-гейт это покрывает — `normGrid.pane`). Без эффекта на PDF
  (в Excel заморозка — view-настройка, не печатается; печатный повтор — Print_Titles). Байт-в-ноль для PDF.
- **SE3 ✓ — HTML sticky-панели.** Проекция вешает `TableProperties.frozen{rows,cols}` на одиночную таблицу
  (не на банды); html-writer эмитит `position:sticky` + накопленные `top`/`left` офсеты (left точный из
  grid-ширин, top из row-height с дефолтом 15pt) + z-index (угол=3, верх=2, лево=1) + opaque-фон. PDF/SVG/
  docx игнорят поле (байт-в-ноль). Тесты `tests/frozen-panes.test.ts` (парс + roundtrip + HTML sticky).

### Риски
- Байт-нулевой SA2 — главный риск; вся опасность волны сконцентрирована в нём (сеть = SA0).
- SheetDoc-«узел vs резолв»: держим сырое значение + styleIndex, резолв в проекции — иначе SA2 не нулевой.
- Дублирование XlsxStyles ↔ резолв: проекция остаётся единственным местом резолва стилей листа.

### Рекомендованный старт
**SA0 → SA1 → SA2** (герметичная граница). До SheetDoc ничего не разблокировано; SA2 — высокорисковый
байт-в-ноль шаг, делать под полным корпус-контролем. Дальше SB1 (экспозиция), затем волна C (фичи).

---

## E-PPTX — pptx-reader (PresentationML → FlowDoc)

**Цель.** `Ream.parse(pptx).convert('pdf'|'svg'|'html'|'docx')` — слайды презентации
как страницы. Четвёртый вход (после docx/xlsx/pdf), замыкающий офисную OOXML-тройку
на чтение. «Новая эра» после 1.8.0.

**Усилие: среднее.** Не новая подсистема рендера и не новый IR — pptx ложится на
УЖЕ готовый FlowDoc и весь конвейер layout→PDF/SVG/HTML. Стоимость — только
pptx-фронт (парс презентации/слайдов/каскада layout↔master) + переиспользование
DrawingML, которого после M5/E-DOCX уже много (shape/chart/theme/preset-geometry).

### Почему ложится чисто (Route A)
Слайд PresentationML — это позиционированный холст: shape'ы с абсолютными `a:xfrm`
(`a:off`+`a:ext` в EMU). Это ТОЧНО модель плавающего docx-drawing'а (`FloatAnchor`
wrap:'none', relativeFrom:'page', offsetPt). Значит:
- **слайд = секция** размером со слайд (`SectionProperties.pageSize` = `p:sldSz`),
- **shape слайда = floating `ShapeBlock`/`ChartBlock`/`ImageBlock`** на позиции `a:xfrm`,
- текст в shape — `ShapeTextBody` (уже есть, M5.7),
- DrawingML (геометрия/заливка/линия/градиент/тема), `chart-parser`, картинки,
  таблицы — переиспользуются КАК ЕСТЬ.

### Главное архитектурное решение (принято ДО кода): Route A
- **A (выбран): pptx → FlowDoc, слайд = секция + floating-shape'ы.** Ноль нового IR,
  ноль нового рендера. Весь pipeline (layout/PDF/SVG/HTML/даже docx-writer) работает
  над pptx бесплатно. Цена — «холст» выражается через float-якоря (выразительно
  достаточно: relativeFrom:'page' = абсолют от края страницы).
- **B (отвергнут): отдельный `SlideDoc` IR.** Дублировал бы всю shape-модель
  (геометрия/заливка/текст-боди/чарты), которая уже живёт в document-model, и
  потребовал бы свой layout + проекции во все writer'ы. Большой эпик без выигрыша —
  слайд И ТАК выражается холстом float'ов.
- **C (отвергнут): pptx-reader `produces:'page'` напрямую в PageDoc.** Обошёл бы
  FlowDoc, но тогда html-writer (потребляет FlowDoc) и docx-выход отвалились бы, а
  пагинацию/измерение пришлось бы делать руками. FlowDoc-путь даёт все 4 выхода даром.

→ **v1 = вариант A.** Слайд проецируется на секцию FlowDoc, контент — floating
BodyElement'ы. Тот же приём «узел важнее рендера», что в E-SHEET (там вариант B).

### Точки привязки в коде (проверено на 1.8.0)
- `FloatAnchor` (`types.ts:748`): `posH/posV.relativeFrom:'page'` + `offsetPt:Pt` —
  абсолют от края страницы. `a:off` EMU / 12700 → offsetPt; `a:ext` → width/height.
- `ShapeBlock` (`types.ts:635`): `float?`, `width/height:Pt`, `geometry/fill/line`,
  `text?:ShapeTextBody`, `paragraphProperties`, `altText?`. `ShapeTextBody`
  (`types.ts:626`): `content:BodyElement[]` + insets + `anchor:'t'|'ctr'|'b'`.
  `BodyElement` несёт `{kind:'shape', shape}` (`types.ts:703`) — floating shape
  кладётся прямо в body.
- `SectionProperties.pageSize` (`types.ts:729`) = размер слайда из `p:sldSz` cx/cy.
- Рычаг пагинации: `styled-layout.ts:3274` — `if (pageBreakBefore &&
  asm.current.length>0) flushPage()`. Пустой параграф схлопывается → страница не
  родится; **каждый слайд несёт in-flow якорь** (один ран с U+200B), чтобы страница
  эмитилась и floating-shape'ы садились на неё. pageBreakBefore на якоре i>0 открывает
  страницу слайда i.
- docx shape-парсер (`drawing-parser.ts` wsp/spPr/txbxContent → ShapeData;
  `document-parser.ts:421` → BodyElement kind:'shape') — образец сборки floating-shape'а;
  pptx-фронт строит то же из `p:sp`.
- `chart-parser.ts:52` `parseChart(chartXml, resolveColor)`, `theme-parser.ts`+`colors.ts`
  ColorResolver — переиспользуются в PX4/PX5.
- EMU_PER_PT = 12700; 16:9 дека = 12192000×6858000 EMU = 960×540pt; 4:3 =
  9144000×6858000 = 720×540pt.

### Декомпозиция (вертикальными срезами; байт-в-ноль для всех существующих выходов)
pptx сниффится отдельно (ZIP с `ppt/presentation.xml`), write-пути не тронуты → каждый
шаг байт-в-ноль для docx/xlsx/pdf.

- **PX0 ✓ — шов ридера** (`0726c59`). `src/pptx/pptx-reader.ts` `pptxReader:
  DocumentReader<FlowDoc>`: sniff `PK` + `ppt/presentation.xml`; read = OPC → `p:sldSz`
  (cx/cy) + число слайдов (`p:sldIdLst/p:sldId@r:id` ↔ `/slide`-rel'ы презентации) →
  одна страница на слайд размером со слайд. Зарегистрирован в `DEFAULT_READERS`.
  Фикстура `tests/fixtures/build-pptx.ts` (`buildPptx(slides[], {cx,cy})`). Тесты
  `tests/pptx-reader.test.ts` (3): sniff→'pptx'; 16:9×3 → 3 страницы 960×540; 4:3 → 720×540.
- **PX1 ✓ — текст слайда** (`d76f634`). `src/pptx/slide-parser.ts` `parseSlideShapes`:
  обходит `p:cSld/p:spTree`; `p:sp` со СВОИМ `a:xfrm` и `p:txBody` → floating `ShapeBlock`
  (geometry:rect, fill:none) с `ShapeTextBody` на позиции off/ext (`relativeFrom:'page'` =
  абсолют от края слайда); параграфы из `a:p`, раны из `a:r`/`a:fld`/`a:t`, прямой `a:rPr`
  (кегль `sz/100`, b/i/u, цвет `a:solidFill/a:srgbClr`, шрифт `a:latin`). Ридер резолвит
  каждый `p:sldId` в часть слайда по порядку и парсит дерево на страницу слайда. Рендерит
  существующий конвейер (текст-боди M5.7 + float-разметка W5d) без правок. Тесты
  `tests/pptx-slide.test.ts` (4): текст в HTML; два слайда — по странице; бокс 2in×1.5in →
  текст на нужной точке PDF (x≈144pt, у верха); жирный `a:rPr` → weighted.
- **PX2 ✓ — плейсхолдеры (каскад slideLayout→slideMaster)** (`f737119`).
  `src/pptx/placeholder-cascade.ts`: shape-плейсхолдер (`p:ph`) без своего `a:xfrm`
  наследует (а) геометрию из совпадающего прототипа layout'а, иначе мастера (матч по
  idx → type → категории стиля, так что ctrTitle ↔ title всё равно сходятся), (б)
  размер/цвет текста из мастер-`p:txStyles` (titleStyle/bodyStyle/otherStyle) по
  уровню — ПОД собственным `a:rPr` рана (прямое форматирование выигрывает). Ридер
  идёт по rel'ам слайд → layout → master, каскад мемоизирован по пути layout'а. Общие
  читалки `p:sp` (`p:ph`, `a:xfrm`-бокс, `a:rPr`/`a:defRPr` → RunProperties) вынесены
  в `sp-helpers.ts` (без цикла); defRPr переиспользует rPr-читалку. Заголовок/тело
  встают на места и нужного размера.
- **PX3 ✓ — картинки + геометрия shape** (PX3a `84b4425`, PX3b `a8696b6`).
  - *PX3a — картинки*: `p:pic` → floating `ImageBlock`; `a:blip@r:embed` резолвится по
    rel'ам слайда в media-парт, байты в контент-адресуемый `ResourceStore` (дедуп) — тем
    же путём, что docx-картинки. Парсер слайда теперь над `SlideContext{cascade,
    resolveImage}`; ридер строит per-slide резолвер картинок. Alt из `p:cNvPr@descr`.
  - *PX3b — видимые фигуры*: `p:spPr` → `a:prstGeom`/`a:custGeom` (геометрия),
    `a:solidFill`/`a:gradFill`/`a:noFill` (заливка), `a:ln` (линия) — переиспользуя
    DrawingML-читалки из `word/drawing-parser` (экспортированы, docx байт-в-ноль). Фигура
    рендерится при видимой заливке/линии ИЛИ тексте; полностью невидимая — отбрасывается.
    Scheme-цвета пока по дефолтной Office-палитре (`defaultColorResolver`), реальная тема
    деки — PX5.
- **PX4 — таблицы + чарты.** `p:graphicFrame` → `a:tbl` в FlowDoc-Table ИЛИ `c:chart`
  (`chart-parser`) в floating `ChartBlock`, позиция из `a:xfrm` рамки.
- **PX5 — тема + фоны + группы.** Цвето/шрифт-схема темы (`theme-parser`+ColorResolver);
  фон слайда/layout/master (`p:bg` solid/gradient/picture) как подложка-shape; `p:grpSp`
  группы (композиция `a:xfrm`/`a:chOff`/`a:chExt`).
- **PX6 — глубина текста + гиперссылки → релиз.** Маркеры (`a:buChar`/`a:buAutoNum`),
  уровни списка, выравнивание (`a:pPr@algn`), вертикальный якорь (`a:bodyPr@anchor`),
  автоподгонка (`a:normAutofit`/`spAutoFit`); гиперссылки (`a:hlinkClick`). Доки на
  сайте + CHANGELOG + эта секция; релиз.

### Риски
- **Холст через float-якоря** — выразительно достаточно (relativeFrom:'page' = абсолют),
  но layout пагинирует in-flow поток. Якорь-параграф (U+200B) на слайд решает рождение
  страницы; floating-shape'ы садятся на текущую страницу. Проверено на PX0.
- **Каскад layout/master** (PX2) — самый «pptx-специфичный» кусок без аналога в docx;
  делается резолвом rel-цепочки слайд→layout→master. Изолирован в PX2.
- **Перенос текста vs autofit** (PX6): pptx часто полагается на autofit-усадку; без неё
  длинный текст переполнит бокс. Консервативно: сначала честный перенос по ширине бокса,
  autofit-усадка отдельным шагом.
- Байт-в-ноль: pptx — отдельный sniff, общие writer'ы не тронуты; снапшоты docx/xlsx/pdf
  не двигаются.

### Прогресс
- **PX0 ✓** (`0726c59`) — шов: pptx → FlowDoc → PDF/SVG/HTML, одна страница на слайд
  нужного размера. 774 теста.
- **PX1 ✓** (`d76f634`) — текст слайда: `p:sp` с собственным `a:xfrm` → floating text-box
  на EMU-позиции; прямой `a:rPr`. Реальный контент на странице, на своих координатах.
  778 тестов.
- **PX2 ✓** (`f737119`) — каскад плейсхолдеров: shape без своего `a:xfrm` берёт геометрию
  из layout/master и размер/цвет из мастер-`txStyles`. Реальные деки (где title/body —
  плейсхолдеры) встают на места. 782 теста.
- **PX3 ✓** (PX3a `84b4425`, PX3b `a8696b6`) — картинки (`p:pic` → `ImageBlock` через
  ResourceStore) + видимая геометрия/заливка/линия/градиент фигур (`p:spPr` через
  переиспользованные DrawingML-читалки). На слайдах появляются картинки и цветные фигуры.
  788 тестов. Дальше: PX4 (таблицы + чарты через `p:graphicFrame`).

---

## Сводка приоритетов

| Эпик    | Усилие        | byteRisk | Связь с core-миссией       | Когда                    |
|---------|---------------|----------|----------------------------|--------------------------|
| E-DOCX  | среднее       | низкий   | docx-выход, roundtrip-гейт | **первым (сейчас)**      |
| E-SHEET | большое       | высокий  | Excel первоклассен → ядро  | вторым (стратегический)  |
| E-PDF   | очень большое | н/д      | универсальность, не ядро   | отдельный крупный заход  |
| E-PPTX  | среднее       | низкий   | pptx-вход, замыкает OOXML  | новая эра (после 1.8.0)   |
