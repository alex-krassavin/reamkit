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

### Волна V — последние «not yet» Excel-конструкты (1.13.0)
- **SV1 ✓ — data validation.** `<dataValidations>` (main-ns) → `ParsedWorksheet.dataValidations`
  (тип/sqref/формулы/промпты round-trip'ятся). `list`-валидация → флаг `CellProperties.dropdown`; слой
  эмитит ▾-кнопку у правого края ячейки в PDF (gated shape-pass, переиспользует CF-icon-машину) и HTML
  (floated inline-SVG). Write-back в `dataValidationsXml` (между `<conditionalFormatting>` и
  `<printOptions>` по §18.3.1.99), `normGrid.dataValidations` в roundtrip-гейте. `showDropDown`
  ИНВЕРТИРОВАН ("1" ПРЯЧЕТ дропдаун); x14 cross-sheet списки — задокументированный пропуск.
  `tests/data-validation.test.ts` (10). Байт-в-ноль (фичечит только файлы с валидацией).
- **SV2 ✓ — slicers.** Резолвятся ЦЕЛИКОМ в reader в `Sheet.slicers` (как чарты): персистентная
  SpreadsheetML-модель, writer и roundtrip-гейт НЕ тронуты. native-table слайсер → кнопки из distinct-
  значений колонки таблицы (`<x14:tableSlicerCache tableId column>`), выбор — из autofilter колонки;
  OLAP/pivot слайсер (items в pivot-cache) → caption-only box. Проекция эмитит каждый как стилизованную
  мини-таблицу после грида (`slicerTable` в print-model) — шапка-caption + кнопки, accent-заливка + белый
  текст на выбранных, светлая полоса на остальных. Стиль — accent-эвристика таблиц/пивотов. Парты НЕ
  пишутся обратно (как пивоты). `tests/slicer.test.ts` (8). Байт-в-ноль.

### Волна W — закрытие остатков Excel-рендера (после 1.13.0)
Честный «not yet» из scope.md, по убыванию заметности. Каждый шаг — байт-в-ноль (фичечит только файлы с
этой конструкцией), тест + ✓ здесь по факту.
- **W1 ✓ — плавающие картинки на листе.** `sheet-drawing.ts` тянет и `xdr:pic` (`blipFill/blip @r:embed`),
  не только чарты; возвращает `{charts, pictures}`. Reader резолвит media-байты → `ResourceStore` (хойстнут
  до цикла) → IR `Sheet.images{resourceId,wPt,hPt}`. Проекция эмитит `ImageBlock` после грида (anchor-
  ordered; placement collapses to inline — как чарты). Рендер бесплатно: `collectImageResources`/
  `prepareImage` → PDF image-XObject, HTML `data:`-URI. `tests/sheet-image.test.ts` (5). Байт-в-ноль (лист
  без рисунка → нет `images`).
- **W2 ✓ — плавающие фигуры/текстбоксы на листе.** `src/excel/sheet-shape-parser.ts` парсит `xdr:sp` через
  `parseXml`→PoNode (DrawingML-ридеры работают на preserveOrder-дереве, не на removeNSPrefix charts/pics —
  поэтому drawing парсится повторно, под гейтом `:sp>`/`:sp `, чтобы chart/pic-only листы не платили).
  Переиспользует `parseGeometry`/`parseTxBody` (экспортнул из slide-parser), `parseFill`/`parseLine`
  (word/drawing-parser), `parseXfrmBox`; cascade=undefined (как SmartArt) → раны по прямому `a:rPr`. Размер из
  ЯКОРЯ (from/to), не из `a:xfrm`. IR `Sheet.shapes: ShapeBlock[]`; проекция эмитит `{kind:'shape'}` после
  грида. Рендер бесплатно (PDF vector + HTML inline-SVG). `tests/sheet-shape.test.ts` (6). Байт-в-ноль.
- **W3 ✓ — ссылки на ячейках.** worksheet `<hyperlinks>` → `ParsedWorksheet.hyperlinks` (raw); reader
  резолвит `r:id`→external URL через ws-rels → IR `Sheet.hyperlinks{ref,url}`. Проекция (через
  `PrintModelOptions`) ставит `run.href` на каждую покрытую ячейку → готовый href-path рисует PDF /Link +
  HTML `<a>`. location-only (внутрикнижные) ссылки без URL — пропуск. Render-only (НЕ в `normGrid`, НЕ
  пишется обратно — как слайсеры). `tests/sheet-hyperlink.test.ts` (7). Байт-в-ноль.
- **W4 ✓ — колонтитулы листа.** `<headerFooter><oddHeader/oddFooter>` → `ParsedWorksheet.headerFooter`;
  `header-footer.ts` парсит &-мини-язык (`&L/&C/&R` регионы, `&P`→PAGE/`&N`→NUMPAGES динамические field-раны,
  `&A`→имя листа, `&B/&I` toggles, `&&`→`&`; `&D/&T/&F/&Z/&"font"/&size/&K` — дропаются). HF-банды рисуют
  ТОЛЬКО параграфы (не таблицы — проверено), поэтому каждый регион = свой выровненный параграф (одиночный
  регион — одна строка; L+R — стопка). Проекция кладёт контент в `FlowDoc.headersFooters` + section.headers/
  footers; динамическое PAGE/NUMPAGES резолвится per-page готовой машиной (docx). Render-only. `tests/sheet-
  header-footer.test.ts` (10, вкл. "Page 1 of 2 / Page 2 of 2"). Байт-в-ноль.
- **W5 ✓ — доп. типы условного форматирования.** `parseCfRule` (worksheet-parser) + `cfRuleXml` (xlsx-
  writer) + эвалуатор (`conditional-format.ts`) расширены на value/text-driven семейства: `top10`
  (top/bottom N|N%, ties включаются), `aboveAverage` (среднее ± N·σ population, `equalAverage`),
  `duplicate/uniqueValues` (частота по диапазону — числа по значению, текст без учёта регистра; ключи
  `n:`/`s:` неймспейснуты, число 5 ≠ строка "5"), текстовые (`containsText`/`notContainsText`/`beginsWith`/
  `endsWith`, регистронезависимо). Все резолвятся в готовый dxf→fill/font (как cellIs); range-семейства
  считаются один раз в `buildConditionalFormatter`. Форматтер получил `text`-параметр (полный, до
  textBudget-обрезки) + `resolveText`-резолвер (для dup/unique по строкам через sharedStrings). `expression`
  (нет формульного движка) и `timePeriod` (зависит от часов — Ream детерминирован) — осознанный graceful
  loss, скипаются. Round-trip-фикспоинт (новые правила пишутся обратно и стабильны). `tests/conditional-
  format-extended.test.ts` (18) + roundtrip-фикстура. Байт-в-ноль (гейт `value===undefined` ослаблен только
  для ячеек с текстом; числовые листы без новых правил → идентичный вывод).
- **W6 ✓ — детали формата ячеек.** Шесть под-фич отрисованы, остальные две — read+round-trip.
  - **in-cell rich text** (render-only): `parseSharedStrings`→`{texts, runs}`; rich-`<si>` (несколько `<r>`
    с собственным `<rPr>`: b/i/u/color/sz/vertAlign) → `SheetRichRun[]` параллельно плоским sharedStrings.
    `SheetDoc.sharedStringRuns` (только когда есть rich) → print-model эмитит по рану на `<r>` (`richRunProps`
    слоит `<rPr>` поверх шрифта ячейки). Writer пишет плоский текст → round-trip/byte-zero не трогаются.
  - **wrapText**: потребляет существующий round-trip-флаг `xf.alignment.wrapText` — пропускает overflow-клип,
    ячейка переносит текст по своей ширине, строка растёт (atLeast).
  - **не-solid заливки** → представительный solid: `PATTERN_DENSITY` (gray125/lightGray/darkGray/hatch…)
    блендит fg над bg по плотности в `shadingFromXf` (без новых полей). **gradientFill** → `averageGradientColor`
    (среднее стопов) в styles-parser кладётся как solid fg → рендерится фоном, round-trip как solid.
  - **indent**: `XlsxCellAlignment.indent` (model+parser+writer round-trip) → `paragraph.indentLeft` =
    `indent×3×TWIPS_PER_EXCEL_CHAR`.
  - **диагональные границы**: `XlsxBorder.diagonal/diagonalUp/diagonalDown` (parser+writer) →
    `CellBorders.diagonalUp/Down` → styled-layout рисует линию-shape через PathBuilder/flipTransform (как
    иконки/dropdown).
  - **textRotation** (рендерится): `XlsxCellAlignment.textRotation` (parser+writer round-trip) → повёрнутая/
    вертикальная ячейка рисует текст **стопкой** (по глифу на центрированный параграф, строка растёт) — честная
    flowed-раскладка для доминирующего кейса вертикальных заголовков (255 точно, ±90° по ориентации чтения),
    работает и в PDF, и в HTML без новых render-примитивов.
  - **shrinkToFit** (рендерится): scale шрифта в print-model по символьной ёмкости столбца (тот же char-модель,
    что и клип; грид авто-ресайзит столбцы по контенту, поэтому усадка по xlsx-ширине, а не по layout-ширине).
    shrinkToFit-ячейка освобождена от клипа. `XlsxCellAlignment.shrinkToFit` round-trip'ится.
  - Гейты: `tests/sheet-rich-text.test.ts` (6) + `tests/sheet-cell-format.test.ts` (10, вкл. round-trip-фикспоинт
    indent/rotation/shrink/diagonal + рендер стопки + усадку шрифта). 945 тестов. Байт-в-ноль (фичечит только
    ячейки с этими конструкциями). Все восемь под-фич РИСУЮТСЯ — недоделок нет.
- **W7 ✓ — комментарии/заметки к ячейкам.** Новый `src/excel/comments-parser.ts`: `parseLegacyComments`
  (`xl/comments`: `<authors>`+`<commentList>`, authorId→author, rich `<text>` флэттенится, leading «Author:»
  снимается), `parseThreadedComments` (`xl/threadedComments`: `<threadedComment ref personId>`) + `parsePersons`
  (`xl/persons` id→displayName, резолвится один раз из workbook-rels). Reader перечисляет ws-rels: `isOoxmlRel
  (…, 'comments')` для legacy + `.endsWith('/threadedComment')` для threaded (MS-2017 ns) → `Sheet.comments`.
  VML-бокс игнорируется (только текст+автор). Проекция эмитит секцию «Comments» после грида (заголовок + строка
  на коммент `<ref> — <author>: <text>`, многострочный текст схлопывается) — Excel-режим «print comments at end
  of sheet». Render-only (НЕ пишется обратно). Фикстура: `comments`/`threadedComments`/`persons` опции +
  `mergeWorksheetRel`-хелпер. `tests/sheet-comments.test.ts` (6: парсеры + end-to-end legacy/threaded + байт-в-
  ноль + PDF). 951 тест. Байт-в-ноль (лист без комментов → нет секции).
- **W8 ✓ — form-контролы.** worksheet-parser `parseFormControls` тянет `<control name r:id>` из
  transitional `<controls>` и из x14 `extLst` (через `mc:AlternateContent`/`Choice`; removeNSPrefix
  убирает x14:/mc:/r:; дедуп по relId) → `ParsedWorksheet.formControls` (render-only, как hyperlinks — НЕ в
  normGrid). Новый `form-control-parser.ts`: `parseFormControlProps` (ctrlProp `<formControlPr objectType
  checked val>`). Reader резолвит relId→ctrlProp→`SheetFormControl{name,objectType,checked,value}`. Проекция
  эмитит секцию «Form controls» после грида + комментов: `[x]`/`[ ]` чекбокс, `(o)`/`( )` радио, `[ name ]`
  кнопка, `name (value N)` spin/scroll, `name (list)` drop/list — ASCII-аффордансы (рендерятся в любом шрифте).
  **ActiveX** — OLE-бинарь, осознанный graceful loss (в scope.md). Фикстура: `formControls` опция (x14 extLst +
  ctrlProp-парты + ws-rels). `tests/sheet-form-controls.test.ts` (4: ctrlProp-парсер + end-to-end аффордансы +
  байт-в-ноль + PDF). 955 тестов. Байт-в-ноль (лист без контролов → нет секции).
- **W9 ✓ — формульный движок (CF expression + timePeriod).** Новый чистый модуль `src/excel/formula/`
  (lexer → Pratt-parser → evaluator + библиотека ~45 функций + дата-хелперы). Ключевой инсайт: на ячейке хранится
  только **кэш-значение** `<v>` (формула `<f>` даже не парсится), а CF-форматтер уже видит весь грид — значит
  пересчётный движок НЕ нужен, надо лишь *вычислить выражение в контексте ячейки над известными значениями*.
  `expression` — формула на ячейку (относительные ссылки сдвигаются от левого-верха sqref, как в Excel);
  `timePeriod` — 10 дата-окон относительно `options.now` (явный вход, НЕ системные часы → детерминизм;
  `TODAY()/NOW()` читают тот же `now`). Оба типа round-trip'ятся через writer (в normGrid). Неизвестная функция /
  sheet-qualified ref / defined name → ошибка → правило просто не применяется (graceful loss, никогда не
  мисрендер). Проводка `now`: `projectSheetDoc(sheet,{now})` → `worksheetToBody` → `buildConditionalFormatter`;
  публично — `ConvertXlsxOptions.now`, `ReamConvertOptions.now` (ре-проекция листа), facade `ConvertOptions.now`.
  Байт-в-ноль: `hasExpr`-гейт сохраняет ранний выход для пустых ячеек, когда expression-правил нет.
  `tests/formula-eval.test.ts` (22: движок) + `tests/conditional-format-expression.test.ts` (10: e2e/roundtrip/
  PDF). 987 тестов. **«Not yet» по Excel-рендеру закрыт** (остался только ActiveX-OLE — осознанный loss).
- **XLS-1/2/3 ✓ — legacy `.xls` (BIFF8) reader.** Бинарный кейстоун — старый `.xls` рендерится во всё
  (PDF/SVG/HTML) и даже переписывается в `.xlsx`.
  - **XLS-1** — `src/core/ole/cfb.ts`: MS-CFB (OLE2) контейнер-ридер (header → DIFAT → FAT → directory →
    miniFAT/mini-stream; `readStream(name)` выбирает FAT vs mini по cutoff 4096). Хардненинг как у
    OpcPackage: bounds-чек каждого сектора, cap цепочек по числу секторов (циклы/overrun → abort), лимит
    вывода. Шарится `.xls`/`.doc`/`.ppt`/ActiveX. Фикстура-билдер `build-cfb` + `tests/cfb.test.ts` (6:
    оба пути хранения + case-insensitive + sniff).
  - **XLS-2** — `src/excel/xls/biff-reader.ts`: запись-стрим `Workbook`. Globals (BOF/BoundSheet8/SST+
    CONTINUE/DATEMODE) + per-sheet (DIMENSIONS/NUMBER/RK/MULRK/LABELSST/LABEL/BOOLERR/FORMULA+STRING/
    MERGECELLS/COLINFO) → WorksheetCell → SheetDoc (как OOXML). Главная заноза — SST: строка может рваться
    через CONTINUE-границу, и первый байт продолжения — свежий fHighByte (continuation-aware `SstReader`).
    RK-декод (int / ×100 / IEEE). BIFF8-only (vers 0x0600); старее → понятная ошибка. Стайлинг (XF) и
    drawings — отложены (default-формат, документированный loss). `build-xls` (CFB+BIFF + патч lbPlyPos) +
    `tests/xls-reader.test.ts` (9: e2e + record-семейства + SST-split + decodeRk).
  - **XLS-3** — `xls-reader.ts` (DocumentReader, sniff = CFB + `Workbook`/`Book` стрим; не цепляет
    `.doc`/`.ppt`/зашифрованный OOXML). В DEFAULT_READERS; facade PDF-диспатч роутит sheet-ридеры
    (`renderSheetReaderToPdf`); Ream работает дженериком. Лоссы: cellFormatting (degraded) + charts
    (dropped). `tests/xls-integration.test.ts` (6: Ream `.sheet`/losses/pdf/html/`.xlsx`-rewrite + facade
    detect + не-misdetect). 1008 тестов.
  - **XLS-4** — `src/excel/xls/biff-styles.ts`: стили из FONT (0x0031, index-4-skip) / FORMAT (0x041E) /
    PALETTE (0x0092, дефолтная 56-цветная таблица + override) / XF (0x00E0, 20 байт: ifnt/ifmt/align/
    rotation/indent + упакованные стили+цвета рамок + fls+цвета заливки) → тот же `XlsxStyles` {fonts, fills
    (dedup, [0]none/[1]gray125), borders (dedup), cellXfs, numFmts}, что и OOXML. Цвет = palette-индекс →
    `FFRRGGBB`; fls→patternType, dg→border-style-имена (совпадают с print-model). `cell.styleIndex = ixfe`.
    Лосс cellFormatting снят — стили читаются (остался только charts/drawings). `tests/xls-styling.test.ts`
    (7: стиль-таблица + cell→XF + e2e шрифт/заливка/рамка/числоформат через проекцию). 1015 тестов.
  **Legacy `.xls`: значения + структура + стили** (`.doc`/`.ppt` — следующий эпик на том же CFB-кейстоуне).
- **XLS-5 ✓ — картинки в `.xls` (Escher).** Новый `src/excel/xls/escher.ts`: ридер Office-Drawing записей
  (`[verInstance][type][len]`, контейнер при ver=0xF). Globals MSODrawingGroup (0x00EB) → BStoreContainer →
  BSE-блипы (magic-scan PNG/JPEG/GIF/BMP/TIFF, индекс-выровнено); per-sheet MSODrawing (0x00EC, +CONTINUE) →
  SpContainer → OPT `pib` (0x0104) + ClientAnchor → картинки. biff-reader: общий `ResourceStore`, blip[pib] →
  `SheetImageRef{resourceId,widthPt,heightPt}` (размер из анкора по дефолт-метрикам). Рендерится через тот же
  пайплайн, что и xlsx-картинки. Фикстура: `msoDrawingGroupRec`/`msoDrawingRec` Escher-билдеры. Лосс снят с
  «images» → остались только charts + autoshapes/textboxes. `tests/xls-image.test.ts` (6: escher-парсер +
  e2e-резолв в ресурсы + проекция + байт-в-ноль + PDF-декод). 1027 тестов.
- **XLS-6 ✓ — чарты в `.xls` (BIFF chart substream).** Новый `src/excel/xls/biff-chart.ts`: вложенный
  chart-substream (BOF dt=0x20 … EOF). Тип из group-рекорда (BAR/LINE/PIE/AREA/SCATTER; barDir из BAR.grbit);
  серии — значения из range'а, на который ссылается AI-рекорд (ptgArea/ptgArea3d), резолвится по **кэш-ячейкам
  листа** (как формульный движок); категории + имена (SeriesText / name-AI). → тот же `Chart`-модель, что и
  OOXML-чарты. **Робастность:** `readSubstream` теперь depth-трекает BOF/EOF — вложенный чарт больше не обрезает
  лист (был латентный баг). biff-reader: детект чарт-substream'ов в `readSheet` → `SheetChartRef` + общий
  `chartData`-map. Фикстура `chartRecords` (BOF/тип/SERIES/AI/SeriesText). `tests/xls-chart.test.ts` (5:
  парсер + тип + не-обрезает-лист + PDF + unit). Лосс снят с charts → остались только autoshapes/textboxes.
  1032 теста.
- **XLS-7 ✓ — drawing shapes в `.xls` (автошейпы + текстбоксы).** escher.ts `parseSheetShapes`: non-picture
  SpContainer'ы → `EscherShape{shapeType (Sp.instance), hasText (ClientTextbox), anchor, fill/lineColor (OPT,
  литеральный RGB)}`. biff-reader: текст текстбокса из TXO-рекорда (0x01B6) + следующего CONTINUE (cch в TXO,
  символы в CONTINUE), ассоциация по порядку; `buildShapes` → `ShapeBlock` (preset из MSOSPT, fill/line, размер
  из анкора, text-body) → `Sheet.shapes` → тот же shape-пайплайн, что и xlsx. Фикстура `msoDrawingShapesRec`/
  `txoRecs`. `tests/xls-shape.test.ts` (4: автошейп+заливка / текстбокс+текст / проекция / байт-в-ноль / PDF).
  1036 тестов. **`.xls` визуально доделан** (значения + стили + картинки + чарты + шейпы → PDF/SVG/HTML +
  rewrite в `.xlsx`). Остался только вторичный слой (CF/комменты/гиперссылки/валидация/defined-names/print-
  setup — есть BIFF-рекорды, но не парсятся; задокументировано как degraded-loss).
- **DOC-1 ✓ — текст legacy `.doc` (Word 97–2003).** Старт следующего бинарного эпика на том же CFB-кейстоуне.
  Новый `src/word/doc/doc-text.ts`: CFB → `WordDocument`-стрим → FIB (wIdent 0xA5EC, fWhichTblStm → `0Table`/
  `1Table`, fcClx/lcbClx @0x1A2/0x1A6, ccpText @0x4C) → piece table (CLX: пропуск Prc-записей, Pcdt → PlcPcd) →
  куски (FcCompressed: bit30 → 8-бит cp1252 @fc/2, иначе 16-бит UTF-16LE @fc) → текст. Новый `doc/doc-reader.ts`
  (DocumentReader id `doc`, sniff = CFB + `WordDocument`-стрим, не цепляет `.xls`/`.ppt`/зашифрованный OOXML):
  текст → абзацы (split по CR/0x0C, control-чары дропаются, tab→пробел) → FlowDoc (Letter, поля 1") → тот же
  путь, что docx/pptx (`Ream.parse → toFlowDoc → renderStyledPdf`), работает в PDF/SVG/HTML + rewrite в `.docx`.
  Defensive (любая структурная аномалия → пустой текст, не throw); encrypted (fEncrypted) → dropped-loss;
  формат-слой (стили/таблицы/картинки/колонтитулы/списки/поля) → degraded-loss, следующая волна. Фикстура
  `build-doc.ts` (CFB + FIB + CLX; compressed/uncompressed/0Table/encrypted — пишет ровно те байты, что читает
  ридер, → round-trip валидирует offset/compression-логику). `tests/doc-reader.test.ts` (10: UTF-16 / cp1252-
  high / смешанные куски / 0Table / control-чары / encrypted / losses / sniff / format / PDF). Зарегистрирован
  в `DEFAULT_READERS` + `SOURCE_MIME`. **Бинарный кейстоун теперь несёт `.xls` и `.doc`** (остался `.ppt`).
- **DOC-2 ✓ — форматирование ранов `.doc` (CHPX).** doc-text расширен: FIB `fcPlcfBteChpx`/`lcb` @0xFA/0xFE →
  PlcBteChpx (PLC: (n+1) FC + n PnFkpChpx, page = low 22 бит) → FKP-страницы (512 байт: rgfc (crun+1 FC) + rgb
  (crun word-офсетов к CHPX, 0 = нет) + crun в байте 511; CHPX = u8 cb + grpprl) → декод sprm'ов
  (sprmCFBold 0x0835 / sprmCFItalic 0x0836 / sprmCKul 0x2A3E / sprmCHps 0x4A43; размер операнда из `spra =
  sprm>>13`, spra=6 → length-prefixed, прочие sprm'ы скипаются с сохранением выравнивания). `extractDocContent`
  возвращает `DocRun[]` (текст + DocCharProps), коалесцируя по свойствам; props ищутся по FC (бинпоиск в
  отсортированных CHPX-ранах), FC символа = `piece.fc + k*step` — стык piece-table (CP→FC) и CHPX (FC→props).
  doc-reader: `buildParagraphs` режет на абзацы по CR, переносит per-run props в `RunProperties`
  (bold/italic/underline kul→UnderlineStyle/fontSizePt = halfPts/2). Фикстура `build-doc` синтезирует FKP +
  PlcBteChpx + sprm-grpprl (single-piece FC-маппинг) → round-trip валидирует sprm/FKP/FC-логику.
  `tests/doc-reader.test.ts` +3 (bold/italic-сплит ранов / size+underline / formatted-PDF), 13 всего.
  Параграф-форматирование/таблицы/картинки/колонтитулы — следующая волна (degraded-loss).
- **DOC-3 ✓ — параграф-форматирование `.doc` (PAPX).** doc-text реструктурирован: вместо плоского
  `DocRun[]` теперь собирает `DocParagraph[]` (где живёт пер-символьный FC-трекинг — стык трёх структур по
  FC: piece-table CP→FC, CHPX FC→run-props, PAPX FC→para-props). FIB `fcPlcfBtePapx`/`lcb` @0x102/0x106 →
  PlcBtePapx (тот же PLC-шейп, что Chpx; общий дженерик `parsePlcBte`) → PAPX FKP (§2.9.23: (cpara+1) FC +
  cpara 13-байтных BX [bOffset + PHE] + cpara в байте 511; PapxInFkp = cb (≠0 → 2cb−1 байт; =0 → 2-й байт
  даёт 2cb2) → istd(2) + grpprl) → параграф-sprm'ы (sprmPJc 0x2403 / PDxaLeft 0x840F / PDxaRight 0x840E /
  PDxaLeft1 0x8411 / PDyaBefore 0xA413 / PDyaAfter 0xA414; sprm-итератор общий с CHPX). Разбивка абзацев по
  CR (0x0D), PAPX берётся по FC para-mark'а; контрол-чары чистятся в doc-text. doc-reader → ParagraphProperties
  (jc→Alignment, twips/20→Pt indent/spacing; индент знаковый — hanging). Фикстура `build-doc` обобщена:
  добавляет CHPX и PAPX FKP-страницы (512-выровненные) + PlcBte-таблицы (4-выровненные) с записью FIB-полей;
  round-trip валидирует PAPX/sprm/FC-логику. `tests/doc-reader.test.ts` +4 (alignment / indent+spacing /
  CHPX+PAPX вместе / per-paragraph alignment), 17 всего. Таблицы/картинки/колонтитулы/списки — следующая волна.
- **DOC-4 ✓ — таблицы `.doc`.** В бинарном Word таблица — не отдельная структура, а абзацы с PAPX-флагами:
  decodePapxGrpprl читает sprmPFInTable (0x2416) → `inTable` и sprmPFTtp (0x2417) → `rowEnd` (TTP, конец строки);
  buildParagraphs дополнительно режет абзацы по cell-mark'у 0x07 (как по CR) и помечает `DocParagraph.cellMark`.
  Фикс: финальный endParagraph теперь условный (только при незакрытом контенте) — иначе хвостовой маркер плодил
  фантомный пустой абзац (и фантомную строку в таблице, которая обязана кончаться на TTP). doc-reader: новый
  `groupTables` — подряд идущие inTable-абзацы → `Table`; строки режутся по rowEnd, каждый абзац = ячейка (пустой
  TTP-терминатор не ячейка). Ширины ячеек в читаемом PAPX нет → колонки делят контент-ширину (468pt) поровну,
  тонкий single-бордюр делает грид видимым. Фикстура `build-doc`: DocParaFormat += inTable/rowEnd →
  sprmPFInTable/PFTtp в grpprl. `tests/doc-reader.test.ts` +3 (2×2-грид с cell-mark'ами / абзацы до и после
  таблицы / PDF), 20 всего. Картинки/колонтитулы/списки/ширины-бордюры-merge — следующие волны.
- **DOC-5 ✓ — встроенные картинки `.doc`.** Картинка в бинарном Word — спец-символ 0x01, чей CHPX несёт
  sprmCFSpec (0x0855, fSpec) + sprmCPicLocation (0x6A03, оффсет в стрим `Data`). doc-text: decodeChpxGrpprl
  читает fSpec/picOffset; buildParagraphs на 0x01 (если fSpec+picOffset+Data) зовёт `extractPicture` →
  PICF (§2.9.158: lcb, cbHeader, dxaGoal/dyaGoal @0x1C/0x1E, mx/my @0x20/0x22) → область картинки сканируется
  на raster-magic (PNG/JPEG/GIF/BMP/TIFF — OfficeArt-blip-хедер идёт перед PNG; метафайлы WMF/EMF без magic →
  скип); возвращает `DocRun.picture{bytes,widthTwips,heightTwips}`. doc-reader: `mapParagraph` теперь
  возвращает `BodyElement[]` — текст-абзац + по `ImageBlock` на картинку (ресурсы в `ResourceStore.put`,
  twips/20→Pt); работает и в ячейках таблиц (buildRow.content = mapParagraph). Фикстура: DocFormatRun +=
  picOffset → sprmCFSpec+CPicLocation; новый экспорт `buildPicf(image,w,h)` + opts.data → стрим `Data` в CFB.
  `tests/doc-reader.test.ts` +3 (картинка→image-block с размерами / нет Data→грейсфул-скип / PDF), 23 всего.
  Колонтитулы/списки/поля/ширины-ячеек — следующие волны.
- **DOC-6 ✓ — поля `.doc`.** Поле в бинарном Word — `0x13 код 0x14 результат 0x15` прямо в тексте; раньше код
  («PAGE», «NUMPAGES \* MERGEFORMAT» …) рендерился как видимый текст (латентный баг). buildParagraphs ведёт
  счётчики fieldDepth/codeDepth: 0x13 → +1/+1, 0x14 → code−1 (переход код→результат), 0x15 → field−1 + клэмп
  (поле без сепаратора). Пока codeDepth>0 — весь текст (и картинки) подавляется; маркеры дропаются. Вложенные
  поля корректны (внешний код подавляет и результат вложенного). Только doc-text. `tests/doc-reader.test.ts`
  +2 (PAGE/NUMPAGES → «Page 1 of 3» / вложенное поле → только внешний результат), 25 всего. Колонтитулы/списки/
  ширины-ячеек — следующие волны.
- **DOC-7 ✓ — ширины ячеек таблиц `.doc` + робастность sprmTDefTable.** Латентный баг: sprmTDefTable (0xD608)
  — spra=6 sprm с **2-байтным** счётчиком операнда, а итератор `sprms` читал 1-байтный для всех spra=6 → на
  реальном `.doc` он съедал бы sprmPFTtp после TAP и ломал детект строки. Фикс: `SPRM_LONG_OPERAND` ={0xD608,
  0xC60D(sprmPChgTabs)} → читаем u16 cb (включает себя), len=cb−1, op=p+2. decodePapxGrpprl на sprmTDefTable
  парсит itcMac + rgdxaCenter ((cols+1) i16) → `DocParaProps.cellEdgesTwips`. doc-reader `buildTable`: рёбра с
  TTP → `columnWidths` (diff соседних, twips/20→Pt) в grid + per-cell `properties.width`; нет рёбер → равная
  делёжка (как было). Бордюры/merge ячеек — отложены. Фикстура: DocParaFormat += cellEdgesTwips → sprmTDefTable
  ПЕРЕД флагами таблицы (тест валидирует, что ридер скипает длинный операнд и всё равно достаёт sprmPFTtp).
  `tests/doc-reader.test.ts` +1 (2-кол грид 144/72pt + строка детектится после длинного sprm), 26 всего.
  Колонтитулы/списки/бордюры-ячеек — следующие волны.
- **W10 ✓ — ActiveX-контролы (xlsx).** Worksheet `<oleObjects><oleObject progId r:id>` (mc:AlternateContent-
  aware, `parseOleObjects` зеркалит W8 `collectControls`) → `OleObjectRef{progId,relId}`. Новый `activex-parser`:
  `xl/activeX/activeX#.xml` `<ax:ocx>`/`<ax:ocxPr name value>` → видимое состояние property-bag'а
  (caption/value/groupName); `activeXType` мапит progId (Forms.CheckBox.1→checkbox, CommandButton→button, …).
  Reader резолвит relId→activeX-парт→`SheetActiveXControl{type,caption,value,groupName}`; проекция — секция
  «ActiveX controls» после грида с теми же аффордансами, что и form-контролы (`[x]`/`(o)`/`[ Caption ]`).
  Фикстура: `oleObjects` опция (`<oleObjects>` + activeX-парты + ws-rels). `tests/sheet-activex.test.ts` (6:
  парсер + type-мап + e2e-аффордансы + SheetDoc + байт-в-ноль + PDF). 1021 тест. Байт-в-ноль (лист без
  oleObjects → нет секции). **Граница:** контрол, персистнутый только в `.bin` (MS-OFORMS, не property-bag),
  рендерится как тип без caption — чтение бинарного bag'а из OLE/CFB-стрима — следующий шаг (кейстоун готов).

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
- **PX4 ✓ — таблицы + чарты** (PX4a `fc61f89`, PX4b `72b1057`). `p:graphicFrame`
  (трансформа в `p:xfrm`, не `a:xfrm` — вынесен `boxFromXfrm`):
  - *PX4a — чарты*: `c:chart@r:id` → resolved chart-парт через `parseChart`+
    `withChartColorStyle` (как docx) → floating `ChartBlock`; распарсенный `Chart` в
    `doc.charts` под глобально-уникальным ключом (`slidePath!relId`, т.к. relId scoped к
    части), `flowRenderOptions` уже пробрасывает карту в рендер.
  - *PX4b — таблицы*: `a:tbl` → FlowDoc-`Table` (grid из `a:gridCol@w`, `a:tr`/`a:tc`,
    текст ячеек через общий `txBodyParagraphs`, заливка `a:tcPr/a:solidFill`). gridSpan →
    colSpan (hMerge-продолжение дропается), vMerge → merge start/middle. У `Table` нет
    `float` → таблица in-flow; поля слайд-секции обнулены → к верх-лево (точная позиция
    рамки — поздний хвост).
- **PX5 — тема + фоны + группы.**
  - *PX5a ✓ — тема деки* (`800be96`). Мастер-`theme` (`a:clrScheme`) → палитра над
    Office-дефолтом → `ColorResolver` (как docx/xlsx), строится рядом с каскадом и
    мемоизируется по layout-пути. Резолвер протянут во ВСЕ цвето-точки: fill/line фигур,
    цвет ранов и мастер-`txStyles` (`rPrToRunProps` берёт резолвер; попутно починен
    латентный баг — run colorHex хранился с лишним `#`), заливка ячеек, цвета чартов.
    Фолбэк на Office-палитру, если темы нет.
  - *PX5b ✓ — фоны* (`283a7db`): `p:bg` слайда (иначе layout/master) → полнослайдовый
    rect за контентом (`float.behind`); `p:bgPr` solid/gradient (через общий `parseFill`),
    `p:bgRef` аппроксимируется solid'ом по цвету. Унаследованный фон резолвится в
    `slideStylesFor` рядом с каскадом/темой.
  - *PX5c ✓ — группы* (`bc534c8`): `p:grpSp` — рекурсия в группы с композицией
    child→slide трансформы (off/ext + chOff/chExt); каждый shape/pic/frame мапит свой
    EMU-бокс через неё. Вложенные группы композятся.
- **PX6 ✓ — глубина текста + гиперссылки → релиз** (PX6a `0628093`, PX6b `f757b0a`,
  release `1.10.0`).
  - *PX6a*: выравнивание (`a:pPr@algn`), вертикальный якорь (`a:bodyPr@anchor`),
    гиперссылки (`a:hlinkClick@r:id` → `Run.href` через slide-rels, External).
  - *PX6b*: буллеты — `a:buChar` (литеральный) и `a:buAutoNum` (счётчик по уровню,
    arabic + суффикс; alpha/roman → arabic) как list-marker-ран; `a:buNone` гасит;
    отступ по уровню (`marL`/`indent`, иначе 0.5"/уровень).
  - *release 1.10.0*: CHANGELOG + bump + `SOURCE_MIME.pptx` + доки сайта (scope/getting-
    started/examples) + README. Отложено (graceful loss): autofit-усадка, picture-фон,
    picture-плейсхолдеры, alpha/roman нумерация, SmartArt.

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
  788 тестов.
- **PX4 ✓** (PX4a `fc61f89`, PX4b `72b1057`) — чарты (`c:chart` → floating `ChartBlock` +
  `doc.charts`) и таблицы (`a:tbl` → in-flow `Table`) через `p:graphicFrame`. Слайд несёт
  диаграммы и таблицы. 794 теста.
- **PX5 ✓** (PX5a `800be96`, PX5b `283a7db`, PX5c `bc534c8`) — тема деки (scheme-цвета во
  всех цвето-точках), фоны слайда/мастера (`p:bg` → подложка `behind`) и группы (`p:grpSp`
  → child→slide трансформа). Дека рендерит верные цвета, фоны и сгруппированный контент.
  801 тест.
- **PX6 ✓** (PX6a `0628093`, PX6b `f757b0a`) — выравнивание/якорь/гиперссылки + буллеты/
  отступы. Глубина текста слайда полная. 806 тестов. **Релиз 1.10.0** (CHANGELOG + bump +
  доки сайта). **E-PPTX завершён.**

**Итог E-PPTX.** `.pptx → FlowDoc → PDF/SVG/HTML/DOCX` — четвёртый вход, замыкающий
офисную OOXML-тройку на чтение. Route A оправдался: ноль нового IR и нового рендера —
слайд = секция, shape'ы = floating-элементы; весь конвейер layout→PDF/SVG/HTML + DrawingML
(shape/chart/theme/table) переиспользован. Новая подсистема `src/pptx/` (pptx-reader,
slide-parser, placeholder-cascade, sp-helpers, build-pptx-фикстура). PX0 шов → PX1 текст →
PX2 каскад плейсхолдеров → PX3 картинки+фигуры → PX4 таблицы+чарты → PX5 тема+фоны+группы →
PX6 глубина текста+ссылки. ~40 тестов в `tests/pptx-slide.test.ts`+`pptx-reader.test.ts`,
honest e2e (parse → flow/HTML/PDF). Байт-в-ноль для docx/xlsx/pdf на каждом шаге (pptx —
отдельный sniff, writer'ы не тронуты). Отложено (graceful loss): autofit-усадка, picture-
фон/плейсхолдеры, alpha/roman нумерация, SmartArt, точная позиция таблицы по рамке.

---

## E-PARITY — паритет с эталонным рендером (Word/LibreOffice)

**Цель.** Сократить визуальное расхождение `Ream.convert('pdf')` с эталонным рендером
до пиксель-близкого — БЕЗ доступа к чужим метрикам. Два рычага: (1) подставлять
шрифты, инженерно повторяющие метрики оригинала (Carlito=Calibri, Caladea=Cambria,
Arimo=Arial, Tinos=Times, Cousine=Courier) — тогда advance-ширины совпадают 1:1 и
чужие метрики не нужны; (2) опциональный `layoutProfile` (`'word'`|`'libreoffice'`),
переключающий модель высоты строки, алгоритм переноса и дефолт кернинга под выбранный
рендер. Дефолт поведения НЕ меняется (byte-в-ноль).

**Усилие: малое→среднее.** FP1 (подстановка) — точечная правка одной цепочки.
FP2–FP4 (профиль) трогают ядро layout, но строго под опт-ин флагом; дефолт = текущее.

### Где реально ломается parity (проверено на 1.10.0)
Глифовые ширины берутся из `hmtx` и масштабируются стандартно (`measure.ts:16`) — тут
расхождений нет. Дрейф даёт другое:

| Источник | Сейчас | Word / LibreOffice |
|---|---|---|
| **Подстановка sans** | `Arial/Calibri → Roboto` (`remote-fonts.ts:73`) — НЕ метрик-совместим | Calibri→Carlito, Arial→Arimo (ширины 1:1) |
| **Высота строки** | хардкод `fontSize*1.2` (`styled-layout.ts:2400`); `OS/2 win*`/`sTypo*` НЕ читаются | из метрик: `usWinAscent+usWinDescent` (Word) / `hhea asc+desc+gap` (LO) |
| **Перенос строк** | Knuth-Plass total-fit (`knuth-plass.ts:94`) | оба — greedy first-fit |
| **Кернинг** | GPOS всегда (`opentype-layout.ts:57`) | в Word по умолчанию ВЫКЛ (без `w:kern`) |

serif→Tinos и mono→Cousine уже метрик-совместимы (Croscore = клоны Times/Courier);
выбивается только sans→Roboto.

### Главное архитектурное решение (принять ДО кода)
- **Метрик-совместимая подстановка ВМЕСТО чужих метрик.** Кавеат «нужны точные метрики
  чужого тула» снимается так: подставляем шрифт, СПРОЕКТИРОВАННЫЙ повторять те метрики
  (Croscore/Carlito/Caladea). Advance-ширины совпадают 1:1 без эмуляции. Покрывает ~90%
  реальных документов (Calibri/Arial/Times/Courier).
- **Паритет сверх этого — опт-ин `layoutProfile`, НИКОГДА не дефолт.** `'ream'` (дефолт)
  = текущее поведение, байт-в-ноль. `'word'`/`'libreoffice'` бандлят {модель leading,
  алгоритм переноса, дефолт кернинга}. Ream остаётся КОРРЕКТНЫМ наборщиком; профиль — это
  режим ЭМУЛЯЦИИ конкретного рендера, явный выбор вызывающего.
- **Measurement-first.** Чужих метрик нет, но расхождение против реального вывода
  LibreOffice измеримо корпус-харнессом (baseline-drift + pixel-mismatch). Сначала
  формализуем метрику в гейт (FP0), потом каждый шаг ДОКАЗЫВАЕТ улучшение числом.

### Точки привязки (проверено на 1.10.0)
- `remote-fonts.ts:30,37,68` — `FamilyKey='roboto'|'tinos'|'cousine'`, `FAMILIES`,
  `resolveFamilyKey` (sans→roboto). `FetchLike`-инъекция для тестов уже есть.
- `ttf-parser.ts:105` hhea asc/desc парсятся (но не для leading); `:148` OS/2 читается
  только под capHeight/xHeight/weight/italic — `usWinAscent/usWinDescent/sTypoAscender/
  sTypoDescender/sTypoLineGap` НЕ читаются.
- `styled-layout.ts:2383` `computeLineHeight` (флэт 1.2×), `:2403` `lineDescent` (0.2×).
  Вызовы leading: `:1177,1607,1836`; `:1620` baselineY через `lineDescent`.
- `knuth-plass.ts:94` `breakLines` (total-fit), `:51` `TOLERANCE_RATIO`. Glue
  stretch/shrink — `styled-layout.ts:~2219`.
- `opentype-layout.ts:57` `parseGposKerning` (GPOS kern), `:40` GSUB liga. Кернинг
  применяется безусловно в shaping.
- `scripts/corpus/run.ts`+`lib.ts` — эталон LibreOffice (`soffice --convert-to pdf`),
  растеризация `mutool`, `structuralDiff`/`visualDiff` (baseline-drift + pixel-mismatch).
  `CORPUS_AUTOFONT=1` гоняет remote-подстановку.
- Опции конвертера: `ConvertDocxOptions extends Omit<StyledRenderOptions,…>` — новый
  `layoutProfile` дотечёт до конвертеров даром (как было с `pdfA`).

### Декомпозиция (опт-ин профиль; дефолт байт-в-ноль)
- **FP0 — parity как число (гейт).** В корпус-харнессе свести divergence к одной
  отслеживаемой метрике: медианный baseline-drift + доля документов с совпавшим числом
  страниц. Зафиксировать текущий baseline (до FP1) в репорте/`corpus/parity-baseline.json`.
  Инструмент проверки для FP1–FP4. Байт-в-ноль (только скрипты).
- **FP1 — метрик-совместимая подстановка.** `remote-fonts.ts`: двухуровневый матч в
  `resolveFamilyKey` — точное семейство → класс.
  - *FP1a — sans→Arimo* (тот же CDN: `@expo-google-fonts/arimo`): Arial/Helvetica/Segoe/
    Verdana → Arimo (1:1 с Arial). Уже улучшает Calibri (Arimo ближе Roboto). Расширить
    `FamilyKey`/`FAMILIES`.
  - *FP1b — точные модерн-Office: Carlito(=Calibri), Caladea(=Cambria).* Точный матч
    `calibri→carlito`, `cambria→caladea`. РИСК-источник: Carlito/Caladea могут
    отсутствовать в `@expo-google-fonts` — выбрать CDN (оба свободны, OFL); если нет —
    фолбэк Calibri→Arimo, Cambria→Tinos (близко, не 1:1) + TODO.
  - Байт-стабильно: snapshot-тесты используют caller-шрифты (детерминизм), не
    remote-цепочку; проверить, что ни один snapshot не запинен на Roboto.
- **FP2 — leading из метрик (профиль).** Расширить ttf-parser: `usWinAscent/usWinDescent`
  + `sTypoAscender/Descender/sTypoLineGap`. Ввести `options.layoutProfile?:
  'ream'|'word'|'libreoffice'` (дефолт `'ream'`). В `computeLineHeight`/`lineDescent` под
  профилем считать высоту из метрик (`'word'`: winAscent+winDescent; `'libreoffice'`: hhea
  asc+desc+lineGap), оставив auto/atLeast/exact/multiple-правила ECMA. `'ream'` = текущая
  ветка нетронута.
- **FP3 — greedy перенос (профиль).** Под `'word'`/`'libreoffice'` — жадный first-fit
  рядом с Knuth-Plass, чтобы число строк/страниц совпадало с эталоном (главный каскадный
  дрейф). `'ream'` остаётся total-fit.
- **FP4 — кернинг под профиль.** `'word'`: гасить GPOS-кернинг без документного `w:kern`;
  проверить чтение legacy `kern`-таблицы (шрифты без GPOS LO кернит). `'libreoffice'`:
  кернинг как есть.

### Риски
- **Carlito/Caladea — источник шрифта** (FP1b): не на Google Fonts; нужен CDN/бандл.
  Свободны (OFL), но требуют решения. Фолбэк на Arimo/Tinos обозначен.
- **layoutProfile ≠ дефолт**: любой профиль трогает ядро layout; строго под флагом,
  `'ream'` байт-в-ноль. Снапшот-гейт стережёт регресс дефолта.
- **Greedy vs total-fit** (FP3): два пути переноса в одном модуле — риск дивергенции;
  общий item-поток, развилка только в выборе точек разрыва.
- **Эмуляция ≠ корректность**: профиль воспроизводит ВЫБОРЫ рендера (в т.ч. его
  «хуже»-типографику). Документировать: `'ream'` — лучшая типографика, `'word'`/
  `'libreoffice'` — паритет.

### Прогресс
- **FP1 ✓** — метрик-совместимая подстановка. `resolveFamilyKey` стал двухуровневым:
  точные твины (Calibri→Carlito, Cambria→Caladea, Arial→Arimo, Times→Tinos,
  Courier→Cousine — ширины 1:1) → классовый фолбэк (прочий serif→Tinos, mono→Cousine,
  иначе sans→Arimo). Дефолт sans сменён Roboto→Arimo (LibreOffice-выровненный нейтрал;
  Roboto убран из набора). `fontUrl` поддержал вложенную раскладку пакета (Carlito:
  `/<variant>/Carlito_<variant>.ttf`); `docx-to-pdf` сид/база `'roboto'`→`'arimo'`. Все 5
  твинов — с того же `@expo-google-fonts` CDN; live-смоук: Carlito/Caladea/Arimo фетчатся
  и парсятся (`parseTtf` numGlyphs>0, все 4 варианта). Байт-в-ноль: снапшоты на
  caller-шрифтах не двигаются. 807 тестов (+1: nested-путь Carlito).
- **FP2 ✓** — leading из метрик под опт-ин `layoutProfile`. `ttf-parser` теперь читает
  hhea `lineGap` + OS/2 `usWin*`/`sTypo*` + бит USE_TYPO_METRICS (фолбэк на hhea, если
  нет OS/2) в `ParsedTtf.vmetrics`. `StyledRenderOptions.layoutProfile?:
  'ream'|'word'|'libreoffice'` (дефолт `'ream'`); `wrap`→`lineFromRange` под профилем
  считает высоту/спуск строки из метрик токен-шрифтов (max по строке) и кладёт на `Line`
  (`metricHeightPt`/`metricDescentPt`); `computeLineHeight`/`lineDescent` их читают,
  иначе прежний флэт 1.2×/0.2. `'word'` = `winAscent+winDescent` (GDI-бокс);
  `'libreoffice'` = hhea-триплет (или typo при USE_TYPO_METRICS). Опция дотекает через
  `ConvertDocxOptions` спредом (как `pdfA`), хопов минимум (1 вызов `wrap`, 1 —
  `lineFromRange`). Байт-в-ноль на `'ream'` (снапшоты не двигаются). 811 тестов (+4:
  дефолт==ream, формулы word/libreoffice по реальному шрифту, vmetrics).
- **FP3 ✓** — greedy-перенос под профилем. `greedyBreakLines` (новый
  `src/core/line-breaker/greedy.ts`) — жадный first-fit: строка набивается до последнего
  влезающего разрыва (glue после box / непрещённый penalty); сверхдлинное слово —
  аварийный разрыв на свою строку; переполнение перед forced-разрывом отрезается по
  последнему хорошему. `wrap` ветвится: `'ream'` → Knuth-Plass total-fit, профиль →
  greedy (тот же формат `breaks[]`, последний — forced-sentinel; потребитель не тронут).
  `'word'`/`'libreoffice'` рвут одинаково (leading на горизонтальные разрывы не влияет),
  но иначе, чем `'ream'`. `FORBIDDEN_BREAK` экспортирован из knuth-plass для общей
  фисибилити. Байт-в-ноль на `'ream'` (ветка только под профилем). 819 тестов (+8: 5 unit
  greedy вкл. детерминированное greedy≠KP, +3 интеграции — дефолт==ream, word==libreoffice,
  word≠ream на узкой колонке, текст сохранён).
- **FP4 ✓** — кернинг под профилем. Находка: эмиттер пишет `Tj` с `/W`-ширинами из
  `hmtx` → GPOS-кернинг в видимый PDF НЕ попадает (ни в одном профиле), влияет только на
  измерение (`textWidthPt` → ширины токенов → переносы). Под `'word'` строим layout-меру
  без кернинга (`createFontMeasure(parsed, kern=false)`): Word по умолчанию не кернит, и
  это совпадает с нашим же un-kerned рендером; `'ream'`/`'libreoffice'` кернинг сохраняют.
  Гейт — одна строка в `collectFontResources` (`kern = profile !== 'word'`); emit-мера
  (cid-font) кернинг и так не трогает (только gids/`/W`). Байт-в-ноль (дефолт `kern=true`).
  821 тест (+2: мера kern on/off на реальных парах Roboto; kern-rich абзац шире под
  `'word'`, чем `'libreoffice'`==`'ream'`). Хвост (measurement-only, маргинально): legacy
  `kern`-таблица для шрифтов без GPOS, пер-ран `w:kern`-фиделити, видимый кернинг через
  TJ-эмит.
- **FP0 ✓** — parity-метрика как гейт + прогон через Docker LibreOffice. Харнесс
  `scripts/corpus/run.ts` получил knob `CORPUS_PROFILE` (`ream`/`word`/`libreoffice`),
  прокинутый в `convertDocx/XlsxToPdf`. Прогон 8 фикстур `inputs/` под `CORPUS_AUTOFONT=1
  CORPUS_SANDBOX=docker` (Ream→Arimo, LO→Liberation Sans — метрически равны, честное
  сравнение): на прозе docx `'libreoffice'` режет median baseline-drift в **4–10×** —
  text-basic 2.2→0.5pt, justified 2.0→0.3, styled 1.4→0.2, list 1.2→0.1, header-footer
  0.8→0.1 (FP2-leading + FP3-greedy подтверждены против реального LO). Находка: профиль
  УХУДШАЛ xlsx (высоты строк Calc ≈ флэт-1.2×, не hhea) → `xlsx-to-pdf` теперь
  деструктурирует `layoutProfile` наружу (геометрия таблиц Excel — Calc-модель, не
  потоковый leading); sheet-* вернулись к бейзлайну 3.3/3.4pt. 822 теста (+1: xlsx
  байт-в-байт с профилем и без).

**Итог E-PARITY (код).** FP1–FP4 закрыты: метрик-совместимая подстановка (Carlito/Caladea/
Arimo, ширины 1:1) + опт-ин `layoutProfile` (`'ream'`|`'word'`|`'libreoffice'`), бандлящий
leading из метрик шрифта (FP2) + greedy-перенос (FP3) + дефолт кернинга (FP4) под выбранный
рендер. Дефолт `'ream'` байт-в-ноль на каждом шаге (всё новое — строго под не-дефолтным
профилем). Новое: `src/core/line-breaker/greedy.ts`, `ParsedTtf.vmetrics`,
`createFontMeasure(kern)`, `StyledRenderOptions.layoutProfile`. ~14 тестов
(`layout-profile`/`line-breaker-greedy`), формулы привязаны к метрикам реального шрифта.
**FP0 закрыт** — harness-knob `CORPUS_PROFILE` + прогон через Docker-LibreOffice
эмпирически подтвердил, что `'libreoffice'` режет docx-drift в 4–10× против реального LO, а
xlsx скоупится прочь. **E-PARITY завершён** (FP1–FP4 + FP0). Открытый хвост (вне эпика):
видимый кернинг через TJ-эмит, пер-ран `w:kern`, тюнинг профиля по большому корпусу.

---

## E-SMARTART — SmartArt-диаграммы через DrawingML-fallback (docx + pptx)

**Цель.** Рендерить SmartArt в docx и pptx через готовый drawing-override
(`diagrams/drawing#.xml`), переиспользуя существующую шейпо-машинерию. Закрывает пункт
«Not yet» из `scope.md` (Word comments / SmartArt / Excel pivot…). Выбран первым по
ROI: бьёт сразу в docx И pptx, переиспользует максимум.

**Усилие: малое→среднее.** Не новый рендер: `dsp:spTree` структурно = группа DrawingML-
шейпов, которую Ream уже рисует. Стоимость — только фронт: резолв рельс до drawing-
override + `dsp:`-walker, кладущий `dsp:spPr`/`dsp:txBody` в существующие `a:`-читалки.

### Главное архитектурное решение (принято ДО кода): drawing-override, не layout-движок
- **A (выбран): читаем готовый `diagrams/drawing#.xml`** (`dsp:drawing/dsp:spTree`) —
  позиционированные шейпы, которые Office УЖЕ разложил. Эмитим как floating-шейпы (docx
  `ShapeBlock` / pptx float-элемент). Ноль нового layout. Покрывает современные файлы
  (Office 2010+ пишет fallback-override).
- **B (отвергнут): исполнять SmartArt layout-алгоритм** (`data1.xml` + `layout1.xml`) —
  огромный движок раскладки узлов, фактически переписать Office. Не оправдано.
- **C (graceful): файлы без override** (старые/редкие) → loss «SmartArt без drawing
  fallback», без падения.

**Ключ по namespace (проверено):** `dsp:` — это ТОЛЬКО обёртки (`dsp:sp`/`dsp:spPr`/
`dsp:txBody`/`dsp:spTree`/`dsp:txXfrm`). Внутри `dsp:spPr` геометрия/заливка/линия/xfrm —
обычные `a:`-элементы (`a:prstGeom`/`a:solidFill`/`a:ln`/`a:xfrm`), т.к. тип `dsp:spPr` =
`a:CT_ShapeProperties`. Значит `parseFill`/`parseLine`/`parsePrstGeom`/`parseXfrmBox`/
`parseTxBody` работают над `dsp:`-узлами БЕЗ правок — нужен лишь walker, достающий обёртки.

### Точки привязки (проверено на 1.11.0)
- docx: `drawing-parser.ts:182` (`graphicData`/`uri`); ветки `WPS_URI`(185)/`CHART_URI`(190)/
  pic-fallback(205); константы uri 35-36 → добавить `DIAGRAM_URI =
  '…/drawingml/2006/diagram'`. `document-parser.ts:363` `tryExtractDrawingFromParagraph`
  диспатчит `content.kind`→`BodyElement`.
- pptx: `slide-parser.ts:195` (`parseGraphicFrame` uri); `CHART_URI`(198)/`TABLE_URI`(213)/
  fallback-undefined(222) → добавить ветку diagram. Группо-машинерия `parseSlideShapes`(85)
  / `composeGroupTransform`(110) / `parseSp`(151) переиспользуется для `dsp:spTree`.
- rels: `pkg.getPartRelationships(path)` + `pkg.resolveRelatedPart(path, rel)`
  (`opc/package.ts:121,140`); пример цепочки от не-корневой части — `docx-reader.ts:255`
  (`loadCharts`), `pptx-reader.ts:195` (per-slide image resolver).
- модель: shapes ПЛОСКИЕ — `BodyElement{kind:'shape', shape:ShapeBlock}` (`types.ts:699`),
  контейнера группы нет; диаграмма = набор floating-`ShapeBlock`'ов на боксе anchor'а.

### Декомпозиция (вертикальными срезами; новая uri-ветка → байт-в-ноль без SmartArt)
- **SA0 — шов: резолв override + `dsp:spTree`→шейпы, проводка в pptx.** Резолв
  `dgm:relIds@r:dm` → `diagrams/data#.xml` → (rel type `diagramDrawing`) → `drawing#.xml`;
  `parseDiagramDrawing` обходит `dsp:spTree/dsp:sp`, эмитит floating-шейпы (geometry/fill/
  line через существующие `a:`-читалки), позиции через frame→child трансформу (как группа).
  End-to-end: pptx со SmartArt → PDF, шейпы на координатах.
- **SA1 — текст в шейпах** (`dsp:txBody` → `ShapeTextBody` через `parseTxBody`): абзацы/
  раны/выравнивание/цвет/буллеты.
- **SA2 — docx-проводка** (`parseDrawing` `DIAGRAM_URI` → floating `ShapeBlock`'ы; резолв
  через `word/document.xml.rels`). Тот же `parseDiagramDrawing`.
- **SA3 — цвета/тема + graceful loss + демо.** scheme-цвета через ColorResolver; нет
  override → loss; демо docx+pptx.

### Риски
- **Нет drawing-override** (старые файлы) → graceful loss, не падение; layout-движок вне
  области.
- **`dsp:`-обёртки vs `a:`-контент** — подтверждено: правок в sub-парсеры не нужно, только
  walker; если где-то всплывёт `dsp:`-контент — точечно добавить префикс.
- Байт-в-ноль: новая uri-ветка не трогает существующие пути; снапшоты docx/xlsx/pdf без
  SmartArt не двигаются.

### Прогресс
- **SA0 ✓** — pptx-шов: drawing-override резолв + `dsp:spTree`→шейпы. `parseGraphicFrame`
  получил ветку `DIAGRAM_URI` (теперь возвращает `BodyElement[]` — диаграмма = много
  шейпов); `dgm:relIds@r:dm` → `makeSlideDiagramResolver` идёт slide→`diagrams/data#.xml`→
  (rel `/diagramDrawing`)→`drawing#.xml`, парсит `dsp:spTree`. `parseDspSp` обходит `dsp:sp`,
  кладёт `dsp:spPr`/`dsp:txBody` в существующие `a:`-читалки (geometry/fill/line/txBody) —
  правок в sub-парсеры НЕ потребовалось, как и предсказано. `diagramTransform` мапит
  child-space диаграммы на бокс фрейма (как группа). Нет override → пусто (graceful). Тесты:
  узлы → floating-шейпы, позиции точно 72/288pt, текст в PDF, no-override→0 шейпов.
  Байт-в-ноль (uri-ветка не трогает существующие пути). 826 тестов (+4). (SA1 «текст-
  полнота» по факту вошёл в SA0 — `parseTxBody` уже даёт буллеты/уровни/выравнивание.)
- **SA2 ✓** — docx-проводка + общий `parseDiagramDrawing`. Вынес walker из slide-parser в
  экспортируемый `parseDiagramDrawing(spTree, transform, makeFloat, colors, resolveLink)` —
  переиспользуют pptx (page-relative `floatAt`) и docx. docx: `DrawingContent` получил
  `diagram`-вариант, `parseDrawing` — ветку, `ParseContext.resolveDiagram`,
  `makeDiagramResolver` (`document.xml.rels`→`diagrams/data#.xml`→`/diagramDrawing`→
  `drawing#.xml`); узлы — column/paragraph-relative floats (inline-кейс, основной для docx).
  `tryExtract` теперь возвращает `BodyElement[]` (диаграмма = много шейпов; image/chart/shape
  байт-идентичны). `parseXml` экспортирован из pptx-reader. Цикла нет (acyclic word→pptx).
  Байт-в-ноль: новая uri-ветка не трогает существующее, no-override → абзац сохраняется.
  829 тестов (+3 docx). Хвост: anchored-docx (page-relative) + точная inline-позиция по
  потоку; вынос `parseDiagramDrawing` в нейтральный `core/drawingml` (сейчас acyclic
  word→pptx).
- **SA3 ✓** — цвета/тема + graceful loss + демо; эпик закрыт. (1) **Явный loss «нет
  override».** Новый `Feature` `shapes.smartArt`; общая фабрика `noDiagramOverrideLoss()`
  (slide-parser) эмитит `dropped`-лосс. Проводка через новый `SlideContext.onLoss` /
  `ParseContext.onLoss`: pptx обогащает `where: 'slide N'`; docx-reader теперь реально
  возвращает `losses` (был жёсткий `[]`) — sink висит на body-контексте (HF/notes
  диаграммы не резолвят → без ложных лоссов). `shapes.smartArt` НЕ в `supports`: полный
  SmartArt = layout-движок (вне области) — рендерим при наличии override, иначе лосс, и это
  честная декларация возможностей. (2) **Scheme-цвета** уже резолвятся общим `ColorResolver`
  (тот же путь, что обычные шейпы); закреплено тестами: `<a:schemeClr val="accent1"/>` в
  override → тема `accent1`=FF8800 в pptx И docx (docx через `word/theme/theme1.xml`). (3)
  **Демо:** docx SmartArt end-to-end → PDF (текст узлов в выводе) + HTML — впервые проверен
  сквозной рендер docx-диаграммы (SA2 проверял только FlowDoc-модель); pptx-демо был в SA0.
  Байт-в-ноль: лоссы — метадата, не байты; фикстуры без override-less SmartArt дают
  `losses: []` как раньше. 832 теста (+3). Хвосты эпика — те же из SA2 (anchored-docx
  page-relative; точная inline-позиция по потоку; вынос `parseDiagramDrawing` в нейтральный
  `core/drawingml`).

---

## E-COMMENTS — рецензентские комментарии Word (docx → PDF/HTML)

**Цель.** Рендерить ревью-комментарии docx (автор/дата/текст), привязанные к точке ссылки в
теле, в PDF и HTML. Сейчас тихо отбрасываются. Следующий пункт «Not yet» из `scope.md`.

**Усилие: малое→среднее.** Не новая машинерия: комментарий структурно = сноска с автором.
Переиспользуем почти весь конвейер footnotes/endnotes (модель-карта по id, `parseNotes`,
`transformNotes`, секция-в-конце, inline-маркер).

### Главное архитектурное решение (ДО кода): FlowDoc-секция + маркер, не балунны на полях
- **A (выбран): inline-маркер у ссылки + секция «Comments» в конце** (как endnotes), плюс
  позже нативные PDF `/Text`-аннотации. Ноль нового layout: переиспользуем notes-рендер.
  Верность: весь текст/автор/дата сохранены, привязка — маркером.
- **B (отвергнут как первый срез): балуны на правом поле с выносными линиями** — требует
  рефактора layout (резерв колонки поля + leader-lines). Вне первого среза.
- **C (graceful): docx без `comments.xml`** → новое поле отсутствует → байт-в-ноль.

### Точки привязки (проверено на 3aa7e1a)
- модель: `flow.ts:41-42` (`footnotes`/`endnotes` карты) → добавить `comments` рядом;
  `document-model/types.ts:247-248` (Run `footnoteRef`/`endnoteRef`) → добавить `commentRef`.
- парсер: `document-parser.ts:695-703` (`w:footnoteReference`/`w:endnoteReference` → ref на
  ране) → ветка `w:commentReference`; заодно ловить `w:commentRangeStart`/`End` @w:id (для
  подсветки диапазона в CM2).
- ридер: `docx-reader.ts:52-53` (`FOOTNOTES_PART`/`ENDNOTES_PART`) → `COMMENTS_PART =
  'word/comments.xml'`; `:96-108` (`noteCtx` + `parseNotes`) → загрузка через comment-aware
  парс (карта id→{author,date,initials,content}); `:155-160` (`transformNotes`-проводка в
  FlowDoc) → добавить `comments`; `:275` `transformNotes`. Лоссы уже протянуты (SA3).
- фича: `features.ts` → `comments: 'comments'`.
- рендер: HTML `html-writer.ts` (notes-секция + ветка `run.footnoteRef` ~482-489) →
  ветка `commentRef` + секция; PDF `styled-page-emitter.ts:192-229` (link-аннотации) → задел
  для нативных `/Text`-аннотаций (CM2).

### Декомпозиция (вертикальными срезами; нет comments.xml → байт-в-ноль)
- **CM0 — модель+парс+маркер+loss.** `Run.commentRef`; парс `w:commentReference`; модель
  `Comment{author,date,initials,content:BodyElement[]}` + карта `comments` на FlowDoc;
  загрузка `comments.xml` (comment-aware парс — `parseNotes` теряет атрибуты, нужен свой).
  End-to-end: комментарии в модели; без комментариев → байт-в-ноль.
- **CM1 — рендер: inline-маркер + секция «Comments» (HTML+PDF через FlowDoc).** Маркер у
  ссылки; список «author (date): text» в конце. Оба таргета сразу, максимум переиспользования.
- **CM2 — нативные PDF `/Text`-аннотации + подсветка диапазона** (`styled-page-emitter`):
  sticky-note у ссылки; подсветка `commentRangeStart..End`. Верность-срез, отложен.
- **CM3 — полиш:** `commentsExtended.xml` (ответы/треды), `people.xml` (резолв автора),
  docx write-back (roundtrip), демо, scope.md.

### Риски
- `parseNotes` теряет атрибуты элемента (author/date) → нужен comment-aware парс.
- Range-маркеры (start/end) бывают несбалансированы/вложены → толерантно (id-карта, не стек).
- Секция-в-конце меняет пагинацию — ок: новый контент только для docx С комментариями.

### Прогресс
- **CM0 ✓** — модель+парс+загрузка. `Run.commentRef` (рядом с footnoteRef/endnoteRef);
  парс `w:commentReference` в `collectRuns`; модель `Comment{content,author,initials,date}`
  (в `document-model/types.ts`, экспорт через barrel) + карта `comments` на FlowDoc; новый
  `parseComments` (comment-aware — `parseNotes` теряет атрибуты) рядом с `parseNotes`; ридер:
  `COMMENTS_PART`, загрузка через `noteCtx`, `transformComments` (те же FlowDoc-трансформы,
  что у нот, метадата проходит насквозь). Тесты: автор/инициалы/дата/текст в `flow.comments`;
  ран помечен `commentRef`; без `comments.xml` → `comments` отсутствует (ран сохраняет
  dangling-ref для рендера CM1). Байт-в-ноль: новые опц-поля, docx без комментариев не
  двигаются. 835 тестов (+3). `commentRangeStart/End` пока игнорируются (для подсветки в CM2).
- **CM1 ✓** — рендер в обоих таргетах (HTML+PDF через FlowDoc, как endnotes — отдельное
  поле, НЕ дописываем в body → docx-writer roundtrip не задет). HTML: `collectNoteNumbers`
  даёт `comments`-карту (нумерация по порядку ссылок); ран `commentRef` → `<sup>[n]</sup>`
  со ссылкой; `emitCommentsSection` — `<section class="comments">` с автором/датой/текстом.
  PDF/layout: `assignNoteNumbers` нумерует комментарии + переписывает ран в маркер `[n]`
  (superscript); хвост после endnotes (`commentTailBlocks` — `[n] author, date: content`);
  `StyledRenderOptions.comments` + `flowRenderOptions` + сбор шрифтов/ресурсов по контенту.
  Байт-в-ноль: guard `hasRefs`/size===0 не трогает доки без комментариев; ни один PDF-снапшот
  не сдвинулся (в фикстурах комментов нет). 837 тестов (+2: HTML + PDF).
- **CM2 ✓ — кликабельный маркер.** Маркер `[n]` в PDF стал кликабельным переходом к записи
  комментария (internal GoTo на закладку `comment-${n}`), ПЕРЕИСПОЛЬЗУЯ проверенный link+
  bookmark путь (W1/W5a) — PDF/A-безопасно и tagged-безопасно, паритет с HTML (там маркер уже
  ссылка на `#cm-n`). Реализация: `commentMarkerRun` (в `assignNoteNumbers`) добавляет
  `anchor: comment-${n}` ТОЛЬКО для рендеримых комментов (проброшен `commentIds` из
  `options.comments` → dangling-ref без `comments.xml` получает маркер без ссылки, без битого
  dest); `commentTailBlocks` кладёт `bookmarks: [comment-${n}]` на первый абзац записи →
  ассемблер пагинации регистрирует dest. Байт-в-ноль: доки без комментов идентичны (anchor
  только на маркер-ране); ни один снапшот не сдвинулся; veraPDF зелёный. Тест: PDF содержит
  `/Subtype /Link` + `/S /GoTo`. 845 тестов (+1). **Сознательно отложено** (диспропорция
  риск/польза): нативные `/Text` sticky-note-аннотации (CM2b) и подсветка диапазона
  `commentRangeStart/End` (CM2c) — обе тянут PDF/A + tagged-PDF annotation-compliance (риск
  veraPDF-гейту) ради фиделити поверх уже-читаемого рендера.
- **CM3 (частично) ✓ — docx write-back.** Комментарии теперь переживают docx→docx (был
  пробел). Зеркало WT2 (footnotes write-back): `emitComments` пишет `word/comments.xml`
  (`<w:comment w:id w:author w:date w:initials>` + контент через общий `emitBlock`, без
  separator-стабов) + content-type + doc-rel; run-write-back эмитит `<w:commentReference>`
  для ранов с `commentRef` (и они переживают `visible`-фильтр пустых ранов). Тест: docx с
  комментом → `convert('docx')` → ре-рид → автор/инициалы/дата/текст и `commentRef` на ране
  сохранены. Байт-в-ноль: `emitComments` только при `flow.comments`; ветка commentRef только
  для commentRef-ранов → docx без комментов идентичны, корпус-roundtrip-гейт (D6) зелёный.
  846 тестов (+1). **E-COMMENTS: ядро + roundtrip готовы (CM0–CM3).** Мелкие хвосты:
  `commentsExtended.xml` (треды/ответы), `people.xml` (резолв автора), демо-артефакт; +
  отложенные CM2b/CM2c (нативные `/Text`/подсветка).

- **CM4 ✓ — треды + resolved через `commentsExtended.xml` (Microsoft w15).** Reader: ключ
  треда — `w14:paraId` последнего абзаца комментария; `w15:paraIdParent` связывает ответ с
  родителем, `w15:done` — закрытый тред. `parseCommentThreads` = `parseCommentsRaw`
  (контент + paraIds) + `parseCommentsExtended` (paraId → {paraIdParent, done}) + линкер →
  на `FlowDoc.comments` едут `parentId`/`done`. Префикс-агностичные `poAttrLocal`/`poIsLocal`
  в `po-helpers` (poAttr знает только w:/r:/m:/xml:, не w14/w15). Рендер: HTML вкладывает
  ответы в `<ol class="comment-replies">` под родителем и метит resolved-тред; PDF-хвост
  отступает ответы по глубине треда (`indentLeft = 18pt·depth`) и добавляет ASCII-подсказки
  `(in reply to [n])` + `(resolved)` (без glyph-риска во встроенном шрифте). Writer:
  `emitComments` штампует детерминированные `w14:paraId` (FNV-1a от id) на последний абзац +
  объявляет `xmlns:w14`, возвращает карту paraId; `emitCommentsExtended` пишет
  `word/commentsExtended.xml` + content-type + doc-rel — только когда есть инфо о треде. Тред
  переживает docx→docx (paraId синтезируются заново, но `parentId`/`done` восстанавливаются).
  Тесты: модель (parentId/done), HTML-вложенность + resolved, PDF (`(resolved)`/`in reply
  to`), roundtrip. Байт-в-ноль: новый парт и `w14`-namespace появляются только в docx с
  комментами → не-комментные документы идентичны; D6-roundtrip + veraPDF (8/8) зелёные.
  **850 тестов (+4).** **E-COMMENTS закрыт (CM0–CM4).** Остались: `people.xml` (резолв
  автора), демо-артефакт; + отложенные CM2b/CM2c.

- **Хвосты закрыты (people.xml / CM2c / CM2b / демо) — E-COMMENTS закрыт ПОЛНОСТЬЮ.**
  - **people.xml ✓** — `parsePeople` (w15:person → presenceInfo/@userId) + `applyAuthorIds`;
    `Comment.authorId` (обычно email), показывается в HTML-метаданных коммента. Префикс-агностично,
    docx без people-парта не затронуты.
  - **CM2c — подсветка диапазона ✓** — `w:commentRangeStart/End` → `Run.commentRangeRefs`
    (open-range стейт в ParseContext, диапазон может пересекать абзацы). HTML вкладывает
    span в `<span class="comment-range">` (мягко-жёлтый); PDF-эмиттер заливает тот же цвет
    под подсвеченными токенами через `TextToken.highlight` + gated pre-pass (ET→`re`/`f`→BT
    только при наличии подсветки → не-комментные документы байт-в-байт идентичны). Токен
    несёт флаг через все line-construction сайты → рендерится под justify/BiDi/пагинацией.
    Write-back маркеров диапазона пока не делается (комменты переживают через reference).
  - **CM2b — нативные `/Text`-аннотации ✓ (opt-in)** — опция `commentAnnotations`: layout
    строит `commentNotes` (anchor `comment-${n}` → {author, contents}) в PdfLayoutAux,
    эмиттер вешает `/Subtype /Text /Name /Comment /T /Contents` рядом с GoTo-ссылкой маркера.
    Строго opt-in и только для интерактивного вывода: подавляется под PDF/A и tagged (где
    нужна annotation/appearance-конформность), по умолчанию отсутствует → veraPDF-гейт и
    байт-снапшоты не тронуты.
  - **Демо ✓** — capstone-тест `reviewedDocx`: один docx со всеми фичами (resolved-родитель +
    тред-ответ, два подсвеченных диапазона, авторы из people.xml, opt-in sticky-notes),
    проверяет весь конвейер (модель + HTML + PDF + docx-roundtrip).
  - Гейты на каждом слайсе зелёные: typecheck 0, lint 0 ошибок, **863 теста**, veraPDF 8/8,
    build, байт-в-ноль. **Хвостов у E-COMMENTS нет.**

---

## E-PIVOT — стиль сводных таблиц Excel (xlsx → PDF/HTML)

**Цель.** Применять СТИЛЬ сводной таблицы (полосатые строки/шапка по `pivotTableStyleInfo`) к
уже-рендеримой сетке pivot; структурная осведомлённость (строки шапки/итогов).

**Усилие: малое→среднее.** Значения pivot УЖЕ рендерятся (см. ниже) — нет нового движка
раскладки, только стиль+структура.

### Главное (проверено): значения pivot уже видны, отсутствует только стиль
- `worksheet-parser.ts:95-141` читает ВСЕ ячейки `sheetData` безусловно. Excel кеширует
  выходные ячейки pivot прямо в лист → они УЖЕ материализуются в SheetDoc и рендерятся как
  обычная сетка. Отсутствует: стиль pivot (полосы/шапка из `pivotTableStyleInfo`) и структура
  (какие строки — субитоги/гранд-итог).

### Архитектурное решение (ДО кода): стиль поверх кешированной сетки, не пересчёт
- **A (выбран): парсим `pivotTable1.xml` (location + styleInfo), резолвим цвета pivot-стиля,
  кладём per-cell shading-карту на диапазон location** — ЗЕРКАЛО Excel-таблиц (SC3): карта
  `print-model.ts:903-917` → слот `shading` на `:470-490`. Переиспользуем резолвер стиля→палитра.
- **B (отвергнут): исполнять движок раскладки pivot из `pivotCacheRecords`** (пересчитать
  сетку) — Excel уже закешировал ячейки; пересчёт = огромный движок, не нужен.
- **C (graceful): нет pivot-партов / неизвестный стиль** → ячейки рендерятся «голыми» (как
  сегодня), без падения.

### Точки привязки (проверено на 3aa7e1a)
- worksheet: `worksheet-parser.ts:81`/`:619` (`parseTableParts`) → `parsePivotTableParts` +
  `pivotTableRelIds`; `:93` (эмит ParsedWorksheet). Ячейки на `:95-141` (безусловно — выход
  pivot уже там).
- модель: `spreadsheet-model/types.ts:80` (ParsedWorksheet), `:105` `tablePartRelIds`, `:107`
  `tables` → добавить `pivotTableRelIds`/`pivotTables`; тип `PivotTable` рядом с `ExcelTable`.
- парс-парт: новый `src/excel/pivot-table-parser.ts` (зеркало `table-parser.ts:22`
  `parseTablePart`) — `location ref` + `pivotTableStyleInfo`(name, firstHeaderRow,
  firstDataRow/Col, showRowStripes/showColStripes).
- резолв стиля: `xlsx-reader.ts:104-111` (tablePartRelIds→parseTablePart→resolveTableStyle) →
  параллель для pivot; `:162` `resolveTableStyle` (regex `TableStyle{Light|Medium|Dark}{N}`) —
  у pivot ОТДЕЛЬНАЯ галерея `PivotStyle{Light|Medium|Dark}{N}` (иной дефолт-маппинг, сверить
  с ECMA §18.10).
- рендер: `print-model.ts:903-917` (tableShadingMap из `worksheet.tables`) → добавить
  pivotShadingMap из `worksheet.pivotTables`; потребляется на `:470-490` (тот же слот `shading`).

### Декомпозиция (вертикальными срезами; нет pivot-партов → байт-в-ноль)
- **PV0 — байт-гейт.** Снапшот PDF-байт/SheetDoc на паре фикстур (`corpus/external/lo-xlsx/
  pivot_dark1.xlsx`, `pivottable_outline_mode.xlsx`) ДО правок — рефактор не должен двигать
  не-pivot выход. (Как E-SHEET SA0.)
- **PV1 — парс `pivotTable1.xml` + модель.** worksheet-rels → `pivotTableRelIds`;
  pivot-table-parser → `PivotTable{ref,name,styleName,firstHeaderRow,firstData*,showRow/Col
  Stripes}`; проводка в ParsedWorksheet. End-to-end: pivot в SheetDoc; стиля нет → байт-в-ноль.
- **PV2 — стиль→цвета + shading-карта.** `resolvePivotStyle` (`PivotStyle…N`→палитра);
  pivotShadingMap на location (шапка `firstHeaderRow..`, полосы данных при `showRowStripes`);
  в слот shading print-model. Сетка pivot рендерится полосатой.
- **PV3 — структура: субитоги/гранд-итог + outline.** парс `rowItems`/`colItems` → пометка
  строк-итогов (отдельная заливка); compact/outline-отступы уже в кешированных ячейках. (Глубина.)
- **PV4 — полиш:** page/filter-поля (шапка), белый текст шапки, демо, scope.md.

### Риски
- Байт-в-ноль: pivot-shading НЕ должен трогать листы без pivot (гейт PV0).
- Галерея pivot-стилей ≠ table-стилей (иные дефолт-цвета — сверить).
- Структурная модель (вложенность `rowItems`) сложна → PV1/PV2 держим плоскими (только
  стиль), структуру в PV3.

### Прогресс
- **PV0 ✓ (свёрнут в PV1).** Отдельный снапшот-харнесс не понадобился: PV1 чисто
  аддитивен (новые опц-поля + rels-дискавери, не трогает чтение ячеек/таблиц), так что
  байт-в-ноль гарантирован конструкцией. Сеть = существующий полный прогон (xlsx-снапшоты) +
  тест «ячейки рендерятся как раньше» + проверка на реальном `pivot_dark1.xlsx`.
- **PV1 ✓** — парс `pivotTable1.xml` + модель. **Ключевое (подтверждено на корпусе):** pivot
  ссылается ТОЛЬКО рельсой листа (`Type=.../pivotTable`), в самом sheet XML элемента нет → не
  трогаем worksheet-parser, перечисляем rels в xlsx-reader. Новый `pivot-table-parser.ts`
  (зеркало `table-parser.ts`) читает `<location ref firstHeaderRow firstDataRow firstDataCol>`
  + `<pivotTableStyleInfo name showRowStripes showColStripes>`; тип `PivotTable` рядом с
  `ExcelTable`; `ParsedWorksheet.pivotTables`; резолв в xlsx-reader (по `isOoxmlRel(...,
  'pivotTable')`) → `grid.pivotTables`. Стиля/шейдинга нет (это PV2) — модель only. Выходные
  ячейки pivot уже кешированы Excel'ем в листе → рендерятся как раньше (байт-в-ноль).
  `build-xlsx` получил опцию `pivotTables` (мерджит rel в sheet-рельсы). Тесты: pivot в
  модели (ref E1:I6→`{0,0,2,2}` на фикстуре, style/firstHeaderRow/stripes), ячейки рендерятся,
  нет pivot→undefined. Проверено на реальном `pivot_dark1.xlsx` (PivotStyleDark1, E1:I6). 840
  тестов (+3). Осталось: PV2 (стиль→цвета + shading-карта), PV3 (структура), PV4 (полиш).
- **PV2 ✓** — стиль→цвета + банды. `resolvePivotStyle` (xlsx-reader, зеркало
  `resolveTableStyle`): `PivotStyle{Light|Medium|Dark}{N}` → accent-колонка `(N-1)%7` →
  header/band-hex (medium/dark — сплошной accent + белый текст; light — тинты); шарит
  `lighten`. Галерея pivot отличается нумерацией от таблиц — берём ту же эвристику как
  аппроксимацию (уточнить в PV4). Рендер: `buildTableFormatLookup` (print-model) вынес общий
  `band(ref, headerRows, style)` и зовёт его для таблиц (`headerRowCount`) И pivot
  (`firstDataRow` как число header-строк) — тот же per-cell shading-слот, что у таблиц SC3.
  Байт-в-ноль: рефактор `band` оставляет таблицы байт-идентичными (полный прогон без сдвигов);
  pivot-шейдинг трогает только листы с pivot, которых в снапшотах нет. Тесты: шапка+полосы по
  `PivotStyleDark2` (band≠header, band1 пуст), без `showRowStripes` — только шапка. Проверено
  на реальном `pivot_dark1.xlsx`: 20 закрашенных ячеек в E1:I6 (grey `7F7F7F` от
  PivotStyleDark1). 842 теста (+2). Осталось: PV3 (субитоги/гранд-итог + outline), PV4 (полиш
  + scope.md).
- **PV3 ✓** — структура: эмфаза строк-итогов. Парсер читает `<rowItems>/<i>@t` →
  `PivotTable.rowItemTypes` (i-й тип = строка данных `firstDataRow + i`; `'grand'`/имя-функции-
  субитога/undefined). Рендер: общий `band` получил опц-предикат `isTotalRow(dataOffset)` —
  строки-итоги (любой `t` ≠ `data`/`blank`, через `isPivotTotal`) красятся как ШАПКА
  (accent + белый текст), а не полосой; таблицы зовут `band` без предиката → байт-идентичны.
  `build-xlsx` pivot-спек получил `rowItemTypes` (эмитит `<rowItems>`). Тест: гранд-итоговая
  строка = цвет шапки (≠ полоса). Проверено на реальном `pivot_dark1.xlsx`: гранд-строка E6:I6
  → `7F7F7F` (шапка), соседняя дата-строка — обычная полоса. Хвост: эмфаза гранд-итоговой
  КОЛОНКИ (`colItems`) и отступы outline. Байт-в-ноль (только листы с pivot). 843 теста (+1).
- **PV4 (частично) ✓ — доки.** `scope.md` синхронизирован под реальность (кросс-катящий, для
  трёх фич сразу): добавлены **Pivot tables** (xlsx), **Review comments** + **SmartArt**
  (docx), **SmartArt** (pptx); из «Not yet» убраны comments/SmartArt/pivot — остались только
  Excel data validation / slicers; SmartArt снят из pptx-loss-списка. Закрывает и хвост
  E-COMMENTS CM3 «scope.md», и стаределость E-SMARTART. Доки-онли (кода нет). Напоминание: на
  сайт попадёт только после мёрджа v1→main + редеплоя (действие пользователя).
- **PV4 ✓ — гранд-колонка; эпик закрыт.** Симметрично PV3, но для КОЛОНОК: парсер читает
  `<colItems>` (общий хелпер `itemTypes` для row/col) → `PivotTable.colItemTypes`; рендер —
  overlay-проход ПОСЛЕ `band` (не трогает общий хелпер → нулевой риск таблицам/строкам):
  тотал-колонки (`isPivotTotal`) перекрываются цветом шапки поверх любой полосы. `build-xlsx`
  pivot-спек получил `colItemTypes` (эмитит `<colItems>`). Тест: гранд-колонка = цвет шапки и
  на небанд-строке, и поверх полосы. На реальном `pivot_dark1.xlsx`: колонка I (грандтотал) вся
  `7F7F7F`; дата-колонка F показывает полосы (`E5E5E5`) + гранд-СТРОКУ (`7F7F7F`) — строки и
  колонки композируются. outline-отступы уже в кешированных ячейках (бесплатно). Байт-в-ноль
  (overlay только для pivot с тотал-колонками). 844 теста (+1). **E-PIVOT закрыт (PV0–PV4).**
  Хвост (не блокирует): точная галерея pivot-стилей (vs аппроксимация), демо-артефакт.

---

## Сводка приоритетов

| Эпик    | Усилие        | byteRisk | Связь с core-миссией       | Когда                    |
|---------|---------------|----------|----------------------------|--------------------------|
| E-DOCX  | среднее       | низкий   | docx-выход, roundtrip-гейт | **первым (сейчас)**      |
| E-SHEET | большое       | высокий  | Excel первоклассен → ядро  | вторым (стратегический)  |
| E-PDF   | очень большое | н/д      | универсальность, не ядро   | отдельный крупный заход  |
| E-PPTX  | среднее       | низкий   | pptx-вход, замыкает OOXML  | новая эра (после 1.8.0)   |
| E-PARITY| малое→среднее | низкий¹  | визуальный паритет с Word/LO | после E-PPTX            |
| E-SMARTART | малое→среднее | низкий | SmartArt в docx+pptx (был пробел) | ✓ закрыт (SA0–SA3) |
| E-COMMENTS | малое→среднее | низкий | ревью-комментарии docx (был пробел) | ✓ закрыт ПОЛНОСТЬЮ (рендер+клик+write-back+треды/resolved+подсветка диапазона+people.xml+opt-in /Text+демо) |
| E-PIVOT | малое→среднее | низкий² | стиль сводных Excel (значения уже видны) | ✓ закрыт (PV0–PV4) |

² E-PIVOT: значения pivot уже рендерятся; риск только в том, чтобы shading не задел листы без pivot (гейт PV0).

¹ E-PARITY: FP1 (подстановка) байт-стабилен; FP2–FP4 — строго опт-ин `layoutProfile`, дефолт `'ream'` не меняется.
