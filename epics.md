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
- **EP4–EP5** — эвристика для нетегированных PDF (глифы→строки→параграфы по координатам) → проводка
  `pdfReader` в фасад (Ream.parse сниффит %PDF-) + losses.

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

## Сводка приоритетов

| Эпик    | Усилие        | byteRisk | Связь с core-миссией       | Когда                    |
|---------|---------------|----------|----------------------------|--------------------------|
| E-DOCX  | среднее       | низкий   | docx-выход, roundtrip-гейт | **первым (сейчас)**      |
| E-SHEET | большое       | высокий  | Excel первоклассен → ядро  | вторым (стратегический)  |
| E-PDF   | очень большое | н/д      | универсальность, не ядро   | отдельный крупный заход  |
