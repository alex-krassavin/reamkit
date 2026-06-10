# IR v1 — design doc (спека, не реализация)

> Статус: **draft для обсуждения**. Решения-основания зафиксированы в
> `handoff.md` § «v1 — слой IR». Сигнатуры ниже — TypeScript-формы для спора об
> именах и контрактах; код по ним не пишется, пока док не согласован.

## 1. Цели / не-цели

**Цели v1:**
- Развязать parser ↔ representer: входные адаптеры (bytes → дерево) и выходные
  (дерево → bytes) за публичными интерфейсами; конвертация = композиция.
- Два дерева (FlowDoc / PageDoc) + направленные трансформы между ними.
- Cross-format fidelity: docx→pdf (паритет с текущим), docx→html, pagedoc→svg —
  третий адаптер как проверка дизайна.
- FontProvider-цепочка; LossReport + `strict`.

**Не-цели v1 (двери держим, но не входим):**
- Round-trip fidelity (docx→IR→docx ≈ исходник) — passthrough-мешки закладываем,
  гарантий не даём.
- reconstruct: PageDoc→FlowDoc (PDF→Word) — отдельная эра.
- SheetDoc (третье дерево для сетки xlsx) — техдолг, пока xlsx-reader проецирует
  в FlowDoc как сейчас.
- SAX-стриминг — материализация + self-limit'ы.

## 2. Словарь

| Термин | Что это |
|---|---|
| **FlowDoc** | Семантическое дерево без страниц/координат (параграфы, раны, таблицы, списки, секции). Родной уровень DOCX/HTML/MD. |
| **PageDoc** | Пагинированное дерево: страницы + позиционированный контент + логическая структура. Родной уровень PDF/SVG/XPS/растра. |
| **Reader** | Адаптер `bytes → FlowDoc \| PageDoc`. |
| **Writer** | Адаптер `FlowDoc \| PageDoc → bytes`. |
| **layout** | Трансформ FlowDoc → PageDoc (Knuth-Plass, таблицы, пагинация — текущий движок). |
| **FontProvider** | Ступень резолва шрифта `(family, style, codepoints) → байты/метрики`. |
| **ResourceStore** | Content-addressed хранилище бинарей (картинки, шрифты); ноды ссылаются по id. |
| **LossReport** | Список потерь конвертации: `dropped \| degraded \| substituted`. |

## 3. Карта

```
DOCX ──reader──▶ FlowDoc ──writer──▶ DOCX / HTML / MD      (генерация, без layout)
XLSX ──reader──▶    │                 (проекция сетки; SheetDoc — техдолг)
                    │ layout(opts: fonts, hyphenator, …)
                    ▼
                 PageDoc ──writer──▶ PDF / SVG / PNG
PDF  ──reader──▶ PageDoc ──reconstruct──▶ FlowDoc          (вне v1)
```

## 4. Общие типы

```ts
/** Канонические пункты (1/72"). Единственная единица длины во всех деревьях. */
type Pt = number & { readonly __brand: 'pt' };

/** IR = дерево + ресурсы. Дерево — чистый JSON (диффабельно, сериализуемо). */
interface FlowDoc  { readonly kind: 'flow';  readonly body: ...; readonly resources: ResourceStore; }
interface PageDoc  { readonly kind: 'page';  readonly pages: ...; readonly resources: ResourceStore; }

interface ResourceStore {
  /** id = хэш байтов (content-addressed) → дедупликация бесплатно. */
  get(id: ResourceId): Uint8Array | undefined;
  put(bytes: Uint8Array): ResourceId;
  ids(): readonly ResourceId[];
}

/** Passthrough для round-trip-двери: нейтральное ядро + формат-скоупный сырец. */
interface NativeBag { readonly [formatId: string]: unknown } // напр. { ooxml: <фрагмент> }
// Каждая нода МОЖЕТ нести `native?: NativeBag`; писатели чужих форматов игнорируют.
```

## 5. FlowDoc — из `document-model`, с нейтрализацией

База — текущий `BodyElement[]` (оба конвертера уже сходятся в него). Изменения:

1. **Единицы**: все `*Twips`, `*EighthPt`, `*HalfPt` → `Pt` (в reader'ах
   конвертация на входе; layout уже считает во float).
2. **Имена**: убрать OOXML-измы из публичной схемы (`w:`-семантика остаётся в
   reader'е). Примеры: `vMerge` → `rowSpan` (резолвится reader'ом),
   `gridSpan` → `colSpan`, `sectPr`-поля → `Section`.
3. **Стили — resolved**: на ранах/параграфах лежат эффективные свойства;
   опциональная `styles`-секция (id → свойства) для writers, генерящих
   именованные стили. Каскад — приватное дело docx-reader'а.
4. **Ресурсы**: `InlineImage.bytes` → `InlineImage.resource: ResourceId`.
5. **`native?: NativeBag`** на нодах (см. §4).

Скелет (имена — на обсуждение):

```
FlowDoc
└─ sections: Section[]          // pageSize/margins/headers/footers (Pt)
   └─ blocks: Block[]           // Paragraph | Table | Figure | MathBlock | …
      Paragraph { runs: Run[], props: ResolvedParaProps, native? }
      Run       { text, props: ResolvedRunProps, native? }   // | InlineImage | MathInline
      Table     { rows, grid: Pt[], props, native? }         // rowSpan/colSpan уже резолвлены
```

## 6. PageDoc — публикация `DrawCommand`

База — приватный `DrawCommand[]` styled-renderer'а. Публикуем как:

```
PageDoc
├─ pages: Page[]
│  └─ { size: {w: Pt, h: Pt}, items: PageItem[] }
│     PageItem = Text | Rect | Image | VectorShape   // ← типизированный DrawCommand
│     Text { glyph-run: font(ResourceId|FaceRef), size, x/y, text }  // для extraction/ActualText
├─ structTree?: StructNode    // логическая структура (tagged) — уже есть, наружу
└─ resources: ResourceStore   // шрифты (subset), картинки
```

Принципы: PageItem самодостаточен (writer SVG/PNG не знает про OOXML); текст
несёт и глифы, и исходную строку (поиск/копирование/ActualText); координаты —
`Pt`, origin — **top-left** (как CSS/SVG; PDF-writer сам флипает в свой y-up —
сейчас флип уже внутри рендерера).

## 7. Интерфейсы адаптеров

```ts
interface DocumentReader<D extends FlowDoc | PageDoc> {
  readonly id: string;                       // 'docx', 'xlsx', 'pdf'
  readonly produces: D['kind'];              // 'flow' | 'page'
  readonly supports: ReadonlySet<Feature>;   // §9
  sniff(bytes: Uint8Array): boolean;         // магия формата (PK…, %PDF-)
  read(bytes: Uint8Array, opts?: ReadOptions): ReadResult<D>;   // SYNC (§ handoff-4)
}

interface DocumentWriter<D extends FlowDoc | PageDoc> {
  readonly id: string;                       // 'pdf', 'html', 'svg', 'docx'
  readonly consumes: D['kind'];
  readonly supports: ReadonlySet<Feature>;
  write(doc: D, opts?: WriteOptions): WriteResult;              // SYNC
}

interface ReadResult<D>  { doc: D;          losses: LossReport; }
interface WriteResult    { bytes: Uint8Array; losses: LossReport; }

/** layout — тоже именованный трансформ (не writer и не reader). */
interface LayoutOptions { fonts: FontResolver; hyphenator?: Hyphenator; /* … */ }
function layout(doc: FlowDoc, opts: LayoutOptions): { doc: PageDoc; losses: LossReport };
```

**Фасад** (единственное async-место — добыча шрифтов/ресурсов):

```ts
const ream = createConverter({ readers: […], writers: […], fonts: providerChain });
const r = await ream.convert(bytes, { to: 'pdf', strict?: boolean, pdfA?, … });
// r: { bytes, losses } — strict: true ⇒ throw ConversionLossError на первой потере
```

Текущие `convertDocxToPdf*` остаются как тонкие обёртки над фасадом
(API-совместимость, корпус-гейты byte-identical).

## 8. FontProvider

```ts
interface FontRequest  { family: string; bold: boolean; italic: boolean; codepoints?: ReadonlySet<number>; }
type FontAnswer =
  | { kind: 'bytes';   bytes: Uint8Array; faceName: string }        // полный резолв
  | { kind: 'metrics'; metrics: FaceMetrics; embedFallback: FontRequest } // canvas-metrics (эксперимент)
  | { kind: 'none' };

interface FontProvider {
  readonly id: 'embedded' | 'caller' | 'local' | 'remote' | 'canvas-metrics' | string;
  resolve(req: FontRequest): Promise<FontAnswer>;   // async: local/remote по природе
}
// Цепочка: embedded → caller → local(fsType-фильтр!) → remote → [canvas-metrics] 
// Всё ниже точного байтового резолва → losses.push({severity:'substituted', …}).
```

`fsType`-правило (OS/2): `restricted` ⇒ провайдер обязан вернуть `none`
(встраивание запрещено лицензией); `preview&print` ⇒ допускаем subset-embed
(PDF-печать — каноничный preview&print кейс), помечаем в losses.

## 9. Features и LossReport

```ts
type Feature = string;   // иерархический словарь: 'tables', 'tables.nested', 'math',
                         // 'rtl', 'images.jp2', 'charts.scatter', 'trackedChanges', …
// Константный реестр в ядре; capability-matrix в доки генерится из supports-сетов.

interface Loss { severity: 'dropped' | 'degraded' | 'substituted';
                 feature: Feature; detail: string; where?: string /* page/para ref */ }
type LossReport = readonly Loss[];
```

## 10. Миграция (каждый этап — зелёный корпус, byte-identical PDF)

| Этап | Что делаем | Гейт |
|---|---|---|
| 1 | Типы ядра: `Pt`, `ResourceStore`, `Loss*`, `Feature` — рядом, ничего не трогаем | tsc + тесты |
| 2 | `FlowDoc` = нейтрализованный document-model (единицы/имена/resources); docx/xlsx-парсеры выдают его | **PDF байт-в-байт** |
| 3 | `PageDoc`: расщепить styled-renderer на `layout(FlowDoc)→PageDoc` и `pdfWriter(PageDoc)→bytes` | **PDF байт-в-байт** |
| 4 | Интерфейсы Reader/Writer/фасад; старые API — обёртки | публичный API не ломается |
| 5 | FontProvider-цепочка (рефактор) + **local**-провайдер с fsType | новые тесты + корпус |
| 6 | Третий адаптер: **svg-writer(PageDoc)** (превью) и/или **html-writer(FlowDoc)** | вскрытие ошибок интерфейсов, фиксация `@experimental`-схемы |

Этап 3 — самый рискованный (внутренности styled-renderer), поэтому отдельный и
с самым строгим гейтом. Порядок 2↔3 можно обернуть.

## 11. Отвергнутые альтернативы (кратко, мотивировка в handoff)

- Один универсальный IR на оба уровня (путь UNO) — union всех фич, всё optional.
- SAX-стриминг как базовая модель — убивает трансформы/дифф/reconstruct.
- Автотрейс пикселей в шрифт — качество/объём/лицензии.
- Round-trip как обещание v1 — ×3 сложности; passthrough-мешки вместо этого.
