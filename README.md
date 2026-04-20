# xlsx Read-Only Viewer (fortune-sheet + fortune-excel)

A React + TypeScript PoC that renders an Excel `.xlsx` file in the browser
with as much of the original formatting (fills, fonts, borders, merges,
wrap, alignment) preserved as possible — styled as a **read-only viewer**
for internal reports, not an editor.

This document is a full implementation spec. An engineer should be able to
reimplement the project from scratch using this README alone.

---

## 1. What it does

- User picks a local `.xlsx` (or `.csv`) file via a file input.
- The file is parsed in the browser — no server round-trip.
- The workbook is rendered inside a `@fortune-sheet/react` `<Workbook>`
  component, **with editing affordances removed**: no toolbar, no formula
  bar, no native sheet tabs, no context menus, no cell editing.
- The rendered viewport shrinks to the actual data extent when the sheet
  is small, and scrolls via the outer container when the data is larger
  than the browser viewport.
- Several known gaps in `@corbe30/fortune-excel`'s xlsx → fortune-sheet
  conversion are patched client-side by a chain of post-processors
  (see §6).

### Sample files used during development

Two real xlsx files exercise different edge cases. They live one
directory above this project root so they're trivially shareable:

- `/home/felix/IPD/sheets-test-fortune/example-formatted.xlsx`
  Small single-sheet workbook with an **Excel Table** (`D13:G20`,
  style "TableStyleLight1 2"). Exercises the table-style post-processor.

- `/home/felix/IPD/sheets-test-fortune/derm-summary-example.xlsx`
  Six-sheet clinical benchmarking workbook with inlineStr rich text,
  custom row heights, zoomed sheet views (`zoomScale="55"`),
  overflow-style cells, and a merged header row. Exercises almost every
  post-processor.

---

## 2. Stack

| Concern | Choice | Why |
|---|---|---|
| Bundler/dev server | **Vite** 7 | Fast HMR, native ESM, plugin ecosystem. |
| UI framework | **React** 19 | Team standard. |
| Language | **TypeScript** 5.8 | Strict types across our post-processors. |
| Styling | **Tailwind CSS** 4 | Utility-first; minimal custom CSS. |
| xlsx → luckysheet data | **`@corbe30/fortune-excel`** 2.3 | Exposes `transformExcelToFortune` and wraps ExcelJS. Community fork of FortuneSheetExcel. |
| Spreadsheet renderer | **`@fortune-sheet/react`** 1.0 | Canvas-based renderer; accepts luckysheet-shaped `Sheet[]` data. |
| Zip / XML access for post-processors | **JSZip** + DOMParser | JSZip is already a transitive dep of `@corbe30/fortune-excel`; no new install. |
| Lint | **ESLint** 9 + `@eslint-react/eslint-plugin` + `typescript-eslint` + `@stylistic/eslint-plugin` + Prettier compat | Included in the Vite template. |
| Format | **Prettier** 3 + `prettier-plugin-tailwindcss` | Class-name ordering. |

**Deliberately rejected / removed during the PoC:**

- `xlsx-populate` + `react-spreadsheet` — evaluated as an alternative
  rendering strategy in an earlier revision; dropped because
  `react-spreadsheet` is an editable data grid, not a formatting-faithful
  viewer (no borders/merges/variable sizing/number formats), and
  `xlsx-populate` doesn't resolve theme colors. See §6.0 for what the
  comparison revealed and why we kept fortune-sheet.
- Swapping the whole stack for **Univer** (the actively-maintained
  successor to luckysheet). Univer has better xlsx feature coverage
  natively — if the post-processor stack here grows unwieldy, that's
  the first escape hatch. For now the targeted post-processors are
  cheaper than a full rewrite.

---

## 3. From zero — full setup

### 3.1 Prerequisites

- Node.js **20.19+** or **22.12+** (Vite 7 requirement; `package.json`'s
  `engines` field enforces this).

### 3.2 Scaffold

Start from the base Vite + React + TS + Tailwind template this repo was
cloned from:

```sh
tiged royrao2333/template-vite-react-ts-tailwind fortune-sheet
cd fortune-sheet
```

or create the equivalent structure manually: `vite create` with
`react-swc-ts`, add `tailwindcss` + `@tailwindcss/vite`, copy the lint
config from this repo's `eslint.config.mjs`.

### 3.3 Install runtime dependencies

```sh
npm install @corbe30/fortune-excel@2.3.2 @fortune-sheet/react@1.0.4 \
           react@19.1.1 react-dom@19.1.1
```

`@corbe30/fortune-excel` drags in `exceljs`, `jszip`, and a handful of
older transitive deps that print deprecation warnings during install
(`inflight`, `rimraf@2`, `glob@7`, `lodash.isequal`, `fstream`). They're
coming from ExcelJS and are unavoidable without forking the package.

### 3.4 Install dev dependencies

Already pulled in by the template. Key choices:

- `vite@7.3.2`, `@vitejs/plugin-react-swc@4`, `@tailwindcss/vite@4.1.12`,
  `tailwindcss@4.1.12`.
- `typescript@5.8.3`, `@types/node@20.19.0` (must satisfy `vite`'s
  `peerOptional @types/node: ^20.19.0 || >=22.12.0` — earlier `20.14.8`
  triggers ERESOLVE).
- `@types/react@19.1.12`, `@types/react-dom@19.1.9`.
- Lint stack: `eslint@9`, `@eslint/js`, `typescript-eslint`,
  `eslint-plugin-react-hooks`, `eslint-plugin-react-refresh`,
  `@eslint-react/eslint-plugin`, `@stylistic/eslint-plugin`,
  `eslint-config-prettier`, `vite-plugin-eslint2`.
- Format: `prettier`, `prettier-plugin-tailwindcss`.

### 3.5 `.npmrc`

```
save-exact=true
```

Locks dependency versions on `npm install`. **Do not** add
`auto-install-peers=true` — that's pnpm-only syntax and npm emits an
"Unknown project config" warning on every command otherwise.

### 3.6 Scripts

```json
"scripts": {
  "dev": "vite",
  "build": "tsc -b && vite build",
  "lint": "eslint .",
  "preview": "vite preview"
}
```

### 3.7 Run it

```sh
npm install
npm run dev        # http://localhost:5173
```

---

## 4. Architecture

```
index.html ──► index.tsx ──► src/App.tsx ──► src/components/FortuneSheet.tsx
                                                   │
                                                   ├── @corbe30/fortune-excel
                                                   │     transformExcelToFortune(file, capture, noop, shimRef)
                                                   │
                                                   └── four post-processors (src/excel/*)
                                                         1. applyTableStyles         (table fills/borders)
                                                         2. backfillInlineStringV    (inlineStr overflow fix)
                                                         3. forceWrapOnOverflowCollisions  (neighbor-collision wrap)
                                                         4. bindDisplayBounds        (trim blank rows/cols)
                                                         ► sizes computed by measureSheet
                                                         ► final Sheet[] ► <Workbook>
```

### Data flow per file load

1. User picks a file via `<input type="file">` in `App.tsx`.
2. `App.tsx` puts the `File` into state and passes it to `<FortuneSheet>`.
3. `FortuneSheet` runs an effect keyed on `file`.
4. Effect calls `transformExcelToFortune(file, capture, noopSetKey, shimRef)`:
   - `capture` stashes the raw converted `Sheet[]` in a closure variable.
   - A **no-op setKey** is passed in so fortune-excel's internal
     `setKey` call after conversion doesn't force a premature Workbook
     remount with the unpatched data.
   - A **shim ref** (`{ current: null }`) is passed so the library's
     setTimeout-scheduled `sheetRef.current.setColumnWidth(...)` pass
     becomes a no-op. That imperative pass was crashing on multi-sheet
     files with `"sheet not found"` (race vs the Workbook's multi-sheet
     init). fortune-sheet reads column widths from `sheet.config.columnlen`
     during render anyway, so the imperative call is redundant.
5. After the fortune-excel promise resolves, the raw sheets run through
   the four post-processors (§6) in order.
6. A single `dispatch({ type: 'load', sheets })` updates the reducer
   state with the patched sheets and bumps `state.key`, which forces
   exactly one `<Workbook>` remount.
7. `FortuneSheet` renders the patched active sheet inside a two-div
   layout: an outer `overflow-auto` container capped at the viewport,
   and an inner container sized to the measured data footprint. Inside
   that, `<Workbook>` receives read-only settings (`allowEdit={false}`,
   `showToolbar={false}`, etc.).

### State model (`FortuneSheet.tsx`)

Three pieces of related state consolidated into a `useReducer` so the
"reset on file change" branch of the effect can call `dispatch` instead
of three separate `setX` calls (the latter tripped
`@eslint-react/hooks-extra/no-direct-set-state-in-use-effect`):

```ts
type ViewerState = { sheets: Sheet[]; activeSheetIndex: number; key: number };

type ViewerAction =
  | { type: 'reset' }                                   // file cleared
  | { type: 'load'; sheets: Sheet[] }                   // patched sheets ready
  | { type: 'selectSheet'; index: number };             // future use: sheet selector dropdown
```

The Workbook `key` is `${state.key}-${safeIndex}` so any reducer action
that bumps `key` or changes `activeSheetIndex` triggers a clean remount.

---

## 5. File-by-file reference

### 5.1 Top-level files

- **`index.html`** — Vite entry. Single `<div id="root">` and a `<script
  type="module" src="/index.tsx">`. No changes from the template.

- **`index.tsx`** — React bootstrap. `createRoot(#root).render(<StrictMode><App/></StrictMode>)`
  with `./src/styles/global.css` imported for Tailwind + our overrides.

- **`vite.config.ts`** — registers `@vitejs/plugin-react-swc`,
  `@tailwindcss/vite`, `vite-plugin-eslint2`. Adds an `@` alias to
  `./src` (not currently used, but kept for future imports).

- **`tsconfig.*`** — standard Vite React strict config. `verbatimModuleSyntax`
  is on, which is why some imports are `import type { ... }`.

- **`eslint.config.mjs`** — flat config pulling in `js.configs.recommended`,
  `tseslint.configs.recommended`, `reactHooks.configs['recommended-latest']`,
  `eslintReact.configs['recommended-typescript']`, and `eslint-config-prettier`.
  The `@eslint-react/hooks-extra/no-direct-set-state-in-use-effect` rule is
  the one that pushed us toward `useReducer` in `FortuneSheet.tsx`.

- **`prettier.config.mjs`** — `prettier-plugin-tailwindcss` for class-name
  sorting.

- **`package.json`** — dependencies as listed in §3.

- **`.npmrc`** — `save-exact=true`. No `auto-install-peers`.

### 5.2 `src/App.tsx`

The top-level page. Holds the `file: File | null` state, renders a
`<header>` with the file picker, and the single framed `<section>`
containing `<FortuneSheet>`.

Layout notes:

- The wrapper is `flex flex-col items-start gap-4` so the section
  doesn't stretch across the page — `items-start` lets `<section
  w-fit>` actually shrink to the viewer's content width.
- The section uses `w-fit max-w-full` so it hugs content but never
  exceeds the page width.

### 5.3 `src/components/FortuneSheet.tsx`

The main component. Responsibilities:

1. Own the `useReducer` view state (`sheets`, `activeSheetIndex`, `key`).
2. Load the file via `transformExcelToFortune` + post-processors inside
   a `useEffect([file])`.
3. Render the `<Workbook>` with all read-only flags.
4. Size the wrapper via `measureSheet(activeSheet)` memoized on
   `activeSheet`.

Two important workarounds live here:

- **`const shimRef = { current: null }`** — see §4 step 4.
- **`const noopSetKey = () => {}`** — passed to fortune-excel so only
  our own `dispatch({ type: 'load' })` bumps the key. Prevents a
  "flicker" where the raw (pre-patched) data rendered for one frame
  before the post-processors finished.

A commented-out `<select>` dropdown for multi-sheet navigation is
retained for future use; the supporting reducer action (`selectSheet`)
is already in place.

### 5.4 `src/excel/xlsxZip.ts`

Wraps `JSZip.loadAsync(file)`. Reads every non-directory entry as text
into a `Map<path, string>` up front (eager because XLSX files are
small and we make multiple passes). Exposes:

- `getXmlDoc(path)` — parses the text at `path` through `DOMParser`
  (`'application/xml'`), returning `null` if missing or parse-error.
- `getText(path)`, `listPaths()` — helpers.

All the subsequent parsers operate on `Document`s returned from this.

### 5.5 `src/excel/theme.ts`

Parses `xl/theme/theme1.xml` into a `ThemePalette` (12 slots:
`[dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink]`). Falls back to the
Office 2013+ default palette if theme1.xml is missing or incomplete.
Matches the namespace-agnostic XML structure via `localName`.

### 5.6 `src/excel/colors.ts`

- `ColorRef` type — captures the union of color forms found in OOXML
  (`{ rgb?, theme?, tint?, indexed?, auto? }`).
- `applyTint(hex, tint)` — OOXML §18.8.19 HLS luminance shift. Accepts
  `tint` in `[-1, 1]`.
- `resolveColor(ref, palette)` — returns a `#RRGGBB` string or `undefined`.
  Applies the **SpreadsheetML theme-index swap**: `theme=0` → `lt1`,
  `theme=1` → `dk1` (the clrScheme XML order is the opposite).
- `readColorRef(el)` — extracts a `ColorRef` from any OOXML color element.

### 5.7 `src/excel/dxf.ts`

Parses the `<dxfs>` block from `xl/styles.xml`. Each dxf can have
`fill`, `font`, and `border` sub-entries. Used by `applyTableStyles`
for Excel table styles (see §6.2).

### 5.8 `src/excel/tableStyles.ts`

Parses `<tableStyles>` from `xl/styles.xml` into a
`Map<styleName, { elements: Partial<Record<TableElementType, dxfId>> }>`.
Element types include `wholeTable`, `headerRow`, `firstColumn`,
`firstColumnStripe`, `firstHeaderCell`.

### 5.9 `src/excel/tables.ts`

- `parseRef(a1Range)` — converts `"D13:G20"` to `{ r1, c1, r2, c2 }`
  (0-based).
- `parseTable(tableDoc)` — extracts the table's name, ref, style info
  flags (`showFirstCol`, `showHeaderRow`, stripe flags).
- `indexSheetTables(getXmlDoc, listPaths)` — walks
  `xl/workbook.xml` + `xl/_rels/workbook.xml.rels` + each
  `xl/worksheets/_rels/sheetN.xml.rels` to build a
  `Map<sheetName, Document[]>` of all table definitions per sheet.
  Browsers disagree on how `r:id` namespaced attributes surface, so
  the lookup tries three fallbacks.

### 5.10 `src/excel/applyTableStyles.ts`

**The big one.** Post-processor that fills a feature gap in
`@corbe30/fortune-excel`: it doesn't apply Excel Table Styles at all.

For each table in each sheet, it walks every cell in the table's range
and layers dxfs in OOXML precedence:

```
wholeTable → (headerRow if r == r1) → (firstColumn if c == c1)
           → (firstHeaderCell if r == r1 && c == c1)
```

Resolved `bg` / `fc` / `bl` / `it` / `un` are written to the cell's
`v`. Borders from `wholeTable.horizontal` and `firstColumn.horizontal`
are emitted as **per-row `border-bottom` entries** in
`sheet.config.borderInfo` — not as a single `border-horizontal` range,
because that leaks beyond the table range in luckysheet's renderer.
The first row of internal borders is skipped so the "empty bottom" on
`headerRow` is honored.

Cells inside a table **overwrite** existing cell-level styles. That's
because fortune-excel mis-resolves certain theme-colored fonts (e.g.,
D13 in the sample file came out white). Inside a table, the authored
table style is the authoritative visual; outside, cell-level wins.

### 5.11 `src/excel/backfillInlineStringV.ts`

Post-processor that fixes a cross-path inconsistency in fortune-sheet
itself (not fortune-excel).

- fortune-sheet's cell-render path uses `isInlineStringCell(cell)` to
  decide whether a cell has content (checks `ct.t === "inlineStr"`).
- fortune-sheet's overflow-trace path (`cellOverflow_trace` in
  `@fortune-sheet/core:73701`) uses `!_.isEmpty(cell.v)` instead —
  which returns "empty" for inlineStr cells because their text lives
  in `cell.v.ct.s[i].v`, not in `cell.v.v`.

Consequence: text from an overflow-mode neighbor bleeds over an
inlineStr cell that fortune-sheet thinks is empty (see §6.3). We
backfill `cell.v.v` with the concatenated run text so the trace check
matches the render check. Rich-text rendering stays unaffected
because it comes from `ct.s`.

### 5.12 `src/excel/forceWrapOnOverflowCollisions.ts`

Post-processor that tackles the Q3-style rendering bug where a
center-aligned overflow cell's text paints across a neighbor that has
content but is narrow enough that fortune-sheet's canvas draw escapes
the cell bounds.

Flow:

1. For each cell with `tb === "1"` (overflow mode) and a text value.
2. Use `canvas.measureText` with the cell's font to get the real pixel
   width.
3. If that width exceeds the column width less a small padding budget,
   check the relevant neighbor(s) — left, right, or both, depending on
   the cell's horizontal alignment (`ht`: `0` center, `1` left, `2` right).
4. If a neighbor is non-empty (via `extractText`, which handles both
   plain-value and inlineStr cells), flip the cell's `tb` to `"2"` so
   fortune-sheet wraps rather than overflows.

A shared hidden `<canvas>` is reused across calls.

### 5.13 `src/excel/bindDisplayBounds.ts`

Post-processor that constrains the visible grid extent. fortune-excel
leaves `sheet.row` / `sheet.column` at large defaults (hundreds of
blank cells past the data). For a read-only viewer we want the grid
to end shortly after the last populated cell.

- Walks `celldata` for `max(r, c)`.
- Folds merged-cell rectangles (`sheet.config.merge`) into the same
  max so merges that stick out past the last populated cell don't get
  clipped.
- Writes `sheet.row = maxR + 2`, `sheet.column = maxC + 2`
  (`+2` leaves one buffer row/col of whitespace so the last data cell
  isn't flush against the edge), and `sheet.addRows = 0`.

### 5.14 `src/excel/measureSheet.ts`

Returns the `{ width, height }` in pixels of the sheet's content area,
mirroring fortune-sheet's internal `calcRowColSize`
(`@fortune-sheet/core:63469`): each row/column contributes
`round((len + 1) * zoom)`, where `zoom` comes from
`sheet.zoomRatio` (set from the xlsx's `<sheetView zoomScale="...">`).

Constants mirror the fortune-sheet defaults at
`@fortune-sheet/core:63380-63387`: `rowHeaderWidth=46`,
`columnHeaderHeight=20`, `defaultcollen=73`, `defaultrowlen=19`.
A `STAT_AREA_HEIGHT=22` accounts for the bottom status strip.

Used by `FortuneSheet.tsx` (memoized on `activeSheet`) to size the
inner wrapper div. The outer wrapper caps that at
`calc(100vw - 4rem)` / `calc(100vh - 12rem)` and handles scrolling when
the data exceeds the viewport.

### 5.15 `src/styles/global.css`

Imports Tailwind. Adds one override:

```css
.luckysheet-scrollbar-x,
.luckysheet-scrollbar-y { display: none !important; }
```

fortune-sheet's internal `ch_width += 120 (maxColumnlen)` and
`rh_height += 80` trailing buffers cause its scroll tracker elements
to always be slightly wider/taller than the canvas — so the library
always shows scrollbars, even when the full data footprint is
visible. These hidden elements aren't the real scroll controls for our
case (the outer wrapper's `overflow-auto` is), so we hide them.

---

## 6. Why four post-processors? — the problems they solve

### 6.0 Why not `xlsx-populate` + `react-spreadsheet` instead?

Initial PoC evaluated that pair. Findings:

- `react-spreadsheet` uses `<table>` with fixed layout. Cannot render
  merged cells, per-cell borders, variable row heights / column widths,
  or number formats. At best you can set background / font via a
  custom `DataViewer`. For "faithful Excel rendering" this is a hard
  stop.
- `xlsx-populate` correctly reads styles but doesn't resolve theme
  colors — themed cells (the overwhelming majority in real-world
  xlsx) render with undefined colors.

Conclusion: `fortune-sheet` is the only viable choice among
browser-JS spreadsheet renderers that accept a data blob and draw
something Excel-faithful. The trade-off is we need to work around its
bugs (that's what the post-processors are).

### 6.1 Excel Table Styles are ignored

`@corbe30/fortune-excel` has **zero** handling of `<tableStyles>` /
`<dxfs>` / `<tableStyleInfo>`. The sample file `example-formatted.xlsx`
defines a Table on `D13:G20` with fills (`theme=1 tint=0.25`
lavender), white font color, and thin horizontal rules coming from
the table style. Without our `applyTableStyles` post-processor, the
table range renders as blank white cells.

### 6.2 `"sheet not found"` on multi-sheet imports

fortune-excel schedules `setTimeout(() => sheetRef.current.setColumnWidth({ id: sheet.id })` `, 1)`
at the end of its transform. For multi-sheet files, this fires before
the Workbook has registered all sheet ids; `setColumnWidth` calls
`getSheet()` which throws `SHEET_NOT_FOUND`. Fix: pass
`{ current: null }` — fortune-excel uses optional chaining, so the
call is skipped. Column widths still land because fortune-sheet reads
`sheet.config.columnlen` during its own render.

### 6.3 inlineStr cells don't block overflow traces

Detailed above in §5.11. Patched by `backfillInlineStringV`.

### 6.4 Overflow text painting into neighbors' rectangles

fortune-sheet's `cellOverflow_trace` correctly refuses an overflow
entry when the neighbor has content, but the canvas text draw call
doesn't clip to the cell bounds. For narrow cells with long text and
populated neighbors, the painted glyphs bleed over. Patched by
`forceWrapOnOverflowCollisions` (flipping `tb` from `"1"` to `"2"`
confines the text to the cell via wrapping).

### 6.5 Phantom scroll trailing buffer

fortune-sheet adds +120 px to `ch_width` and +80 px to `rh_height`
inside its `calcRowColSize`. That makes the scroll tracker larger
than the canvas → scrollbars are always present. Patched by the CSS
override in §5.15.

### 6.6 Viewport overflows when data is bigger than the screen

Since the `<Workbook>` is canvas-based and its canvas size is latched
to the placeholder element at mount, the wrapper must be at least as
big as the content or cells get clipped. Solution: two nested divs —
inner one sized to content (from `measureSheet`), outer one
`overflow-auto` capped at the viewport. The browser's own scrollbars
on the outer div provide real scrolling, because the inner content
genuinely extends past the outer.

---

## 7. Read-only rendering settings

`<Workbook>` receives:

```tsx
<Workbook
  allowEdit={false}
  showToolbar={false}
  showFormulaBar={false}
  showSheetTabs={false}
  cellContextMenu={[]}
  headerContextMenu={[]}
/>
```

- `allowEdit={false}` blocks F2/Enter/Delete/paste/contentEditable.
- Empty `cellContextMenu` / `headerContextMenu` arrays cause
  `handleContextMenu` in `@fortune-sheet/core:77727` to
  `e.preventDefault()` and return — no menu is rendered.
- Row / column indicator headers (`A/B/C…`, `1/2/3…`) are retained
  because fortune-sheet has no setting to hide them and they help
  readers orient.

---

## 8. Limitations / known gaps

- **Excel Tables**: post-processor handles `wholeTable`, `headerRow`,
  `firstColumn`, `firstHeaderCell`, `firstColumnStripe`. Row stripes
  and column stripes are not applied.
- **Conditional formatting**: not supported (fortune-excel doesn't
  parse it; we don't patch it in).
- **Number formats from dxfs**: not applied.
- **Non-default theme files**: we parse `xl/theme/theme1.xml`, so
  custom themes work. Theme palette fallback kicks in only when the
  theme file is missing.
- **Rich text within a single cell**: renders via fortune-excel's
  `inlineStr` path; our `backfillInlineStringV` preserves that path.
- **Pixel-exact Excel rendering**: impossible without forking
  fortune-sheet's canvas draw path. Our goal is "looks like the
  original to a human reader", not "byte-for-byte identical to Excel".

---

## 9. Verifying a working build

From a fresh clone:

```sh
npm install
npm run dev           # http://localhost:5173
```

Then:

1. Pick `example-formatted.xlsx`. Expect the Items table (D13:G20) to
   render with lavender fills on column D, dark text in the header,
   and thin horizontal rules between the body rows.
2. Pick `derm-summary-example.xlsx`. Expect all 6 sheets' first sheet
   ("Plaque Psoriasis - Benchmarking") to render without Q3 bleeding
   into P3. Resize the browser window narrower than the sheet; the
   outer container grows scrollbars and lets you pan the full data.
3. Right-click a cell — nothing happens.
4. Double-click a cell — nothing happens; cell does not become
   editable.

Also run:

```sh
npm run build          # tsc -b && vite build — should be clean
npm run lint           # should report 0 errors, 0 warnings
```

---

## 10. Directory map

```
fortune-sheet/
├── README.md                         (this file)
├── package.json
├── package-lock.json
├── .npmrc
├── vite.config.ts
├── tsconfig.json, tsconfig.app.json, tsconfig.node.json
├── eslint.config.mjs
├── prettier.config.mjs
├── index.html
├── index.tsx
├── public/                           (vite.svg, react.svg — unused placeholders)
└── src/
    ├── App.tsx                       (file picker + FortuneSheet mount)
    ├── styles/
    │   └── global.css                (Tailwind + phantom-scrollbar hide)
    ├── types/
    │   └── vite-env.d.ts             (Vite-injected env types)
    ├── components/
    │   └── FortuneSheet.tsx          (viewer component, reducer, useEffect pipeline)
    └── excel/
        ├── xlsxZip.ts                (JSZip + DOMParser)
        ├── colors.ts                 (ColorRef, resolveColor, applyTint)
        ├── theme.ts                  (parseTheme ← theme1.xml)
        ├── dxf.ts                    (parseDxfs ← styles.xml)
        ├── tableStyles.ts            (parseTableStyles ← styles.xml)
        ├── tables.ts                 (parseTable, indexSheetTables)
        ├── applyTableStyles.ts       (post-processor #1)
        ├── backfillInlineStringV.ts  (post-processor #2)
        ├── forceWrapOnOverflowCollisions.ts (post-processor #3)
        ├── bindDisplayBounds.ts      (post-processor #4)
        └── measureSheet.ts           (wrapper size calc)
```

---

## 11. Upgrade / escape-hatch notes

If the post-processor stack keeps growing with each new xlsx file,
**Univer** (https://github.com/dream-num/univer) is the most common
reference point — actively maintained, TypeScript-first, the
spiritual successor to luckysheet. Licensing caveat to be aware of
before adopting:

- Univer's **core framework is Apache-2.0** (true OSS).
- Univer's **xlsx import/export** (plus pivot tables, charts,
  printing, collaborative editing, history) ships in Univer's
  **non-OSS** edition. Their README describes it as "free for
  commercial use, with paid upgrade plans." It's source-available
  with a commercial tier rather than fully OSS.

For a read-only xlsx viewer, the xlsx reader is exactly the piece
we'd need, so switching to Univer is not a pure "swap for a more
permissive OSS lib" — it's adopting a source-available dependency.
Whether that's acceptable depends on your org's policy on non-OSI
licenses.

For the current scope (report viewers), the post-processor approach
is cheaper to maintain, fully open-source (Apache/MIT all the way
down), and easy to reason about. Revisit only if a real file shows up
that none of the post-processors can faithfully render.
