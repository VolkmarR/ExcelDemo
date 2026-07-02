# Excel Preview POC

A proof-of-concept that renders a **high-fidelity, read-only preview** of an Excel file
(`DemoData.xlsx`) in the browser. The preview is a table you navigate with the **arrow
keys or the mouse**, with a **focused cell** and a formula bar, just like Excel in
read-only mode.

The same workbook is rendered **three different ways** behind a tab switcher so the
approaches can be compared on fidelity, performance, and — the focus of this document —
**how many dependencies each one adds and how big they are**.

The Excel file is never sent to any external service: the .NET backend serves it
same-origin and all parsing happens on the backend or in the browser.

- `ExcelApi/` — ASP.NET Core (.NET 10) backend
- `excel-web/` — React 19 + Vite 8 frontend
- `DemoData.xlsx` — the sample workbook: sheet `Tabelle1` (4001×4 with cached `A+B+C`
  formulas, formatted as an Excel **table** with a **color-scale** on column D, two
  merged-cell regions, and an embedded photo) plus a small `Diagram` sheet (a second table
  and an embedded pie chart)

---

## Dependencies added per option

Every option reuses the pre-existing stack (ASP.NET minimal API, React, Vite). The table
below lists only what each **approach added on top of that**.

| Option | Library added | Direct deps | Transitive packages | Installed size | Shipped to the browser |
|--------|---------------|:-----------:|:-------------------:|---------------:|-----------------------:|
| **A — Backend · ClosedXML** | `ClosedXML` (NuGet, MIT) | **1** | 7 | **~9.4 MB** of DLLs (server-side) | **0 B** — the browser only receives compact JSON |
| **B — Frontend · ExcelJS** | `exceljs` (npm, MIT) | **1** | a few | 23 MB in `node_modules` (dev only) | **~0.93 MB** (`exceljs.min.js`, self-contained) |
| **C — Univer SDK** | `@univerjs/presets` + `@univerjs/preset-sheets-core` (npm, Apache-2.0) | **2** (+ 30 leaf deps, see note) | **~100** `@univerjs/*` packages | **~177 MB** in `node_modules` | **~9.3 MB** of JS (the Univer engine) |
| _shared by A & B_ | `@tanstack/react-virtual` (npm, MIT) | 1 | 0 | 64 KB | a few KB |

### Notes on the numbers

- **Measurement method.** "Installed size" is the on-disk size of the package(s) in the
  pnpm store / NuGet output. "Shipped to the browser" is the minified library code the
  browser actually downloads (`exceljs.min.js` for B; the bundled Univer engine for C).
  Option A ships **no library code** — it sends a compact JSON model, so the browser cost
  is just `@tanstack/react-virtual` (~few KB) shared with B.
- **Option A — ClosedXML (7 transitive):** `ClosedXML.Parser`, `DocumentFormat.OpenXml`,
  `DocumentFormat.OpenXml.Framework`, `ExcelNumberFormat`, `RBush.Signed`,
  `SixLabors.Fonts`, `System.IO.Packaging`. Bulk of the ~9.4 MB is
  `DocumentFormat.OpenXml.dll` (6.1 MB) + `ClosedXML.dll` (1.7 MB).
- **Option B — ExcelJS:** the 23 MB in `node_modules` is dev-only (source + multiple
  builds); at runtime the browser loads the self-contained **928 KB** minified bundle.
- **Option C — Univer:** the two preset packages pull in **73 `@univerjs/*` + 27
  `@univerjs-pro/*` = 100 packages (~164 MB)**. On top of that, **30 leaf packages**
  (Radix UI, `prop-types`, `opentype.js`, `sonner`, `react-transition-group`, `clsx`, …)
  were added as direct `devDependencies` purely to work around Vite 8's dependency
  optimizer (see the "Univer + Vite 8" note below). By far the heaviest option.

### Rough comparison

```
Browser payload:   A  ~0 (JSON only) │ B  ~0.93 MB │ C  ~9.3 MB
Packages added:    A  1 (+7 trans.)  │ B  1        │ C  2 (+30 leaves, ~100 transitive)
Install footprint: A  ~9 MB (DLLs)   │ B  ~23 MB   │ C  ~177 MB
```

**Takeaway:** A and B give the same custom grid (identical UX) at a tiny footprint —
A keeps everything server-side and ships only JSON; B keeps everything in the browser for
under 1 MB. C offers the richest, most "spreadsheet-like" engine out of the box but is
one to two orders of magnitude heavier and needs extra build configuration.

---

## Running it

```bash
# 1) Backend  (from ExcelApi/)
dotnet run --launch-profile http          # → http://localhost:5269

# 2) Frontend (from excel-web/)
pnpm install
pnpm dev                                  # → http://localhost:5173
```

Open **http://localhost:5173** and switch between the three tabs. Vite proxies `/api` to
the backend, so no CORS setup is needed.

The three tabs:

| Tab | Parses | Renders with |
|-----|--------|--------------|
| **Backend · ClosedXML** | `.NET` server (`GET /api/workbook` → JSON) | shared custom virtualized grid (`src/grid/SheetGrid.tsx`) |
| **Frontend · ExcelJS** | the browser (`GET /api/workbook/file` → ExcelJS) | the same shared grid |
| **Univer SDK** | (same model) → Univer snapshot | Univer canvas engine, read-only |

---

## Pictures & merged cells

`Tabelle1` includes two **merged-cell regions** (`F3:H3`, `F6:F8`) and an **embedded photo**
(anchored over columns J–S). Both render on **all three tabs**:

| Option | How pictures are read | How they're drawn |
|--------|-----------------------|-------------------|
| **Backend · ClosedXML** | `IXLWorksheet.Pictures` (native API — bytes, format, pixel anchor), emitted as a base64 data URL in the JSON model | `<img>` overlay in the shared grid |
| **Frontend · ExcelJS** | `ws.getImages()` + `wb.getImage()` (native API — buffer + anchor), no manual unzip | the same `<img>` overlay |
| **Univer SDK** | (same base64 model from the backend) | `fWorksheet.newOverGridImage()…insertImages()` via the OSS `@univerjs/preset-sheets-drawing` |

Unlike charts, **pictures need no library workaround**: ClosedXML and ExcelJS both expose images
through their normal read APIs, and Univer's image support is plain **Apache-2.0**
(`@univerjs/preset-sheets-drawing`), not the commercial `@univerjs-pro` tier. The photo's embedded
external hyperlink (a source-credit URL) is **ignored** — only the bytes already inside the file
are rendered, so nothing leaves localhost.

**Merged cells** were already supported by all three renderers; the fix was that the
Backend/ExcelJS **used range is now grown to cover overlays** (merges or pictures) that extend past
the last cell with content — otherwise a wide merge like `F3:H3` (reaching column H, past the
data's last column F) was clipped, and the picture had no columns to anchor to.

Dependencies added: **backend 0** (ClosedXML's picture API is built in), **Backend/ExcelJS rendering
0** (a plain `<img>`), and **`@univerjs/preset-sheets-drawing`** (Apache-2.0, largely already present
transitively) for the Univer tab.

---

## Tables & conditional formatting

`DemoData.xlsx` formats its data as **Excel Tables** and adds a conditional-format **color scale**:

- **`Tabelle1`** (sheet *Tabelle1*, `A1:D4001`) — style **`TableStyleMedium4`** (solid green
  header, banded rows). Column **D** also carries a 2-color **color scale**: green `#63BE7B`
  for the smallest values → pale yellow `#FFEF9C` for the largest.
- **`Tabelle2`** (sheet *Diagram*, `A1:B4`) — style **`TableStyleLight9`** (teal header
  underline, thin borders, banded rows).

Both render on the **Backend · ClosedXML** and **Frontend · ExcelJS** tabs, with the **exact
theme colors** Excel uses:

| Option | How tables / color scale are read | How they're drawn |
|--------|-----------------------------------|-------------------|
| **Backend · ClosedXML** | `IXLWorksheet.Tables` (name, range, style name, header/stripe flags) + `IXLWorksheet.ConditionalFormats` (color-scale stop colors); the workbook palette from `wb.Theme` | shared grid (`SheetGrid.tsx`) |
| **Frontend · ExcelJS** | `ws.tables` / `ws.getTables()` + `ws.conditionalFormattings` (native read APIs); the palette parsed from the theme XML ExcelJS already holds — no re-unzip | the same shared grid |
| **Univer SDK** | — not rendered (see below) | — |

Neither library resolves the *colors* of a table style — Excel derives those live from the
built-in style plus the workbook theme. So both adapters extract only **structured metadata**,
and a single shared helper (`excel-web/src/grid/sheetStyling.ts`) resolves the exact colors
once: it maps the style name to a theme accent (e.g. `TableStyleMedium4` → accent 3 = `#196B24`;
`TableStyleLight9` → accent 1 = `#156082`), applies Excel's HSL **theme-tint** transform for the
banded rows, and computes the color scale by interpolating between the stop colors across the
column's actual min/max. The two grid tabs therefore render **identically**. Precedence matches
Excel: **conditional-format fill > explicit cell formatting > table style**.

**Univer** is intentionally left minimal here: its real table and conditional-formatting
features are in the commercial `@univerjs-pro/*` packages (the same license line that keeps
charts out — see below), so its tab shows the data without table chrome or the color scale.

Dependencies added: **0** on all paths — ClosedXML and ExcelJS expose tables/CF/theme through
their normal read APIs, and the color resolution is a small dependency-free TypeScript module.

---

## Charts / diagrams

The `Diagram` sheet carries an embedded **pie chart** — categories Italy / Germany /
Austria = 200,000 / 100,000 / 50,000, titled "Summary". Its labels and values are
**cached inside the chart part** (`xl/charts/chart1.xml`, in `<c:strCache>` / `<c:numCache>`),
so reading the chart *data* would be trivial — much like the cached formula results the
grid already relies on.

Even so, **none of the three tabs renders the chart**, and the reason differs per option —
which makes it a useful comparison in its own right:

| Option | Chart? | Why — and what it would take (not done) |
|--------|:------:|-----------------------------------------|
| **Backend · ClosedXML** | ❌ | ClosedXML 0.105.0 has **no chart-read API** (it reads images, not charts) and does not expose its internal Open XML document. Reading the chart would need a **separate parse pass** with the Open XML SDK (`DocumentFormat.OpenXml`, already a transitive dependency) walking `WorksheetPart → DrawingsPart → ChartPart` — a second parser disjoint from the ClosedXML one. Deliberately not added. |
| **Frontend · ExcelJS** | ❌ | ExcelJS 4.4.0 **cannot read charts** ([issue #1569](https://github.com/exceljs/exceljs/issues/1569)). The only route is to **unzip the `.xlsx` in the browser and hand-parse the OOXML chart XML** (e.g. `fflate` + `DOMParser`), bypassing ExcelJS. Deliberately not built. |
| **Univer SDK** | ❌ | Univer *can* draw charts, but only via the **commercial `@univerjs-pro/*`** packages (`engine-chart`, `sheets-chart`, `sheets-chart-ui`) gated behind `@univerjs-pro/license` — unlicensed use is watermarked/restricted. Outside this POC's MIT/Apache-only rule. |

**The common thread:** across the open-source spreadsheet ecosystem, chart **reading** is the
consistently missing piece. ClosedXML and ExcelJS both parse cells, styles and merges but not
charts (they *do* handle images — see above); the one engine with charts built in (Univer) puts
them behind a paid license. Actually
*drawing* the chart would be easy — the data is cached in the file and a small MIT renderer such
as [Recharts](https://recharts.org) would draw the pie in a few lines — but every path first
requires **bypassing the parsing library or paying for it**, so it is intentionally out of scope
for this POC.

---

## Known limitations

- **Charts / diagrams are not rendered** on any tab. Chart *reading* is unsupported by ClosedXML
  and ExcelJS, and is a paid (`@univerjs-pro`) feature in Univer — see
  [Charts / diagrams](#charts--diagrams) for the per-option reasons.
- **Tables & conditional formatting render on the Backend and ExcelJS tabs only.** Univer's
  table/CF support is commercial (`@univerjs-pro`), so its tab shows the data without table
  styling or the color scale — see
  [Tables & conditional formatting](#tables--conditional-formatting). Only 2-color/3-color
  **color scales** are implemented (the file's only CF type); data bars, icon sets and
  cell-value rules are modelled as a discriminated union but not yet drawn.
- **`pnpm build` (production) fails on the Univer chunk.** Vite 8 ships an experimental
  Rolldown/oxc bundler whose parser overflows (`WebAssembly.Memory.grow`) on Univer's
  ~10 MB bundle. The Backend and ExcelJS parts build fine; only the isolated lazy Univer
  chunk trips it. **`pnpm dev` runs all three tabs correctly** — that is how this POC is
  demoed. Fixing the production build means splitting Univer via `manualChunks`, loading
  its UMD build, or pinning stable Vite 7 (esbuild optimizer).
- **Univer + Vite 8.** Univer needs the `optimizeDeps` config and the 30 extra
  `devDependencies` in `excel-web/` (all documented in `vite.config.ts`) to load in dev; the
  drawing packages (`@univerjs/preset-sheets-drawing` and its plugins) join that same
  `optimizeDeps.exclude` list for image support. Options A and B need none of that.
- **Picture sizing on the ExcelJS tab is approximate.** ExcelJS exposes an image's cell anchors
  rather than its stored pixel size, so the width is reconstructed from column widths and can
  differ slightly from the Backend tab (which reads the exact size from ClosedXML); position and
  height match.
- **Licenses:** ClosedXML, ExcelJS, `@tanstack/react-virtual` are MIT; Univer is
  Apache-2.0. All permissive.
