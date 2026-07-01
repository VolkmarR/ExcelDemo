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
- `DemoData.xlsx` — the sample workbook (1 sheet `Tabelle1`, 4001×4 with cached `A+B+C`
  formulas, plus a small `Diagram` sheet that now carries an embedded pie chart)

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
charts; the one engine with charts built in (Univer) puts them behind a paid license. Actually
*drawing* the chart would be easy — the data is cached in the file and a small MIT renderer such
as [Recharts](https://recharts.org) would draw the pie in a few lines — but every path first
requires **bypassing the parsing library or paying for it**, so it is intentionally out of scope
for this POC.

---

## Known limitations

- **Charts / diagrams are not rendered** on any tab. Chart *reading* is unsupported by ClosedXML
  and ExcelJS, and is a paid (`@univerjs-pro`) feature in Univer — see
  [Charts / diagrams](#charts--diagrams) for the per-option reasons.
- **`pnpm build` (production) fails on the Univer chunk.** Vite 8 ships an experimental
  Rolldown/oxc bundler whose parser overflows (`WebAssembly.Memory.grow`) on Univer's
  ~10 MB bundle. The Backend and ExcelJS parts build fine; only the isolated lazy Univer
  chunk trips it. **`pnpm dev` runs all three tabs correctly** — that is how this POC is
  demoed. Fixing the production build means splitting Univer via `manualChunks`, loading
  its UMD build, or pinning stable Vite 7 (esbuild optimizer).
- **Univer + Vite 8.** Univer needs the `optimizeDeps` config and the 30 extra
  `devDependencies` in `excel-web/` (all documented in `vite.config.ts`) to load in dev.
  Options A and B need none of that.
- **Licenses:** ClosedXML, ExcelJS, `@tanstack/react-virtual` are MIT; Univer is
  Apache-2.0. All permissive.
