// Computes the *visual effect* of Excel Tables and conditional formatting for the
// shared SheetGrid (Backend + ExcelJS tabs). Both adapters only extract structured
// metadata (table style name, ranges, color-scale stops, theme palette); the exact
// colors are resolved here, once, so the two grid tabs render identically.
//
// Exact-color path: table styles reference the workbook theme's accent palette plus a
// built-in per-family treatment; theme tints use Excel's HSL luminance transform.
//
// Precedence (matches Excel): conditional-format fill > explicit cell style > table style.

import type {
  Borders,
  BorderSide,
  Cell,
  CellStyle,
  ColorScaleFormat,
  Sheet,
  ThemePalette,
} from '../model'

const clamp = (v: number, lo: number, hi: number) => (v < lo ? lo : v > hi ? hi : v)

// ---- color math ---------------------------------------------------------------

function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace('#', '')
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]
}

function rgbToHex(r: number, g: number, b: number): string {
  const to2 = (n: number) => Math.round(clamp(n, 0, 255)).toString(16).padStart(2, '0')
  return ('#' + to2(r) + to2(g) + to2(b)).toUpperCase()
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  r /= 255
  g /= 255
  b /= 255
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = (max + min) / 2
  if (max === min) return [0, 0, l] // achromatic
  const d = max - min
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
  let h = 0
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0)
  else if (max === g) h = (b - r) / d + 2
  else h = (r - g) / d + 4
  return [h / 6, s, l]
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) return [l * 255, l * 255, l * 255] // achromatic
  const hue2rgb = (p: number, q: number, t: number) => {
    if (t < 0) t += 1
    if (t > 1) t -= 1
    if (t < 1 / 6) return p + (q - p) * 6 * t
    if (t < 1 / 2) return q
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6
    return p
  }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s
  const p = 2 * l - q
  return [hue2rgb(p, q, h + 1 / 3) * 255, hue2rgb(p, q, h) * 255, hue2rgb(p, q, h - 1 / 3) * 255]
}

/**
 * Excel's theme-tint transform (the `tint` attribute on a themed color): modulate the
 * HSL luminance. tint<0 darkens, tint>0 lightens. This matches how Excel derives the
 * lighter/darker variants used by built-in table styles.
 */
function applyTint(hex: string, tint: number): string {
  if (!tint) return hex
  const [r, g, b] = hexToRgb(hex)
  const [h, s, l] = rgbToHsl(r, g, b)
  const nl = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint
  const [nr, ng, nb] = hslToRgb(h, s, clamp(nl, 0, 1))
  return rgbToHex(nr, ng, nb)
}

function lerpHex(a: string, b: string, t: number): string {
  const [ar, ag, ab] = hexToRgb(a)
  const [br, bg, bb] = hexToRgb(b)
  return rgbToHex(ar + (br - ar) * t, ag + (bg - ag) * t, ab + (bb - ab) * t)
}

// Standard Excel theme-tint constants (lighter N%), for reference / built-in table styles.
const LIGHTER_40 = 0.3999755859375
const LIGHTER_80 = 0.7999816888943144

// ---- theme resolution ---------------------------------------------------------

// Fallback = the classic Office palette, used only if the workbook carries no theme.
const DEFAULT_THEME: ThemePalette = {
  accent1: '#4472C4',
  accent2: '#ED7D31',
  accent3: '#A5A5A5',
  accent4: '#FFC000',
  accent5: '#5B9BD5',
  accent6: '#70AD47',
  dk1: '#000000',
  lt1: '#FFFFFF',
  dk2: '#44546A',
  lt2: '#E7E6E6',
}

interface ResolvedTheme {
  accent(i: number): string // 1..6
  dk1: string
  lt1: string
}

function resolveTheme(theme?: ThemePalette | null): ResolvedTheme {
  const t = theme ?? DEFAULT_THEME
  const accents = [t.accent1, t.accent2, t.accent3, t.accent4, t.accent5, t.accent6]
  return {
    accent: (i) => accents[clamp(i, 1, 6) - 1] ?? DEFAULT_THEME.accent1,
    dk1: t.dk1 || '#000000',
    lt1: t.lt1 || '#FFFFFF',
  }
}

// ---- built-in table styles -----------------------------------------------------

interface TableStyleDef {
  headerBg?: string
  headerColor?: string
  headerBold: boolean
  headerBottomBorder?: BorderSide
  band1Bg?: string
  cellBorder?: BorderSide // whole-table thin border (Light family)
}

/** Parse "TableStyleMedium4" → family + accent index (0 = neutral/no-accent variant). */
function parseTableStyle(name?: string | null): { family: 'light' | 'medium' | 'dark' | 'none'; accent: number } {
  const m = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(name ?? '')
  if (!m) return { family: 'none', accent: 0 }
  const family = m[1].toLowerCase() as 'light' | 'medium' | 'dark'
  const n = parseInt(m[2], 10)
  if (family === 'medium') return { family, accent: n === 1 ? 0 : ((n - 2) % 6) + 1 } // Medium2..7 → accent1..6
  if (family === 'light') {
    const pos = (n - 1) % 7 // gallery rows of 7: col 1 = neutral, cols 2..7 = accent1..6
    return { family, accent: pos === 0 ? 0 : pos }
  }
  return { family, accent: n <= 2 ? 0 : ((n - 3) % 6) + 1 }
}

function tableStyleDef(name: string | null | undefined, rt: ResolvedTheme): TableStyleDef {
  const { family, accent } = parseTableStyle(name)
  const accentHex = accent === 0 ? '#808080' : rt.accent(accent)
  if (family === 'medium') {
    // Solid accent header, white bold text, light banded rows, no cell borders.
    return { headerBg: accentHex, headerColor: rt.lt1, headerBold: true, band1Bg: applyTint(accentHex, LIGHTER_80) }
  }
  if (family === 'dark') {
    return { headerBg: applyTint(accentHex, -0.25), headerColor: rt.lt1, headerBold: true, band1Bg: applyTint(accentHex, LIGHTER_80) }
  }
  // Light (and neutral): no header fill; bold header with an accent bottom border;
  // thin accent-tinted borders around the cells; light banded rows.
  return {
    headerBold: true,
    headerBottomBorder: { style: 'medium', color: accentHex },
    band1Bg: applyTint(accentHex, LIGHTER_80),
    cellBorder: { style: 'thin', color: applyTint(accentHex, LIGHTER_40) },
  }
}

// ---- color scale ---------------------------------------------------------------

interface ScaleAnchor {
  value: number
  color: string
}
interface PreparedScale {
  r0: number
  c0: number
  r1: number
  c1: number
  anchors: ScaleAnchor[] // sorted ascending by value; length 2 or 3
}

function percentile(sorted: number[], p: number): number {
  if (sorted.length === 0) return 0
  const idx = clamp(p, 0, 100) / 100 * (sorted.length - 1)
  const lo = Math.floor(idx)
  const hi = Math.ceil(idx)
  if (lo === hi) return sorted[lo]
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo)
}

function prepareScale(cf: ColorScaleFormat, sheet: Sheet): PreparedScale | null {
  const { r0, c0, r1, c1 } = cf.range
  const values: number[] = []
  for (const cell of sheet.cells) {
    if (cell.r < r0 || cell.r > r1 || cell.c < c0 || cell.c > c1) continue
    if (typeof cell.raw === 'number') values.push(cell.raw)
  }
  if (values.length === 0 || cf.stops.length < 2) return null
  let min = values[0]
  let max = values[0]
  for (const v of values) {
    if (v < min) min = v
    if (v > max) max = v
  }
  const sorted = [...values].sort((a, b) => a - b)
  const anchors = cf.stops.map((s) => {
    let value: number
    switch (s.kind) {
      case 'min': value = min; break
      case 'max': value = max; break
      case 'num': value = s.value ?? min; break
      case 'percent': value = min + (max - min) * ((s.value ?? 0) / 100); break
      case 'percentile': value = percentile(sorted, s.value ?? 50); break
      default: value = min // 'formula' can't be evaluated here
    }
    return { value, color: s.color.toUpperCase() }
  })
  anchors.sort((a, b) => a.value - b.value)
  return { r0, c0, r1, c1, anchors }
}

function scaleColor(v: number, anchors: ScaleAnchor[]): string {
  if (v <= anchors[0].value) return anchors[0].color
  const last = anchors[anchors.length - 1]
  if (v >= last.value) return last.color
  for (let i = 0; i < anchors.length - 1; i++) {
    const a = anchors[i]
    const b = anchors[i + 1]
    if (v <= b.value) {
      const t = b.value === a.value ? 0 : (v - a.value) / (b.value - a.value)
      return lerpHex(a.color, b.color, t)
    }
  }
  return last.color
}

// ---- table region index --------------------------------------------------------

interface PreparedTable {
  r0: number
  c0: number
  r1: number
  c1: number
  hasHeader: boolean
  banded: boolean
  dataTop: number
  dataBottom: number
  def: TableStyleDef
}

function prepareTable(sheet: Sheet, rt: ResolvedTheme): PreparedTable[] {
  return (sheet.tables ?? []).map((t) => {
    const { r0, c0, r1, c1 } = t.range
    return {
      r0,
      c0,
      r1,
      c1,
      hasHeader: t.showHeaderRow,
      banded: t.showRowStripes,
      dataTop: r0 + (t.showHeaderRow ? 1 : 0),
      dataBottom: r1 - (t.showTotalsRow ? 1 : 0),
      def: tableStyleDef(t.styleName, rt),
    }
  })
}

// Excel shades the first data row of a striped table, then alternates.
const FIRST_DATA_ROW_BANDED = true

// ---- border merge helpers (explicit cell borders always win) -------------------

function withSide(border: Borders | null | undefined, side: keyof Borders, value: BorderSide): Borders {
  const b: Borders = { ...(border ?? {}) }
  if (b[side] == null) b[side] = value
  return b
}

function withAllSides(border: Borders | null | undefined, value: BorderSide): Borders {
  const b: Borders = { ...(border ?? {}) }
  for (const side of ['top', 'right', 'bottom', 'left'] as (keyof Borders)[]) {
    if (b[side] == null) b[side] = value
  }
  return b
}

// ---- public API -----------------------------------------------------------------

export interface DecoratedCell {
  style: CellStyle | null
  isTableHeader: boolean
  /** True when the cell lies inside a table region (used to suppress the sheet zebra). */
  inTable: boolean
}

export interface SheetStyling {
  decorate(r: number, c: number, cell: Cell | undefined): DecoratedCell
}

const PLAIN: DecoratedCell = { style: null, isTableHeader: false, inTable: false }

/**
 * Build a per-sheet styling context once (memoize on [sheet, theme]). `decorate` is
 * cheap per cell — it scans the sheet's handful of tables / color-scale rules.
 */
export function buildSheetStyling(sheet: Sheet, theme?: ThemePalette | null): SheetStyling {
  const rt = resolveTheme(theme)
  const tables = prepareTable(sheet, rt)
  const scales: PreparedScale[] = []
  for (const cf of sheet.conditionalFormats ?? []) {
    if (cf.type !== 'colorScale') continue
    const prepared = prepareScale(cf, sheet)
    if (prepared) scales.push(prepared)
  }

  if (tables.length === 0 && scales.length === 0) {
    return { decorate: (_r, _c, cell) => (cell?.style ? { style: cell.style, isTableHeader: false, inTable: false } : PLAIN) }
  }

  const colorScaleBgAt = (r: number, c: number, cell: Cell | undefined): string | undefined => {
    if (!cell || typeof cell.raw !== 'number') return undefined
    for (const s of scales) {
      if (r >= s.r0 && r <= s.r1 && c >= s.c0 && c <= s.c1) return scaleColor(cell.raw, s.anchors)
    }
    return undefined
  }

  return {
    decorate(r, c, cell) {
      const table = tables.find((t) => r >= t.r0 && r <= t.r1 && c >= t.c0 && c <= t.c1)
      const csBg = colorScaleBgAt(r, c, cell)
      if (!table && !csBg) return cell?.style ? { style: cell.style, isTableHeader: false, inTable: false } : PLAIN

      const base = cell?.style
      const out: CellStyle = { ...(base ?? {}) }
      let isTableHeader = false

      if (table) {
        const { def } = table
        if (table.hasHeader && r === table.r0) {
          isTableHeader = true
          if (out.bg == null && def.headerBg) out.bg = def.headerBg
          if (out.color == null && def.headerColor) out.color = def.headerColor
          if (out.bold == null && def.headerBold) out.bold = true
          if (def.headerBottomBorder) out.border = withSide(out.border, 'bottom', def.headerBottomBorder)
          if (def.cellBorder) out.border = withAllSides(out.border, def.cellBorder)
        } else if (r >= table.dataTop && r <= table.dataBottom) {
          const dataRow = r - table.dataTop
          const banded = table.banded && (dataRow % 2 === 0) === FIRST_DATA_ROW_BANDED
          if (out.bg == null && banded && def.band1Bg) out.bg = def.band1Bg
          if (def.cellBorder) out.border = withAllSides(out.border, def.cellBorder)
        }
      }

      if (csBg) out.bg = csBg // conditional-format fill wins over cell + table

      return { style: out, isTableHeader, inTable: !!table }
    },
  }
}
