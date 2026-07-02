// Shared workbook model consumed by all three approaches.
// The backend (ClosedXML) serializes to exactly this shape (camelCase); the
// ExcelJS adapter maps into it; the Univer adapter maps out of it.

export type HAlign = 'left' | 'center' | 'right'
export type VAlign = 'top' | 'middle' | 'bottom'

export interface BorderSide {
  /** Excel border-style token, e.g. "thin" | "medium" | "thick" | "double" | "dashed" | "dotted". */
  style: string
  color?: string | null
}

export interface Borders {
  top?: BorderSide | null
  right?: BorderSide | null
  bottom?: BorderSide | null
  left?: BorderSide | null
}

export interface CellStyle {
  bold?: boolean | null
  italic?: boolean | null
  underline?: boolean | null
  color?: string | null
  bg?: string | null
  fontFamily?: string | null
  fontSize?: number | null
  hAlign?: HAlign | null
  vAlign?: VAlign | null
  wrap?: boolean | null
  border?: Borders | null
}

export interface Cell {
  /** 0-based row index. */
  r: number
  /** 0-based column index. */
  c: number
  /** Display string (number format already applied). */
  text: string
  /** Typed value; used for right-aligning numbers. */
  raw?: string | number | boolean | null
  /** Formula including the leading "=", when the cell is a formula. */
  formula?: string | null
  style?: CellStyle | null
}

export interface Merge {
  r0: number
  c0: number
  r1: number
  c1: number
}

export interface Picture {
  /** Data URL (base64), e.g. "data:image/jpeg;base64,…". */
  src: string
  /** 0-based anchor cell (top-left). */
  fromCol: number
  fromRow: number
  /** Pixel offset within the anchor cell. */
  offsetX: number
  offsetY: number
  /** Display size in px. */
  width: number
  height: number
}

/** A rectangular cell range, 0-based inclusive (same shape as Merge). */
export interface Range {
  r0: number
  c0: number
  r1: number
  c1: number
}

/** An Excel Table (ListObject) region. Colors come from the workbook theme + the built-in style. */
export interface TableModel {
  name: string
  range: Range
  /** Built-in table-style name, e.g. "TableStyleMedium4" | "TableStyleLight9". */
  styleName?: string | null
  showHeaderRow: boolean
  showTotalsRow: boolean
  showRowStripes: boolean
  showColumnStripes: boolean
  showFirstColumn: boolean
  showLastColumn: boolean
}

/** Where a color-scale stop is anchored. min/max/percentile are computed from the data. */
export type ColorScaleStopKind = 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula'

export interface ColorScaleStop {
  kind: ColorScaleStopKind
  /** Literal value for kind 'num'/'percent'/'percentile'; ignored for 'min'/'max'. */
  value?: number | null
  /** Stop color as #RRGGBB. */
  color: string
}

/** A 2- or 3-color color scale. The only conditional-format type implemented so far. */
export interface ColorScaleFormat {
  type: 'colorScale'
  range: Range
  stops: ColorScaleStop[]
}

/** Discriminated union so cellIs/dataBar/iconSet can be added later without reshaping. */
export type ConditionalFormat = ColorScaleFormat

/** Resolved workbook theme palette (hex). Needed to render exact table-style colors. */
export interface ThemePalette {
  accent1: string
  accent2: string
  accent3: string
  accent4: string
  accent5: string
  accent6: string
  dk1: string
  lt1: string
  dk2: string
  lt2: string
}

export interface Sheet {
  name: string
  rowCount: number
  colCount: number
  cells: Cell[]
  /** Per-column width in px (length === colCount). */
  colWidths: number[]
  /** Per-row height in px (length === rowCount). */
  rowHeights: number[]
  merges: Merge[]
  freeze?: { rows: number; cols: number } | null
  pictures?: Picture[] | null
  tables?: TableModel[] | null
  conditionalFormats?: ConditionalFormat[] | null
}

export interface WorkbookModel {
  sheets: Sheet[]
  /** Workbook theme palette; used to resolve exact table-style colors. */
  theme?: ThemePalette | null
}

/** Column index -> spreadsheet column letters (0 -> "A", 26 -> "AA"). */
export function columnLabel(index: number): string {
  let n = index
  let label = ''
  do {
    label = String.fromCharCode(65 + (n % 26)) + label
    n = Math.floor(n / 26) - 1
  } while (n >= 0)
  return label
}

/** (row, col) -> A1 address, e.g. (1, 3) -> "D2". */
export function cellAddress(r: number, c: number): string {
  return columnLabel(c) + (r + 1)
}
