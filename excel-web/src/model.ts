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
}

export interface WorkbookModel {
  sheets: Sheet[]
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
