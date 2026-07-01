import ExcelJS from 'exceljs'
import type { Borders, Cell, CellStyle, Merge, Sheet } from '../model'
import type { LoadResult } from './backendModel'

/**
 * Approach B — the backend serves the raw .xlsx bytes; ExcelJS parses them
 * in the browser and we map into the same shared WorkbookModel that the
 * SheetGrid renders. Nothing leaves the browser.
 */
export async function loadFromExcelJs(): Promise<LoadResult> {
  const t0 = performance.now()
  const res = await fetch('/api/workbook/file')
  if (!res.ok) throw new Error(`/api/workbook/file responded with ${res.status}`)
  const buf = await res.arrayBuffer()
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buf)
  const sheets = wb.worksheets.map(mapWorksheet)
  return { model: { sheets }, ms: performance.now() - t0 }
}

function mapWorksheet(ws: ExcelJS.Worksheet): Sheet {
  const rowCount = ws.rowCount
  const colCount = ws.columnCount
  const cells: Cell[] = []

  ws.eachRow({ includeEmpty: false }, (row, rn) => {
    row.eachCell({ includeEmpty: false }, (cell, cn) => {
      const r = rn - 1
      const c = cn - 1
      const text = (cell.text ?? '').toString()
      let raw: Cell['raw'] = null
      let formula: string | undefined

      if (cell.formula) {
        // Note: for shared-formula cells ExcelJS reports the master formula,
        // whereas the ClosedXML backend translates it per row — a nice contrast.
        formula = '=' + cell.formula
        const result = cell.result
        raw = typeof result === 'number' || typeof result === 'boolean' ? result : result != null ? String(result) : null
      } else {
        const v = cell.value
        if (typeof v === 'number' || typeof v === 'boolean' || typeof v === 'string') raw = v
        else if (v instanceof Date) raw = v.toISOString()
      }

      cells.push({ r, c, text, raw, formula, style: mapStyle(cell) })
    })
  })

  const colWidths: number[] = []
  for (let c = 1; c <= colCount; c++) {
    const w = ws.getColumn(c).width
    colWidths.push(Math.round((w ?? 8.43) * 7 + 5))
  }
  const rowHeights: number[] = []
  for (let r = 1; r <= rowCount; r++) {
    const h = ws.getRow(r).height
    rowHeights.push(Math.round((h ?? 15) * 4 / 3))
  }

  return {
    name: ws.name,
    rowCount,
    colCount,
    cells,
    colWidths,
    rowHeights,
    merges: parseMerges(ws),
    freeze: parseFreeze(ws),
  }
}

function colorHex(color?: Partial<ExcelJS.Color>): string | undefined {
  const argb = color?.argb
  if (!argb) return undefined // theme/indexed colors are skipped (mirrors the backend guard)
  const hex = argb.length === 8 ? argb.slice(2) : argb
  return ('#' + hex).toUpperCase()
}

function mapStyle(cell: ExcelJS.Cell): CellStyle | null {
  const style: CellStyle = {}
  const f = cell.font
  if (f) {
    if (f.bold) style.bold = true
    if (f.italic) style.italic = true
    if (f.underline) style.underline = true
    const col = colorHex(f.color)
    if (col && col !== '#000000') style.color = col
    // fontFamily/size are intentionally omitted unless clearly non-default,
    // to avoid attaching a style object to every one of thousands of cells.
    if (f.size && f.size !== 11) style.fontSize = f.size
  }

  const fill = cell.fill
  if (fill && fill.type === 'pattern' && fill.pattern === 'solid') {
    const bg = colorHex(fill.fgColor)
    if (bg && bg !== '#FFFFFF') style.bg = bg
  }

  const al = cell.alignment
  if (al) {
    if (al.horizontal === 'left' || al.horizontal === 'center' || al.horizontal === 'right') style.hAlign = al.horizontal
    if (al.vertical === 'top' || al.vertical === 'middle' || al.vertical === 'bottom') style.vAlign = al.vertical
    if (al.wrapText) style.wrap = true
  }

  const border = mapBorder(cell.border)
  if (border) style.border = border

  return Object.keys(style).length ? style : null
}

function mapBorder(bd?: Partial<ExcelJS.Borders>): Borders | null {
  if (!bd) return null
  const side = (s?: Partial<ExcelJS.Border>) => (s && s.style ? { style: s.style, color: colorHex(s.color) ?? null } : null)
  const top = side(bd.top)
  const right = side(bd.right)
  const bottom = side(bd.bottom)
  const left = side(bd.left)
  if (!top && !right && !bottom && !left) return null
  return { top, right, bottom, left }
}

// Read merges without touching ws.model (which would serialize the whole sheet).
function parseMerges(ws: ExcelJS.Worksheet): Merge[] {
  const out: Merge[] = []
  try {
    const container = (ws as unknown as { _merges?: Record<string, unknown> })._merges
    const map = (container as { merges?: Record<string, unknown> })?.merges ?? container ?? {}
    for (const key of Object.keys(map)) {
      const rng = (map as Record<string, { top?: number; left?: number; bottom?: number; right?: number; model?: { top: number; left: number; bottom: number; right: number } }>)[key]
      const top = rng.top ?? rng.model?.top
      const left = rng.left ?? rng.model?.left
      const bottom = rng.bottom ?? rng.model?.bottom
      const right = rng.right ?? rng.model?.right
      if ([top, left, bottom, right].every((n) => typeof n === 'number')) {
        out.push({ r0: (top as number) - 1, c0: (left as number) - 1, r1: (bottom as number) - 1, c1: (right as number) - 1 })
      }
    }
  } catch {
    /* best-effort */
  }
  return out
}

function parseFreeze(ws: ExcelJS.Worksheet): Sheet['freeze'] {
  try {
    const view = ws.views?.find((v) => v.state === 'frozen') as { xSplit?: number; ySplit?: number } | undefined
    if (view) return { rows: view.ySplit ?? 0, cols: view.xSplit ?? 0 }
  } catch {
    /* ignore */
  }
  return null
}
