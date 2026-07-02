import ExcelJS from 'exceljs'
import type { Borders, Cell, CellStyle, Merge, Picture, Sheet } from '../model'
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
  const sheets = wb.worksheets.map((ws) => mapWorksheet(ws, wb))
  return { model: { sheets }, ms: performance.now() - t0 }
}

function mapWorksheet(ws: ExcelJS.Worksheet, wb: ExcelJS.Workbook): Sheet {
  let rowCount = ws.rowCount
  let colCount = ws.columnCount
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

  const merges = parseMerges(ws)
  const rawImages = ws.getImages()

  // Extend the used range so overlays that reach past the last content cell
  // (wide merges, floating pictures) have grid geometry to anchor against.
  for (const m of merges) {
    colCount = Math.max(colCount, m.c1 + 1)
    rowCount = Math.max(rowCount, m.r1 + 1)
  }
  for (const img of rawImages) {
    colCount = Math.max(colCount, Math.floor(img.range.br.nativeCol) + 1)
    rowCount = Math.max(rowCount, Math.floor(img.range.br.nativeRow) + 1)
  }

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
    merges,
    freeze: parseFreeze(ws),
    pictures: parsePictures(rawImages, wb, colWidths, rowHeights),
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

// Floating pictures via ExcelJS's native API (no manual unzip). Each image's bytes
// come from wb.getImage(); position/size are derived from the tl/br anchors against
// the (already extended) column/row pixel sizes — mirrors the ClosedXML backend path.
function parsePictures(
  rawImages: ReturnType<ExcelJS.Worksheet['getImages']>,
  wb: ExcelJS.Workbook,
  colWidths: number[],
  rowHeights: number[],
): Picture[] {
  const cum = (sizes: number[], upto: number, deflt: number) => {
    let acc = 0
    for (let i = 0; i < upto; i++) acc += sizes[i] ?? deflt
    return acc
  }
  const out: Picture[] = []
  for (const img of rawImages) {
    try {
      const media = wb.getImage(Number(img.imageId))
      if (!media) continue
      const b64 = media.base64 ?? uint8ToBase64(media.buffer)
      if (!b64) continue
      const ext = media.extension === 'jpg' ? 'jpeg' : (media.extension ?? 'png')
      const src = b64.startsWith('data:') ? b64 : `data:image/${ext};base64,${b64}`
      const { tl, br } = img.range
      // ExcelJS reports anchor offsets in EMU (914400 per inch → 9525 per px at 96 DPI).
      const px = (emu: number) => (emu || 0) / 9525
      const fromCol = Math.floor(tl.nativeCol)
      const fromRow = Math.floor(tl.nativeRow)
      const offsetX = px(tl.nativeColOff)
      const offsetY = px(tl.nativeRowOff)
      const left = cum(colWidths, fromCol, 64) + offsetX
      const top = cum(rowHeights, fromRow, 20) + offsetY
      const right = cum(colWidths, Math.floor(br.nativeCol), 64) + px(br.nativeColOff)
      const bottom = cum(rowHeights, Math.floor(br.nativeRow), 20) + px(br.nativeRowOff)
      out.push({
        src,
        fromCol,
        fromRow,
        offsetX,
        offsetY,
        width: Math.max(1, Math.round(right - left)),
        height: Math.max(1, Math.round(bottom - top)),
      })
    } catch {
      /* a single unreadable image must not break the preview */
    }
  }
  return out
}

// Browser-safe base64 for an image byte buffer (small POC images).
function uint8ToBase64(buf?: unknown): string {
  if (!buf) return ''
  const bytes =
    buf instanceof Uint8Array
      ? buf
      : ArrayBuffer.isView(buf)
        ? new Uint8Array((buf as ArrayBufferView).buffer)
        : buf instanceof ArrayBuffer
          ? new Uint8Array(buf)
          : null
  if (!bytes || !bytes.length) return ''
  let bin = ''
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i])
  return btoa(bin)
}
