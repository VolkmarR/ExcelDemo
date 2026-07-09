import ExcelJS from 'exceljs'
import type {
  Borders,
  Cell,
  CellStyle,
  ColorScaleStop,
  ConditionalFormat,
  Merge,
  Picture,
  Range,
  Sheet,
  TableModel,
  ThemePalette,
} from '../model'
import type { LoadResult } from './backendModel'

/**
 * Approach B — the backend serves the raw .xlsx bytes; ExcelJS parses them
 * in the browser and we map into the same shared WorkbookModel that the
 * SheetGrid renders. Nothing leaves the browser.
 */
export async function loadFromExcelJs(
    selectedFile: File
): Promise<LoadResult> {
  const t0 = performance.now()

  const buf = await selectedFile.arrayBuffer()

  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buf)

  const sheets = wb.worksheets.map((ws) => mapWorksheet(ws, wb))

  return {
    model: {
      sheets,
      theme: parseTheme(wb),
    },
    ms: performance.now() - t0,
  }
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
    tables: parseTables(ws),
    conditionalFormats: parseConditionalFormats(ws),
  }
}

// Parse an A1 range like "A1:D4001" (or a single cell "B2") into 0-based bounds.
function parseA1Range(ref: string): Range | null {
  const m = /^\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$/.exec(ref.trim())
  if (!m) return null
  const col = (s: string) => {
    let n = 0
    for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
    return n - 1
  }
  const c0 = col(m[1])
  const r0 = parseInt(m[2], 10) - 1
  const c1 = m[3] ? col(m[3]) : c0
  const r1 = m[4] ? parseInt(m[4], 10) - 1 : r0
  return { r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) }
}

// Excel Tables via ExcelJS's parsed `ws.tables` map (each value wraps a raw table model
// with tableRef/style/headerRow). Not in the public typings, so read defensively.
function parseTables(ws: ExcelJS.Worksheet): TableModel[] {
  const out: TableModel[] = []
  try {
    const tablesObj = (ws as unknown as { tables?: Record<string, unknown> }).tables
    for (const entry of Object.values(tablesObj ?? {})) {
      try {
        const t = entry as { model?: Record<string, unknown>; name?: string; style?: Record<string, unknown> }
        const model = (t.model ?? t) as Record<string, unknown>
        const ref = (model.tableRef ?? model.ref ?? '') as string
        const range = parseA1Range(ref)
        if (!range) continue
        const style = (model.style ?? t.style ?? {}) as Record<string, unknown>
        out.push({
          name: (model.name ?? t.name ?? 'Table') as string,
          range,
          styleName: (style.theme as string) ?? null,
          // ExcelJS reports headerRow=false even when a header is shown (OOXML headerRowCount
          // defaults to 1 when absent). Excel tables show a header by default, so assume true.
          showHeaderRow: true,
          showTotalsRow: !!model.totalsRow,
          showRowStripes: !!style.showRowStripes,
          showColumnStripes: !!style.showColumnStripes,
          showFirstColumn: !!style.showFirstColumn,
          showLastColumn: !!style.showLastColumn,
        })
      } catch {
        /* skip a single unreadable table */
      }
    }
  } catch {
    /* best-effort */
  }
  return out
}

// Conditional formatting: ExcelJS parses it into `ws.conditionalFormattings`
// (array of { ref, rules }) on load — merging the x14 extension variant. Only color
// scales are mapped; the data min/max is computed later in the shared styling helper.
function parseConditionalFormats(ws: ExcelJS.Worksheet): ConditionalFormat[] {
  const out: ConditionalFormat[] = []
  try {
    const cfs = (ws as unknown as { conditionalFormattings?: Array<{ ref?: string; rules?: unknown[] }> }).conditionalFormattings
    if (!Array.isArray(cfs)) return out
    for (const cf of cfs) {
      const refs = (cf.ref ?? '').split(/\s+/).filter(Boolean)
      for (const raw of cf.rules ?? []) {
        const rule = raw as { type?: string; cfvo?: Array<{ type?: string; value?: unknown }>; color?: Array<Partial<ExcelJS.Color>> }
        if (rule.type !== 'colorScale') continue
        const cfvo = rule.cfvo ?? []
        const colors = rule.color ?? []
        const stops: ColorScaleStop[] = cfvo.map((v, i) => ({
          kind: mapCfvoKind(v?.type),
          value: v?.value != null ? Number(v.value) : null,
          color: colorHex(colors[i]) ?? '#FFFFFF',
        }))
        if (stops.length < 2) continue
        for (const ref of refs) {
          const range = parseA1Range(ref)
          if (range) out.push({ type: 'colorScale', range, stops })
        }
      }
    }
  } catch {
    /* best-effort */
  }
  return out
}

function mapCfvoKind(t: string | undefined): ColorScaleStop['kind'] {
  switch (t) {
    case 'min':
    case 'max':
    case 'num':
    case 'percent':
    case 'percentile':
    case 'formula':
      return t
    default:
      return 'num'
  }
}

// Resolve the workbook theme palette from the theme XML ExcelJS already holds
// (wb.model.themes.theme1) — no re-unzip. Falls back to undefined on any error.
function parseTheme(wb: ExcelJS.Workbook): ThemePalette | undefined {
  try {
    const xml = (wb.model as unknown as { themes?: Record<string, string> }).themes?.theme1
    if (!xml || typeof xml !== 'string') return undefined
    const doc = new DOMParser().parseFromString(xml, 'application/xml')
    const slot = (name: string, dflt: string): string => {
      const el = doc.getElementsByTagName('a:' + name)[0]
      if (!el) return dflt
      const srgb = el.getElementsByTagName('a:srgbClr')[0]
      if (srgb) return ('#' + (srgb.getAttribute('val') ?? '')).toUpperCase()
      const sys = el.getElementsByTagName('a:sysClr')[0]
      if (sys) {
        const last = sys.getAttribute('lastClr')
        if (last) return ('#' + last).toUpperCase()
        return sys.getAttribute('val') === 'window' ? '#FFFFFF' : '#000000'
      }
      return dflt
    }
    return {
      accent1: slot('accent1', '#4472C4'),
      accent2: slot('accent2', '#ED7D31'),
      accent3: slot('accent3', '#A5A5A5'),
      accent4: slot('accent4', '#FFC000'),
      accent5: slot('accent5', '#5B9BD5'),
      accent6: slot('accent6', '#70AD47'),
      dk1: slot('dk1', '#000000'),
      lt1: slot('lt1', '#FFFFFF'),
      dk2: slot('dk2', '#44546A'),
      lt2: slot('lt2', '#E8E8E8'),
    }
  } catch {
    return undefined
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
      const ext = (media.extension as string) === 'jpg' ? 'jpeg' : (media.extension ?? 'png')
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
