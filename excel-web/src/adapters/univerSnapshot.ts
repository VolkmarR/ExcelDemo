import { BooleanNumber, HorizontalAlign, VerticalAlign, WrapStrategy } from '@univerjs/presets'
import type { ICellData, IStyleData, IWorkbookData } from '@univerjs/presets'
import type { CellStyle, WorkbookModel } from '../model'

// Univer uses terse style keys (bl=bold, it=italic, cl=color, bg=fill, ht/vt=align, tb=wrap).
function mapStyle(s: CellStyle): IStyleData {
  const st: IStyleData = {}
  if (s.bold) st.bl = BooleanNumber.TRUE
  if (s.italic) st.it = BooleanNumber.TRUE
  if (s.underline) st.ul = { s: BooleanNumber.TRUE }
  if (s.color) st.cl = { rgb: s.color }
  if (s.bg) st.bg = { rgb: s.bg }
  if (s.fontFamily) st.ff = s.fontFamily
  if (s.fontSize) st.fs = s.fontSize
  if (s.hAlign) st.ht = s.hAlign === 'left' ? HorizontalAlign.LEFT : s.hAlign === 'center' ? HorizontalAlign.CENTER : HorizontalAlign.RIGHT
  if (s.vAlign) st.vt = s.vAlign === 'top' ? VerticalAlign.TOP : s.vAlign === 'middle' ? VerticalAlign.MIDDLE : VerticalAlign.BOTTOM
  if (s.wrap) st.tb = WrapStrategy.WRAP
  // Borders are omitted for the Univer path to keep the POC mapping small.
  return st
}

/**
 * Approach C — map the shared WorkbookModel into a Univer IWorkbookData snapshot,
 * loaded fully client-side via univerAPI.createWorkbook (no server exchange service).
 */
export function toUniverSnapshot(model: WorkbookModel): IWorkbookData {
  const sheetOrder: string[] = []
  const sheets: Record<string, unknown> = {}

  model.sheets.forEach((sheet, i) => {
    const id = `sheet-${i}`
    sheetOrder.push(id)

    const cellData: Record<number, Record<number, ICellData>> = {}
    for (const cell of sheet.cells) {
      const rowObj = cellData[cell.r] ?? (cellData[cell.r] = {})
      const cd: ICellData = {
        v: typeof cell.raw === 'number' || typeof cell.raw === 'boolean' ? cell.raw : cell.text,
      }
      if (cell.formula) cd.f = cell.formula
      if (cell.style) cd.s = mapStyle(cell.style)
      rowObj[cell.c] = cd
    }

    const columnData: Record<number, { w: number }> = {}
    sheet.colWidths.forEach((w, c) => {
      columnData[c] = { w }
    })
    const rowData: Record<number, { h: number }> = {}
    sheet.rowHeights.forEach((h, r) => {
      rowData[r] = { h }
    })

    sheets[id] = {
      id,
      name: sheet.name,
      rowCount: Math.max(sheet.rowCount, 1),
      columnCount: Math.max(sheet.colCount, 1),
      cellData,
      mergeData: sheet.merges.map((m) => ({ startRow: m.r0, startColumn: m.c0, endRow: m.r1, endColumn: m.c1 })),
      columnData,
      rowData,
      ...(sheet.freeze
        ? { freeze: { xSplit: sheet.freeze.cols, ySplit: sheet.freeze.rows, startRow: sheet.freeze.rows, startColumn: sheet.freeze.cols } }
        : {}),
    }
  })

  return {
    id: 'excel-preview',
    name: 'DemoData',
    sheetOrder,
    styles: {},
    sheets,
  } as unknown as IWorkbookData
}
