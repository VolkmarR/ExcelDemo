import { useMemo, useState } from 'react'
import type { ClipboardEvent } from 'react'
import { DataGrid } from 'react-data-grid'
import type { CellCopyArgs, CellSelectArgs, Column } from 'react-data-grid'
import 'react-data-grid/lib/styles.css'
import type { Cell, HAlign, Merge, Sheet, WorkbookModel } from '../model'
import { cellAddress, columnLabel } from '../model'
import { buildSheetStyling } from '../grid/sheetStyling'
import { styleToCss } from '../grid/cellCss'
import '../grid/SheetGrid.css'
import './ReactDataGridView.css'

// Approach D — the same ClosedXML JSON model that feeds the Backend tab, rendered by
// react-data-grid (adazzle, MIT) instead of the hand-rolled SheetGrid. Holding the data
// source constant makes this a renderer-vs-renderer comparison. Cell/table/CF styling is
// reused verbatim from sheetStyling.decorate() + styleToCss(); the library supplies the
// virtualization, keyboard navigation and frozen gutter for free. See README for the gaps.

const GUTTER_W = 54 // row-number column width (px); matches SheetGrid
const HEAD_H = 24 // column-header height (px)
const ROW_H = 20 // constant row height — DemoData rows are uniform, and a constant keeps the
// picture overlay math exact against react-data-grid's internal scroll.
const DEFAULT_COL_W = 80

interface Props {
  workbook: WorkbookModel
}

// react-data-grid needs a row object per visible row; the sheet row index is all we carry,
// cell content is pulled from cellMap in renderCell.
interface GridRow {
  __r: number
}

export default function ReactDataGridView({ workbook }: Props) {
  const [sheetIndex, setSheetIndex] = useState(0)
  const sheet: Sheet = workbook.sheets[sheetIndex] ?? workbook.sheets[0]
  const { rowCount, colCount } = sheet

  const [active, setActive] = useState({ r: 0, c: 0 })
  const [scroll, setScroll] = useState({ left: 0, top: 0 })

  const cellMap = useMemo(() => {
    const m = new Map<number, Cell>()
    for (const cell of sheet.cells) m.set(cell.r * colCount + cell.c, cell)
    return m
  }, [sheet, colCount])

  const styling = useMemo(() => buildSheetStyling(sheet, workbook.theme), [sheet, workbook.theme])

  // Cumulative pixel offsets for the picture overlay. Rows use the constant ROW_H (same as the
  // grid's rowHeight) so the overlay tracks scroll exactly.
  const { colOffsets, rowOffsets } = useMemo(() => {
    const co: number[] = []
    let acc = 0
    for (let c = 0; c < colCount; c++) {
      co.push(acc)
      acc += sheet.colWidths[c] ?? DEFAULT_COL_W
    }
    const ro: number[] = [0]
    for (let r = 0; r < rowCount; r++) ro.push(ro[r] + ROW_H)
    return { colOffsets: co, rowOffsets: ro }
  }, [sheet, colCount, rowCount])

  // Merges: single-row spans use react-data-grid's native colSpan; merges that cover more than one
  // row (vertical or rectangular blocks) can't — colSpan is horizontal-only — so they're drawn as
  // absolute overlays (same technique as the picture layer), with their covered cells blanked.
  const { spansByCol, blockMerges, overlaidCells } = useMemo(() => {
    const spans = new Map<number, { r0: number; r1: number; span: number }[]>()
    const blocks: Merge[] = []
    const covered = new Set<number>()
    for (const mg of sheet.merges) {
      if (mg.r1 > mg.r0) {
        blocks.push(mg)
        for (let r = mg.r0; r <= mg.r1; r++) for (let c = mg.c0; c <= mg.c1; c++) covered.add(r * colCount + c)
      } else if (mg.c1 > mg.c0) {
        const list = spans.get(mg.c0) ?? []
        list.push({ r0: mg.r0, r1: mg.r1, span: mg.c1 - mg.c0 + 1 })
        spans.set(mg.c0, list)
      }
    }
    return { spansByCol: spans, blockMerges: blocks, overlaidCells: covered }
  }, [sheet, colCount])

  const columns = useMemo<Column<GridRow>[]>(() => {
    const cols: Column<GridRow>[] = [
      {
        key: '__row',
        name: '',
        frozen: true,
        width: GUTTER_W,
        cellClass: 'rdg-gutter',
        headerCellClass: 'rdg-gutter-corner',
        renderCell: ({ row }) => row.__r + 1,
      },
    ]
    for (let c = 0; c < colCount; c++) {
      const spans = spansByCol.get(c)
      cols.push({
        key: String(c),
        name: columnLabel(c),
        width: sheet.colWidths[c] ?? DEFAULT_COL_W,
        headerCellClass: 'rdg-col-head',
        cellClass: 'xcell-host',
        colSpan: spans
          ? (args) => (args.type === 'ROW' ? spans.find((s) => args.row.__r >= s.r0 && args.row.__r <= s.r1)?.span : undefined)
          : undefined,
        renderCell: ({ row, column }) => {
          const r = row.__r
          const cc = Number(column.key)
          if (overlaidCells.has(r * colCount + cc)) return <div className="xcell" /> // hidden under a block-merge overlay
          const cell = cellMap.get(r * colCount + cc)
          const { style, isTableHeader } = styling.decorate(r, cc, cell)
          const align: HAlign = style?.hAlign ?? (typeof cell?.raw === 'number' ? 'right' : 'left')
          return (
            <div className={isTableHeader ? 'xcell xcell-th' : 'xcell'} style={styleToCss(style, align)}>
              {cell?.text}
              {isTableHeader && <span className="filter-glyph">▾</span>}
            </div>
          )
        },
      })
    }
    return cols
  }, [sheet, colCount, cellMap, styling, spansByCol, overlaidCells])

  const rows = useMemo<GridRow[]>(() => Array.from({ length: rowCount }, (_, r) => ({ __r: r })), [rowCount])

  const activeCell = cellMap.get(active.r * colCount + active.c)
  const formulaText = activeCell?.formula ?? activeCell?.text ?? ''
  const nameBoxText = cellAddress(active.r, active.c)

  const onSelectedCellChange = (args: CellSelectArgs<GridRow>) => {
    if (args.column.key === '__row') return // the frozen row-number gutter isn't a real cell
    setActive({ r: args.rowIdx, c: Number(args.column.key) })
  }

  // react-data-grid fires this on Ctrl/Cmd+C; we own writing the clipboard payload.
  const onCellCopy = (args: CellCopyArgs<GridRow>, event: ClipboardEvent<HTMLDivElement>) => {
    if (args.column.key === '__row') return
    const text = cellMap.get(args.row.__r * colCount + Number(args.column.key))?.text ?? ''
    event.clipboardData.setData('text/plain', text)
  }

  return (
    <div className="sheet rdg-view">
      <div className="formula-bar">
        <span className="name-box">{nameBoxText}</span>
        <span className="fx-label">fx</span>
        <span className="fx-content">{formulaText}</span>
      </div>

      <div className="rdg-scroll-wrap">
        <DataGrid
          className="rdg-excel rdg-light"
          columns={columns}
          rows={rows}
          rowKeyGetter={(row) => row.__r}
          rowHeight={ROW_H}
          headerRowHeight={HEAD_H}
          onSelectedCellChange={onSelectedCellChange}
          onCellCopy={onCellCopy}
          onScroll={(e) => setScroll({ left: e.currentTarget.scrollLeft, top: e.currentTarget.scrollTop })}
        />
        {(blockMerges.length > 0 || (sheet.pictures?.length ?? 0) > 0) && (
          <div className="overlay-layer">
            {blockMerges.map((m, i) => {
              const anchor = cellMap.get(m.r0 * colCount + m.c0)
              const { style, isTableHeader } = styling.decorate(m.r0, m.c0, anchor)
              const align: HAlign = style?.hAlign ?? (typeof anchor?.raw === 'number' ? 'right' : 'left')
              let width = 0
              for (let c = m.c0; c <= m.c1; c++) width += sheet.colWidths[c] ?? DEFAULT_COL_W
              return (
                <div
                  key={`m${i}`}
                  className="merge-cell"
                  style={{
                    left: (colOffsets[m.c0] ?? 0) - scroll.left,
                    top: (rowOffsets[m.r0] ?? 0) - scroll.top,
                    width,
                    height: (m.r1 - m.r0 + 1) * ROW_H,
                    ...styleToCss(style, align),
                  }}
                >
                  {anchor?.text}
                  {isTableHeader && <span className="filter-glyph">▾</span>}
                </div>
              )
            })}
            {sheet.pictures?.map((p, i) => (
              <img
                key={`pic${i}`}
                src={p.src}
                alt=""
                draggable={false}
                style={{
                  left: (colOffsets[p.fromCol] ?? 0) + p.offsetX - scroll.left,
                  top: (rowOffsets[p.fromRow] ?? 0) + p.offsetY - scroll.top,
                  width: p.width,
                  height: p.height,
                }}
              />
            ))}
          </div>
        )}
      </div>

      <div className="tabs" role="tablist">
        {workbook.sheets.map((s, i) => (
          <button
            key={s.name + i}
            type="button"
            role="tab"
            aria-selected={i === sheetIndex}
            className={`tab${i === sheetIndex ? ' active' : ''}`}
            onClick={() => {
              setSheetIndex(i)
              setActive({ r: 0, c: 0 })
              setScroll({ left: 0, top: 0 })
            }}
          >
            {s.name}
          </button>
        ))}
      </div>
    </div>
  )
}
