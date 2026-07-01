import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { CSSProperties, KeyboardEvent as ReactKeyboardEvent } from 'react'
import { useVirtualizer } from '@tanstack/react-virtual'
import type { BorderSide, Cell, CellStyle, HAlign, Sheet, WorkbookModel } from '../model'
import { cellAddress, columnLabel } from '../model'
import './SheetGrid.css'

const GUTTER_W = 54 // row-number column width (px)
const DEFAULT_ROW_H = 20
const DEFAULT_COL_W = 80

interface Props {
  workbook: WorkbookModel
}

const clamp = (v: number, lo: number, hi: number) => (v < lo ? lo : v > hi ? hi : v)

function borderToCss(side?: BorderSide | null): string | undefined {
  if (!side) return undefined
  const s = side.style.toLowerCase()
  const width = s === 'thick' ? '3px' : s.startsWith('medium') || s === 'double' ? '2px' : '1px'
  const kind = s === 'double' ? 'double' : s.includes('dash') || s === 'dotted' || s === 'hair' ? 'dashed' : 'solid'
  return `${width} ${kind} ${side.color ?? 'var(--grid-line-strong)'}`
}

function styleToCss(style: CellStyle | null | undefined, align: HAlign): CSSProperties {
  const css: CSSProperties = {
    justifyContent: align === 'right' ? 'flex-end' : align === 'center' ? 'center' : 'flex-start',
    textAlign: align,
  }
  if (!style) return css
  if (style.bold) css.fontWeight = 700
  if (style.italic) css.fontStyle = 'italic'
  if (style.underline) css.textDecoration = 'underline'
  if (style.color) css.color = style.color
  if (style.bg) css.background = style.bg
  if (style.fontFamily) css.fontFamily = style.fontFamily
  if (style.fontSize) css.fontSize = `${style.fontSize}pt`
  if (style.vAlign) css.alignItems = style.vAlign === 'top' ? 'flex-start' : style.vAlign === 'middle' ? 'center' : 'flex-end'
  if (style.wrap) {
    css.whiteSpace = 'normal'
    css.wordBreak = 'break-word'
  }
  const b = style.border
  if (b) {
    const t = borderToCss(b.top)
    const r = borderToCss(b.right)
    const bo = borderToCss(b.bottom)
    const l = borderToCss(b.left)
    if (t) css.borderTop = t
    if (r) css.borderRight = r
    if (bo) css.borderBottom = bo
    if (l) css.borderLeft = l
  }
  return css
}

export default function SheetGrid({ workbook }: Props) {
  const [sheetIndex, setSheetIndex] = useState(0)
  const sheet: Sheet = workbook.sheets[sheetIndex] ?? workbook.sheets[0]
  const { rowCount, colCount } = sheet

  const [focus, setFocus] = useState({ r: 0, c: 0 })
  const scrollRef = useRef<HTMLDivElement>(null)

  // Fast (row,col) -> Cell lookup.
  const cellMap = useMemo(() => {
    const m = new Map<number, Cell>()
    for (const cell of sheet.cells) m.set(cell.r * colCount + cell.c, cell)
    return m
  }, [sheet, colCount])

  // Column geometry.
  const { colWidths, colOffsets, bodyWidth } = useMemo(() => {
    const widths: number[] = []
    const offsets: number[] = []
    let acc = 0
    for (let c = 0; c < colCount; c++) {
      offsets.push(acc)
      const w = sheet.colWidths[c] ?? DEFAULT_COL_W
      widths.push(w)
      acc += w
    }
    return { colWidths: widths, colOffsets: offsets, bodyWidth: acc }
  }, [sheet, colCount])

  // Cumulative row offsets (match the virtualizer, which uses the same estimateSize).
  const rowOffsets = useMemo(() => {
    const offs = new Float64Array(rowCount + 1)
    for (let r = 0; r < rowCount; r++) offs[r + 1] = offs[r] + (sheet.rowHeights[r] ?? DEFAULT_ROW_H)
    return offs
  }, [sheet, rowCount])

  // Merges: which cells are covered, and geometry for each anchor overlay.
  const { coveredInFlow, coveredToAnchor, mergeBlocks } = useMemo(() => {
    const covered = new Set<number>()
    const toAnchor = new Map<number, number>()
    const blocks = sheet.merges.map((m) => {
      const anchorKey = m.r0 * colCount + m.c0
      for (let r = m.r0; r <= m.r1; r++) {
        for (let c = m.c0; c <= m.c1; c++) {
          const k = r * colCount + c
          covered.add(k)
          toAnchor.set(k, anchorKey)
        }
      }
      return m
    })
    return { coveredInFlow: covered, coveredToAnchor: toAnchor, mergeBlocks: blocks }
  }, [sheet, colCount])

  const rowVirtualizer = useVirtualizer({
    count: rowCount,
    getScrollElement: () => scrollRef.current,
    estimateSize: (i) => sheet.rowHeights[i] ?? DEFAULT_ROW_H,
    overscan: 12,
  })

  // Keep the focused cell scrolled into view (vertical via virtualizer, horizontal manually).
  useEffect(() => {
    rowVirtualizer.scrollToIndex(focus.r, { align: 'auto' })
    const el = scrollRef.current
    if (!el) return
    const left = colOffsets[focus.c] ?? 0
    const right = left + (colWidths[focus.c] ?? DEFAULT_COL_W)
    const viewLeft = el.scrollLeft
    const viewWidth = el.clientWidth - GUTTER_W
    if (left < viewLeft) el.scrollLeft = left
    else if (right > viewLeft + viewWidth) el.scrollLeft = right - viewWidth
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [focus, colOffsets, colWidths])

  const pageRows = () => Math.max(1, Math.floor((scrollRef.current?.clientHeight ?? 400) / DEFAULT_ROW_H) - 1)

  const onKeyDown = useCallback(
    (e: ReactKeyboardEvent) => {
      let handled = true
      setFocus((f) => {
        let { r, c } = f
        switch (e.key) {
          case 'ArrowUp': r--; break
          case 'ArrowDown': case 'Enter': r++; break
          case 'ArrowLeft': c--; break
          case 'ArrowRight': case 'Tab': c++; break
          case 'PageUp': r -= pageRows(); break
          case 'PageDown': r += pageRows(); break
          case 'Home': c = 0; if (e.ctrlKey) r = 0; break
          case 'End': c = colCount - 1; if (e.ctrlKey) r = rowCount - 1; break
          default: handled = false
        }
        return { r: clamp(r, 0, rowCount - 1), c: clamp(c, 0, colCount - 1) }
      })
      if (handled) e.preventDefault()
    },
    [rowCount, colCount],
  )

  const selectCell = (r: number, c: number) => {
    setFocus({ r, c })
    scrollRef.current?.focus()
  }

  // The focused cell resolves to its merge anchor if it sits on a covered cell.
  const focusKey = focus.r * colCount + focus.c
  const focusAnchorKey = coveredToAnchor.get(focusKey) ?? focusKey
  const focusedCell = cellMap.get(focusKey)
  const formulaText = focusedCell?.formula ?? focusedCell?.text ?? ''

  const columns = useMemo(() => Array.from({ length: colCount }, (_, c) => c), [colCount])
  const virtualRows = rowVirtualizer.getVirtualItems()
  const totalWidth = bodyWidth + GUTTER_W

  return (
    <div className="sheet">
      <div className="formula-bar">
        <span className="name-box">{cellAddress(focus.r, focus.c)}</span>
        <span className="fx-label">fx</span>
        <span className="fx-content">{formulaText}</span>
      </div>

      <div className="grid" ref={scrollRef} tabIndex={0} onKeyDown={onKeyDown} role="grid" aria-rowcount={rowCount} aria-colcount={colCount}>
        {/* Column headers (sticky top). */}
        <div className="head-row" style={{ width: totalWidth }}>
          <div className="corner" style={{ width: GUTTER_W }} />
          {columns.map((c) => (
            <div key={c} className={`col-head${focus.c === c ? ' hl' : ''}`} style={{ width: colWidths[c] }}>
              {columnLabel(c)}
            </div>
          ))}
        </div>

        {/* Body (absolute rows). */}
        <div className="body" style={{ height: rowVirtualizer.getTotalSize(), width: totalWidth }}>
          {virtualRows.map((vr) => {
            const r = vr.index
            return (
              <div
                key={r}
                className={`row${r % 2 === 1 ? ' odd' : ''}`}
                style={{ position: 'absolute', top: vr.start, height: vr.size, width: totalWidth }}
              >
                <div className={`row-head${focus.r === r ? ' hl' : ''}`} style={{ width: GUTTER_W }}>
                  {r + 1}
                </div>
                {columns.map((c) => {
                  const key = r * colCount + c
                  if (coveredInFlow.has(key)) return null // rendered as a merge overlay instead
                  const cell = cellMap.get(key)
                  const isNum = typeof cell?.raw === 'number'
                  const align: HAlign = cell?.style?.hAlign ?? (isNum ? 'right' : 'left')
                  const focused = key === focusAnchorKey
                  return (
                    <div
                      key={c}
                      className={`cell${focused ? ' focused' : ''}`}
                      style={{ width: colWidths[c], ...styleToCss(cell?.style, align) }}
                      onMouseDown={() => selectCell(r, c)}
                    >
                      {cell?.text}
                    </div>
                  )
                })}
              </div>
            )
          })}

          {/* Merge overlays (absolute, above the flow). */}
          {mergeBlocks.map((m, i) => {
            const key = m.r0 * colCount + m.c0
            const cell = cellMap.get(key)
            const left = GUTTER_W + (colOffsets[m.c0] ?? 0)
            const top = rowOffsets[m.r0]
            let width = 0
            for (let c = m.c0; c <= m.c1; c++) width += colWidths[c] ?? DEFAULT_COL_W
            const height = rowOffsets[m.r1 + 1] - rowOffsets[m.r0]
            const isNum = typeof cell?.raw === 'number'
            const align: HAlign = cell?.style?.hAlign ?? (isNum ? 'right' : 'left')
            const focused = key === focusAnchorKey
            return (
              <div
                key={`m${i}`}
                className={`cell merged${focused ? ' focused' : ''}`}
                style={{ position: 'absolute', left, top, width, height, ...styleToCss(cell?.style, align) }}
                onMouseDown={() => selectCell(m.r0, m.c0)}
              >
                {cell?.text}
              </div>
            )
          })}
        </div>
      </div>

      {/* Sheet tabs. */}
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
              setFocus({ r: 0, c: 0 })
              rowVirtualizer.scrollToIndex(0)
              if (scrollRef.current) scrollRef.current.scrollLeft = 0
            }}
          >
            {s.name}
          </button>
        ))}
      </div>
    </div>
  )
}
