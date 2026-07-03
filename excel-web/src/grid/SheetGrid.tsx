import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent, MouseEvent as ReactMouseEvent } from 'react'
import { useVirtualizer } from '@tanstack/react-virtual'
import type { Cell, HAlign, Range as CellRange, Sheet, WorkbookModel } from '../model'
import { cellAddress, columnLabel } from '../model'
import { buildSheetStyling } from './sheetStyling'
import { styleToCss } from './cellCss'
import './SheetGrid.css'

const GUTTER_W = 54 // row-number column width (px)
const HEAD_H = 24 // column-header row height (px); must match .head-row in SheetGrid.css
const DEFAULT_ROW_H = 20
const DEFAULT_COL_W = 80
const EDGE = 24 // distance from a grid edge (px) at which drag auto-scroll engages
const AUTO_SCROLL_SPEED = 18 // px per animation frame while auto-scrolling

interface Props {
  workbook: WorkbookModel
}

interface CellPos {
  r: number
  c: number
}

// An Excel-style selection: one rectangle plus the corners driving extension. `anchor` is the
// fixed corner and also the active ("current") cell shown in the name box; `extent` is the moving
// corner. For a drag/Shift-extend `range === norm(anchor, extent)`; Ctrl+A and Row/Column select
// set `range` explicitly while keeping the anchor put.
interface Selection {
  range: CellRange
  anchor: CellPos
  extent: CellPos
}

const clamp = (v: number, lo: number, hi: number) => (v < lo ? lo : v > hi ? hi : v)

const singleSel = (r: number, c: number): Selection => ({
  range: { r0: r, c0: c, r1: r, c1: c },
  anchor: { r, c },
  extent: { r, c },
})

const norm = (a: CellPos, b: CellPos): CellRange => ({
  r0: Math.min(a.r, b.r),
  c0: Math.min(a.c, b.c),
  r1: Math.max(a.r, b.r),
  c1: Math.max(a.c, b.c),
})

const inRange = (r: number, c: number, x: CellRange) => r >= x.r0 && r <= x.r1 && c >= x.c0 && c <= x.c1

// Grow the selection so its range spans anchor→extent. Returns the same object when the extent is
// unchanged so React can bail out of a re-render (important during auto-scroll, where the pointer
// may sit still against an edge).
const extendTo = (s: Selection, extent: CellPos): Selection => {
  if (s.extent.r === extent.r && s.extent.c === extent.c) return s
  return { ...s, extent, range: norm(s.anchor, extent) }
}

const fmtNum = (n: number) => n.toLocaleString(undefined, { maximumFractionDigits: 2 })

export default function SheetGrid({ workbook }: Props) {
  const [sheetIndex, setSheetIndex] = useState(0)
  const sheet: Sheet = workbook.sheets[sheetIndex] ?? workbook.sheets[0]
  const { rowCount, colCount } = sheet

  const [sel, setSel] = useState<Selection>(() => singleSel(0, 0))
  const [isDragging, setIsDragging] = useState(false) // drives the name-box "NR × MC" readout
  const scrollRef = useRef<HTMLDivElement>(null)

  // Drag state kept in refs so the window listeners read live values without re-binding.
  const dragging = useRef(false)
  const ptr = useRef({ x: 0, y: 0 })
  const vel = useRef({ x: 0, y: 0 })
  const raf = useRef(0)

  // Fast (row,col) -> Cell lookup.
  const cellMap = useMemo(() => {
    const m = new Map<number, Cell>()
    for (const cell of sheet.cells) m.set(cell.r * colCount + cell.c, cell)
    return m
  }, [sheet, colCount])

  // Table + conditional-format styling, resolved once per sheet (exact theme colors).
  const styling = useMemo(() => buildSheetStyling(sheet, workbook.theme), [sheet, workbook.theme])

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

  // Selection predicates — O(1) membership against the single selected rectangle (a full-column /
  // Ctrl+A selection is just one big rect, so this stays cheap even with virtualization).
  const isSelected = (r: number, c: number) => inRange(r, c, sel.range)
  const colInSel = (c: number) => c >= sel.range.c0 && c <= sel.range.c1
  const rowInSel = (r: number) => r >= sel.range.r0 && r <= sel.range.r1

  // Map a client point to a cell. Drag auto-scroll needs this because a stationary pointer at the
  // edge gets no mouseenter as the content scrolls under it, so this is the single drag-extend path.
  const hitTest = useCallback(
    (clientX: number, clientY: number): CellPos => {
      const el = scrollRef.current
      if (!el) return { r: 0, c: 0 }
      const rect = el.getBoundingClientRect()
      const bodyX = clientX - rect.left + el.scrollLeft - GUTTER_W
      const bodyY = clientY - rect.top + el.scrollTop - HEAD_H
      let c = 0
      while (c < colCount - 1 && colOffsets[c + 1] <= bodyX) c++
      // Largest row index whose cumulative offset is <= bodyY (binary search over rowOffsets).
      let lo = 0
      let hi = rowCount - 1
      while (lo < hi) {
        const mid = (lo + hi + 1) >> 1
        if (rowOffsets[mid] <= bodyY) lo = mid
        else hi = mid - 1
      }
      return { r: clamp(lo, 0, rowCount - 1), c: clamp(c, 0, colCount - 1) }
    },
    [colOffsets, rowOffsets, colCount, rowCount],
  )

  // Keep the active cell in view (vertical via virtualizer, horizontal manually). Skip while
  // mouse-dragging — the auto-scroll loop owns scrolling then.
  useEffect(() => {
    if (dragging.current) return
    rowVirtualizer.scrollToIndex(sel.extent.r, { align: 'auto' })
    const el = scrollRef.current
    if (!el) return
    const left = colOffsets[sel.extent.c] ?? 0
    const right = left + (colWidths[sel.extent.c] ?? DEFAULT_COL_W)
    const viewLeft = el.scrollLeft
    const viewWidth = el.clientWidth - GUTTER_W
    if (left < viewLeft) el.scrollLeft = left
    else if (right > viewLeft + viewWidth) el.scrollLeft = right - viewWidth
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sel.extent.r, sel.extent.c, colOffsets, colWidths])

  // Window-level drag tracking + edge auto-scroll. Re-bound only when grid geometry changes
  // (via hitTest), and gated on the dragging ref so it's inert unless a drag is in progress.
  useEffect(() => {
    const extendToPtr = () => setSel((s) => (dragging.current ? extendTo(s, hitTest(ptr.current.x, ptr.current.y)) : s))
    const step = () => {
      if (!dragging.current || (!vel.current.x && !vel.current.y)) {
        raf.current = 0
        return
      }
      const el = scrollRef.current
      if (el) {
        el.scrollLeft += vel.current.x
        el.scrollTop += vel.current.y
        extendToPtr()
      }
      raf.current = requestAnimationFrame(step)
    }
    const onMove = (e: MouseEvent) => {
      if (!dragging.current) return
      ptr.current = { x: e.clientX, y: e.clientY }
      extendToPtr()
      const el = scrollRef.current
      if (!el) return
      const rect = el.getBoundingClientRect()
      let vx = 0
      let vy = 0
      if (e.clientX > rect.right - EDGE) vx = AUTO_SCROLL_SPEED
      else if (e.clientX < rect.left + GUTTER_W + EDGE) vx = -AUTO_SCROLL_SPEED
      if (e.clientY > rect.bottom - EDGE) vy = AUTO_SCROLL_SPEED
      else if (e.clientY < rect.top + HEAD_H + EDGE) vy = -AUTO_SCROLL_SPEED
      vel.current = { x: vx, y: vy }
      if ((vx || vy) && !raf.current) raf.current = requestAnimationFrame(step)
    }
    const onUp = () => {
      if (!dragging.current) return
      dragging.current = false
      vel.current = { x: 0, y: 0 }
      if (raf.current) {
        cancelAnimationFrame(raf.current)
        raf.current = 0
      }
      setIsDragging(false)
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
    return () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
      if (raf.current) {
        cancelAnimationFrame(raf.current)
        raf.current = 0
      }
    }
  }, [hitTest])

  const pageRows = () => Math.max(1, Math.floor((scrollRef.current?.clientHeight ?? 400) / DEFAULT_ROW_H) - 1)

  // Copy the selected rectangle as TSV. localhost is a secure context and keydown is a user gesture.
  const copySelection = () => {
    if (!navigator.clipboard) return
    const x = sel.range
    const lines: string[] = []
    for (let r = x.r0; r <= x.r1; r++) {
      const cols: string[] = []
      for (let c = x.c0; c <= x.c1; c++) cols.push(cellMap.get(r * colCount + c)?.text ?? '')
      lines.push(cols.join('\t'))
    }
    navigator.clipboard.writeText(lines.join('\n')).catch(() => { /* clipboard may be blocked (permissions / non-secure context) */ })
  }

  const onKeyDown = (e: ReactKeyboardEvent) => {
    const shift = e.shiftKey
    const ctrl = e.ctrlKey || e.metaKey

    // Ctrl+A: whole sheet.
    if (ctrl && (e.key === 'a' || e.key === 'A')) {
      setSel((s) => ({ range: { r0: 0, c0: 0, r1: rowCount - 1, c1: colCount - 1 }, anchor: s.anchor, extent: s.anchor })) // keep the active cell where it was
      e.preventDefault()
      return
    }
    // Ctrl+C: copy the live range.
    if (ctrl && (e.key === 'c' || e.key === 'C')) {
      copySelection()
      e.preventDefault()
      return
    }
    // Shift+Space: whole row(s) of the live range; Ctrl+Space: whole column(s).
    if (e.key === ' ' && (shift || ctrl)) {
      setSel((s) => {
        const x = s.range
        const full: CellRange = ctrl ? { r0: 0, r1: rowCount - 1, c0: x.c0, c1: x.c1 } : { r0: x.r0, r1: x.r1, c0: 0, c1: colCount - 1 }
        return { ...s, range: full }
      })
      e.preventDefault()
      return
    }

    let handled = true
    let dr = 0
    let dc = 0
    let homeEnd: 'home' | 'end' | null = null
    switch (e.key) {
      case 'ArrowUp': dr = -1; break
      case 'ArrowDown': dr = 1; break
      case 'ArrowLeft': dc = -1; break
      case 'ArrowRight': dc = 1; break
      case 'Enter': dr = 1; break
      case 'Tab': dc = 1; break
      case 'PageUp': dr = -pageRows(); break
      case 'PageDown': dr = pageRows(); break
      case 'Home': homeEnd = 'home'; break
      case 'End': homeEnd = 'end'; break
      default: handled = false
    }
    if (!handled) return
    e.preventDefault()
    const isArrow = e.key.startsWith('Arrow')
    const jump = ctrl && isArrow // Ctrl+Arrow → jump to the sheet edge
    const extend = shift && (isArrow || e.key === 'PageUp' || e.key === 'PageDown' || e.key === 'Home' || e.key === 'End')

    setSel((s) => {
      // Plain moves start from the active cell (the anchor); Shift-extends move the extent corner.
      const a = extend ? s.extent : s.anchor
      let target: CellPos
      if (homeEnd === 'home') target = { r: ctrl ? 0 : a.r, c: 0 }
      else if (homeEnd === 'end') target = { r: ctrl ? rowCount - 1 : a.r, c: colCount - 1 }
      else if (jump)
        target = {
          r: clamp(dr ? (dr < 0 ? 0 : rowCount - 1) : a.r, 0, rowCount - 1),
          c: clamp(dc ? (dc < 0 ? 0 : colCount - 1) : a.c, 0, colCount - 1),
        }
      else target = { r: clamp(a.r + dr, 0, rowCount - 1), c: clamp(a.c + dc, 0, colCount - 1) }
      return extend ? extendTo(s, target) : singleSel(target.r, target.c)
    })
  }

  const onCellMouseDown = (e: ReactMouseEvent, r: number, c: number) => {
    if (e.button !== 0) return // left button only; leave right/middle for the browser
    scrollRef.current?.focus()
    dragging.current = true
    ptr.current = { x: e.clientX, y: e.clientY }
    setIsDragging(true)
    if (e.shiftKey) setSel((s) => extendTo(s, { r, c })) // extend the range from its anchor
    else setSel(singleSel(r, c)) // fresh single-cell selection
  }

  // Status-bar stats (Count/Sum/Average of the selected cells), Excel-style. Iterating the sparse
  // cell list is cheap; membership is tested against the range rather than a materialized Set.
  const selStats = useMemo(() => {
    let count = 0
    let sum = 0
    let nums = 0
    for (const cell of sheet.cells) {
      if (!inRange(cell.r, cell.c, sel.range)) continue
      if (cell.text !== '') count++
      if (typeof cell.raw === 'number') {
        sum += cell.raw
        nums++
      }
    }
    const totalCells = (sel.range.r1 - sel.range.r0 + 1) * (sel.range.c1 - sel.range.c0 + 1)
    return { count, sum, nums, avg: nums ? sum / nums : 0, totalCells }
  }, [sel, sheet])

  // The active cell resolves to its merge anchor if it sits on a covered cell.
  const activeKey = sel.anchor.r * colCount + sel.anchor.c
  const activeAnchorKey = coveredToAnchor.get(activeKey) ?? activeKey
  const activeCell = cellMap.get(activeKey)
  const formulaText = activeCell?.formula ?? activeCell?.text ?? ''

  const selRows = sel.range.r1 - sel.range.r0 + 1
  const selCols = sel.range.c1 - sel.range.c0 + 1
  const nameBoxText = isDragging && (selRows > 1 || selCols > 1) ? `${selRows}R × ${selCols}C` : cellAddress(sel.anchor.r, sel.anchor.c)

  // Geometry of the single selection outline (absolute, like the merge overlays).
  const selBox = (() => {
    const x = sel.range
    let width = 0
    for (let c = x.c0; c <= x.c1; c++) width += colWidths[c] ?? DEFAULT_COL_W
    return { left: GUTTER_W + (colOffsets[x.c0] ?? 0), top: rowOffsets[x.r0], width, height: rowOffsets[x.r1 + 1] - rowOffsets[x.r0] }
  })()

  const columns = useMemo(() => Array.from({ length: colCount }, (_, c) => c), [colCount])
  const virtualRows = rowVirtualizer.getVirtualItems()
  const totalWidth = bodyWidth + GUTTER_W

  return (
    <div className="sheet">
      <div className="formula-bar">
        <span className="name-box">{nameBoxText}</span>
        <span className="fx-label">fx</span>
        <span className="fx-content">{formulaText}</span>
      </div>

      <div className="grid" ref={scrollRef} tabIndex={0} onKeyDown={onKeyDown} role="grid" aria-rowcount={rowCount} aria-colcount={colCount}>
        {/* Column headers (sticky top). */}
        <div className="head-row" style={{ width: totalWidth }}>
          <div className="corner" style={{ width: GUTTER_W }} />
          {columns.map((c) => (
            <div key={c} className={`col-head${colInSel(c) ? ' hl' : ''}`} style={{ width: colWidths[c] }}>
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
                <div className={`row-head${rowInSel(r) ? ' hl' : ''}`} style={{ width: GUTTER_W }}>
                  {r + 1}
                </div>
                {columns.map((c) => {
                  const key = r * colCount + c
                  if (coveredInFlow.has(key)) return null // rendered as a merge overlay instead
                  const cell = cellMap.get(key)
                  const { style: cellStyle, isTableHeader, inTable } = styling.decorate(r, c, cell)
                  const isNum = typeof cell?.raw === 'number'
                  const align: HAlign = cellStyle?.hAlign ?? (isNum ? 'right' : 'left')
                  const selected = isSelected(r, c)
                  const active = key === activeAnchorKey
                  return (
                    <div
                      key={c}
                      className={`cell${inTable ? ' in-table' : ''}${isTableHeader ? ' table-header' : ''}${selected && !active ? ' sel' : ''}${active ? ' active' : ''}`}
                      style={{ width: colWidths[c], ...styleToCss(cellStyle, align) }}
                      onMouseDown={(e) => onCellMouseDown(e, r, c)}
                    >
                      {cell?.text}
                      {isTableHeader && <span className="filter-glyph">▾</span>}
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
            const { style: cellStyle, isTableHeader, inTable } = styling.decorate(m.r0, m.c0, cell)
            const left = GUTTER_W + (colOffsets[m.c0] ?? 0)
            const top = rowOffsets[m.r0]
            let width = 0
            for (let c = m.c0; c <= m.c1; c++) width += colWidths[c] ?? DEFAULT_COL_W
            const height = rowOffsets[m.r1 + 1] - rowOffsets[m.r0]
            const isNum = typeof cell?.raw === 'number'
            const align: HAlign = cellStyle?.hAlign ?? (isNum ? 'right' : 'left')
            const selected = isSelected(m.r0, m.c0)
            const active = key === activeAnchorKey
            return (
              <div
                key={`m${i}`}
                className={`cell merged${inTable ? ' in-table' : ''}${isTableHeader ? ' table-header' : ''}${selected && !active ? ' sel' : ''}${active ? ' active' : ''}`}
                style={{ position: 'absolute', left, top, width, height, ...styleToCss(cellStyle, align) }}
                onMouseDown={(e) => onCellMouseDown(e, m.r0, m.c0)}
              >
                {cell?.text}
                {isTableHeader && <span className="filter-glyph">▾</span>}
              </div>
            )
          })}

          {/* Selection outline: one crisp green rectangle. Absolute geometry (like the merge
              overlays) so the perimeter is correct even when the range extends past the virtualized
              rows. pointer-events:none so it never eats clicks. */}
          <div className="sel-outline" style={{ left: selBox.left, top: selBox.top, width: selBox.width, height: selBox.height }} />

          {/* Floating pictures (absolute, above the grid; positioned by pixel anchor). */}
          {sheet.pictures?.map((p, i) => (
            <img
              key={`pic${i}`}
              className="cell-picture"
              src={p.src}
              alt=""
              draggable={false}
              style={{
                position: 'absolute',
                left: GUTTER_W + (colOffsets[p.fromCol] ?? 0) + p.offsetX,
                top: (rowOffsets[p.fromRow] ?? 0) + p.offsetY,
                width: p.width,
                height: p.height,
              }}
            />
          ))}
        </div>
      </div>

      {/* Sheet tabs + Excel-style selection summary. */}
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
              setSel(singleSel(0, 0))
              rowVirtualizer.scrollToIndex(0)
              if (scrollRef.current) scrollRef.current.scrollLeft = 0
            }}
          >
            {s.name}
          </button>
        ))}
        {selStats.totalCells > 1 && (
          <span className="sel-stats">
            {selStats.nums > 0 && <span>Average: {fmtNum(selStats.avg)}</span>}
            <span>Count: {selStats.count}</span>
            {selStats.nums > 0 && <span>Sum: {fmtNum(selStats.sum)}</span>}
          </span>
        )}
      </div>
    </div>
  )
}
