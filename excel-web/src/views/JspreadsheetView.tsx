import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import { Spreadsheet, Worksheet } from '@jspreadsheet-ce/react'
import 'jspreadsheet-ce/dist/jspreadsheet.css'
import 'jsuites/dist/jsuites.css'
import type { Cell, HAlign, Sheet, WorkbookModel } from '../model'
import { cellAddress } from '../model'
import { buildSheetStyling } from '../grid/sheetStyling'
import { styleToCssText } from '../grid/cellCss'
import '../grid/SheetGrid.css'
import './JspreadsheetView.css'

// Approach E — the same ClosedXML JSON model that feeds the Backend tab, rendered by
// Jspreadsheet CE (MIT) through its React wrapper. This tab exists to answer whether the
// two features a prior analysis flagged can be supported:
//   • Conditional formatting — YES: CE has no CF engine, but decorate() already resolves the
//     color scale + table styling into a concrete per-cell CellStyle, which we hand to CE's
//     `style` option as CSS text. Full fidelity, same colors as the Backend/rdg tabs.
//   • Floating picture — YES: CE's native floating images are Pro-only, so (exactly like
//     SheetGrid and rdg) we draw the photo as our own absolute <img> overlay synced to CE's
//     scroll. Merges, by contrast, ARE native here (mergeCells, both orientations).
// One worksheet is rendered at a time with our own tab strip (remounted via `key`), so there
// is a single, predictable `.jss_content` scroller for the overlay + formula bar to track.

const DEFAULT_COL_W = 80

interface Props {
  workbook: WorkbookModel
}

interface Geom {
  headerH: number // sticky column-header height (px)
  indexW: number // row-number column width (px)
  colX: number[] // cumulative left offset of each data column, measured from CE's header cells
  rowH: number // uniform body row height (px)
}

export default function JspreadsheetView({ workbook }: Props) {
  const [sheetIndex, setSheetIndex] = useState(0)
  const sheet: Sheet = workbook.sheets[sheetIndex] ?? workbook.sheets[0]
  const { rowCount, colCount } = sheet

  const [active, setActive] = useState({ r: 0, c: 0 })
  const [scroll, setScroll] = useState({ left: 0, top: 0 })
  const [gridH, setGridH] = useState(0)
  const [geom, setGeom] = useState<Geom | null>(null)

  const wrapRef = useRef<HTMLDivElement>(null)
  // Holds the CE instance the wrapper assigns. Reset to null on sheet switch so the freshly
  // keyed <Spreadsheet> rebuilds (the wrapper only builds when the ref is empty).
  const ssRef = useRef<unknown>(null)

  const cellMap = useMemo(() => {
    const m = new Map<number, Cell>()
    for (const cell of sheet.cells) m.set(cell.r * colCount + cell.c, cell)
    return m
  }, [sheet, colCount])

  const styling = useMemo(() => buildSheetStyling(sheet, workbook.theme), [sheet, workbook.theme])

  // Map the model to CE's worksheet inputs. Table + color-scale conditional formatting is
  // resolved by decorate() and carried through the per-cell `style` map (CSS text). Merges use
  // CE's native mergeCells ([colspan, rowspan]) — no overlay needed for either orientation.
  const { data, columns, styleMap, mergeMap } = useMemo(() => {
    const data: string[][] = Array.from({ length: rowCount }, () => new Array<string>(colCount).fill(''))
    const styleMap: Record<string, string> = {}
    for (const cell of sheet.cells) {
      if (cell.r < rowCount && cell.c < colCount) data[cell.r][cell.c] = cell.text ?? ''
      const { style } = styling.decorate(cell.r, cell.c, cell)
      const align: HAlign = style?.hAlign ?? (typeof cell.raw === 'number' ? 'right' : 'left')
      const css = styleToCssText(style, align)
      if (css) styleMap[cellAddress(cell.r, cell.c)] = css
    }
    const columns = Array.from({ length: colCount }, (_, c) => ({
      type: 'text',
      width: sheet.colWidths[c] ?? DEFAULT_COL_W,
      readOnly: true,
    }))
    const mergeMap: Record<string, [number, number]> = {}
    for (const m of sheet.merges) mergeMap[cellAddress(m.r0, m.c0)] = [m.c1 - m.c0 + 1, m.r1 - m.r0 + 1]
    return { data, columns, styleMap, mergeMap }
  }, [sheet, rowCount, colCount, styling])

  const activeCell = cellMap.get(active.r * colCount + active.c)
  const formulaText = activeCell?.formula ?? activeCell?.text ?? ''

  // Measure the available height once, before building CE, so tableHeight (which CE reads only
  // at build time) matches the viewport. Resizing afterwards keeps the initial height — a POC
  // limitation, documented in the README.
  useLayoutEffect(() => {
    const el = wrapRef.current
    if (!el) return
    if (el.clientHeight > 0) {
      setGridH(el.clientHeight)
      return
    }
    const raf = requestAnimationFrame(() => {
      if (wrapRef.current) setGridH(wrapRef.current.clientHeight || 400)
    })
    return () => cancelAnimationFrame(raf)
  }, [])

  // After CE builds, read its real geometry (header height, index-column width, per-column
  // offsets, row height) from the DOM and follow the scroll container. Must be useEffect, not
  // useLayoutEffect: the wrapper builds CE in a useEffect, and every useLayoutEffect fires
  // before any useEffect, so .jss_content would not exist yet. A rAF retry covers lazy layout.
  useEffect(() => {
    if (!gridH) return
    let raf = 0
    let tries = 0
    let bound: HTMLElement | null = null
    const onScroll = () => {
      if (bound) setScroll({ left: bound.scrollLeft, top: bound.scrollTop })
    }
    const bind = () => {
      const content = wrapRef.current?.querySelector('.jss_content') as HTMLElement | null
      const thead = content?.querySelector('thead') as HTMLElement | null
      const headCells = thead ? Array.from(thead.querySelectorAll('td')) : []
      if (!content || !thead || headCells.length < 2) {
        if (tries++ < 30) raf = requestAnimationFrame(bind)
        return
      }
      const colX: number[] = []
      let acc = 0
      for (let c = 0; c < colCount; c++) {
        colX.push(acc)
        acc += (headCells[c + 1] as HTMLElement | undefined)?.offsetWidth ?? DEFAULT_COL_W
      }
      const firstRow = content.querySelector('tbody tr') as HTMLElement | null
      setGeom({
        headerH: thead.offsetHeight,
        indexW: (headCells[0] as HTMLElement).offsetWidth,
        colX,
        rowH: firstRow?.offsetHeight || 20,
      })
      bound = content
      content.addEventListener('scroll', onScroll, { passive: true })
    }
    raf = requestAnimationFrame(bind)
    return () => {
      cancelAnimationFrame(raf)
      if (bound) bound.removeEventListener('scroll', onScroll)
    }
  }, [gridH, sheetIndex, colCount])

  const selectSheet = (i: number) => {
    ssRef.current = null // force the keyed <Spreadsheet> to rebuild for the new sheet
    setSheetIndex(i)
    setActive({ r: 0, c: 0 })
    setScroll({ left: 0, top: 0 })
    setGeom(null)
  }

  return (
    <div className="sheet jss-view">
      <div className="formula-bar">
        <span className="name-box">{cellAddress(active.r, active.c)}</span>
        <span className="fx-label">fx</span>
        <span className="fx-content">{formulaText}</span>
      </div>

      <div className="jss-wrap" ref={wrapRef}>
        {gridH > 0 && (
          <Spreadsheet
            key={sheetIndex}
            ref={ssRef}
            tabs={false}
            onselection={(_instance: unknown, x1: number, y1: number) => setActive({ r: y1, c: x1 })}
          >
            <Worksheet
              data={data}
              columns={columns}
              style={styleMap}
              mergeCells={mergeMap}
              minDimensions={[colCount, rowCount]}
              worksheetName={sheet.name}
              editable={false}
              tableOverflow={true}
              tableHeight={gridH}
              lazyLoading={true}
            />
          </Spreadsheet>
        )}

        {geom && (sheet.pictures?.length ?? 0) > 0 && (
          <div className="overlay-layer" style={{ top: geom.headerH }}>
            {sheet.pictures?.map((p, i) => (
              <img
                key={`pic${i}`}
                src={p.src}
                alt=""
                draggable={false}
                style={{
                  left: geom.indexW + (geom.colX[p.fromCol] ?? 0) + p.offsetX - scroll.left,
                  top: p.fromRow * geom.rowH + p.offsetY - scroll.top,
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
            onClick={() => selectSheet(i)}
          >
            {s.name}
          </button>
        ))}
      </div>
    </div>
  )
}
