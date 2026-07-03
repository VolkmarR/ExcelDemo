// Shared cell → CSS translation for the DOM-based grids (SheetGrid and the
// react-data-grid view). Keeping this in one place means both grids render a
// given CellStyle identically. `decorate()` (sheetStyling.ts) resolves table +
// conditional-format styling into a CellStyle; this turns that into inline CSS.

import type { CSSProperties } from 'react'
import type { BorderSide, CellStyle, HAlign } from '../model'

function borderToCss(side?: BorderSide | null): string | undefined {
  if (!side) return undefined
  const s = side.style.toLowerCase()
  const width = s === 'thick' ? '3px' : s.startsWith('medium') || s === 'double' ? '2px' : '1px'
  const kind = s === 'double' ? 'double' : s.includes('dash') || s === 'dotted' || s === 'hair' ? 'dashed' : 'solid'
  return `${width} ${kind} ${side.color ?? 'var(--grid-line-strong)'}`
}

export function styleToCss(style: CellStyle | null | undefined, align: HAlign): CSSProperties {
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
