// Shared cell → CSS translation for the DOM-based grids (SheetGrid and the
// react-data-grid view). Keeping this in one place means both grids render a
// given CellStyle identically. `decorate()` (sheetStyling.ts) resolves table +
// conditional-format styling into a CellStyle; this turns that into inline CSS.
// `styleToCss` returns a React style object (for the flex-based cell <div>s);
// `styleToCssText` returns a CSS declaration string for Jspreadsheet CE, whose
// per-cell `style` option takes `{ A1: "prop:val;…" }` applied to <td>s.

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

// Same visual result as styleToCss, but emitted as a CSS declaration string for a
// table cell (<td>). Jspreadsheet CE applies these via its `style` option, so there
// is no flex container — alignment uses text-align / vertical-align directly.
export function styleToCssText(style: CellStyle | null | undefined, align: HAlign): string {
  const parts: string[] = [`text-align:${align}`]
  if (style) {
    if (style.bold) parts.push('font-weight:700')
    if (style.italic) parts.push('font-style:italic')
    if (style.underline) parts.push('text-decoration:underline')
    if (style.color) parts.push(`color:${style.color}`)
    if (style.bg) parts.push(`background-color:${style.bg}`)
    if (style.fontFamily) parts.push(`font-family:${style.fontFamily}`)
    if (style.fontSize) parts.push(`font-size:${style.fontSize}pt`)
    if (style.vAlign) parts.push(`vertical-align:${style.vAlign}`)
    if (style.wrap) parts.push('white-space:normal', 'word-break:break-word')
    const b = style.border
    if (b) {
      const t = borderToCss(b.top)
      const r = borderToCss(b.right)
      const bo = borderToCss(b.bottom)
      const l = borderToCss(b.left)
      if (t) parts.push(`border-top:${t}`)
      if (r) parts.push(`border-right:${r}`)
      if (bo) parts.push(`border-bottom:${bo}`)
      if (l) parts.push(`border-left:${l}`)
    }
  }
  return parts.join(';')
}
