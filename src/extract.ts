import type { LayoutBox, Paragraph, Run, RunStyle } from './types.js'

/**
 * Serialized for injection via page.evaluate. Must be self-contained:
 * no imports referenced at runtime inside the browser. Types above are
 * compile-time only and erased.
 */
export function extractSlide(section: Element): { sizePx: { w: number; h: number }; boxes: LayoutBox[]; background?: { fill?: string; needsRaster?: boolean } } {
  const sRect = section.getBoundingClientRect()
  const origin = { x: sRect.left, y: sRect.top }
  const sizePx = { w: sRect.width, h: sRect.height }
  const boxes: LayoutBox[] = []

  const hex = (rgb: string): string => {
    const m = rgb.match(/\d+(\.\d+)?/g)
    if (!m) return '000000'
    const [r, g, b] = m.map((n) => Math.round(parseFloat(n)))
    return [r, g, b].map((n) => n.toString(16).padStart(2, '0')).join('')
  }
  const isTransparent = (rgb: string): boolean =>
    rgb === 'transparent' || /rgba\(\s*0,\s*0,\s*0,\s*0\s*\)/.test(rgb) || rgb === ''

  const runStyleOf = (el: Element): RunStyle => {
    const cs = getComputedStyle(el)
    const bg = cs.backgroundColor
    const SYS = ['-apple-system', 'BlinkMacSystemFont', 'system-ui', 'ui-sans-serif', 'ui-serif', 'ui-monospace', 'ui-rounded']
    const fams = cs.fontFamily.split(',').map((f) => f.replace(/["']/g, '').trim()).filter(Boolean)
    const fontFace = fams.find((f) => !SYS.includes(f)) ?? fams[0] ?? 'sans-serif'
    return {
      fontFace,
      sizePt: parseFloat(cs.fontSize) * 0.75,
      bold: parseInt(cs.fontWeight, 10) >= 600,
      italic: cs.fontStyle === 'italic',
      color: hex(cs.color),
      fill: isTransparent(bg) ? undefined : hex(bg),
      underline: cs.textDecorationLine.includes('underline'),
      strike: cs.textDecorationLine.includes('line-through'),
    }
  }

  // Segment a block element's inline content into runs.
  const runsOf = (block: Element): Run[] => {
    const runs: Run[] = []
    const walk = (node: Node, inheritedEl: Element) => {
      for (const child of Array.from(node.childNodes)) {
        if (child.nodeType === Node.TEXT_NODE) {
          const text = child.textContent ?? ''
          if (text.length === 0) continue
          runs.push({ text, style: runStyleOf(inheritedEl) })
        } else if (child.nodeType === Node.ELEMENT_NODE) {
          const tag = (child as Element).tagName
          if (tag === 'UL' || tag === 'OL') continue // nested lists are emitted as their own boxes
          walk(child, child as Element)
        }
      }
    }
    walk(block, block)
    // merge adjacent runs with identical style
    const merged: Run[] = []
    for (const r of runs) {
      const last = merged[merged.length - 1]
      if (last && JSON.stringify(last.style) === JSON.stringify(r.style)) last.text += r.text
      else merged.push({ ...r })
    }
    return merged
  }

  const BLOCK = new Set(['H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'P', 'LI', 'BLOCKQUOTE', 'FIGCAPTION', 'HEADER', 'FOOTER'])
  const rectOf = (el: Element) => {
    const r = el.getBoundingClientRect()
    return { xPx: r.left - origin.x, yPx: r.top - origin.y, wPx: r.width, hPx: r.height }
  }

  const visit = (el: Element) => {
    if (el.tagName === 'IMG') {
      boxes.push({ kind: 'image', rect: rectOf(el), src: (el as HTMLImageElement).src })
      return
    }
    if (el.tagName === 'TABLE') {
      const rows: { paras: { runs: ReturnType<typeof runsOf>; align: string; lineSpacingPt: number }[]; fill?: string; header: boolean }[][] = []
      const trs = Array.from(el.querySelectorAll('tr'))
      const colWidthsPx: number[] = []
      for (const tr of trs) {
        const cells = Array.from(tr.children).filter((c) => c.tagName === 'TD' || c.tagName === 'TH')
        const row = cells.map((cell, ci) => {
          const cs = getComputedStyle(cell)
          const bg = cs.backgroundColor
          const r = (cell as HTMLElement).getBoundingClientRect()
          if (rows.length === 0) colWidthsPx[ci] = r.width
          return {
            paras: [{
              runs: runsOf(cell),
              align: (cs.textAlign as string) || 'left',
              lineSpacingPt: parseFloat(cs.lineHeight) * 0.75 || parseFloat(cs.fontSize) * 0.75 * 1.2,
            }],
            fill: isTransparent(bg) ? undefined : hex(bg),
            header: cell.tagName === 'TH',
          }
        })
        if (row.length > 0) rows.push(row)
      }
      if (rows.length > 0) boxes.push({ kind: 'table', rect: rectOf(el), rows: rows as any, colWidthsPx })
      return
    }
    if (BLOCK.has(el.tagName)) {
      const cs = getComputedStyle(el)
      const para: Paragraph = {
        runs: runsOf(el),
        align: (cs.textAlign as Paragraph['align']) || 'left',
        lineSpacingPt: parseFloat(cs.lineHeight) * 0.75 || parseFloat(cs.fontSize) * 0.75 * 1.2,
        bullet: el.tagName === 'LI'
          ? { type: el.parentElement?.tagName === 'OL' ? 'number' : 'bullet', indentLevel: 0 }
          : undefined,
      }
      // M1: only emit a text box if there is real (non-whitespace) text.
      if (para.runs.some((r) => r.text.trim() !== '')) {
        boxes.push({ kind: 'text', rect: rectOf(el), paras: [para], valign: 'top' })
      }
      // C1: capture any <img> nested in this block (e.g. Marp's <p><img></p>),
      // which the run-segmentation does not pick up.
      for (const img of Array.from(el.querySelectorAll('img'))) {
        boxes.push({ kind: 'image', rect: rectOf(img), src: (img as HTMLImageElement).src })
      }
      // I1: descend into nested lists so sub-items are not lost (correct
      // indent levels are deferred to a later plan; this prevents data loss).
      for (const child of Array.from(el.children)) {
        if (child.tagName === 'UL' || child.tagName === 'OL') visit(child)
      }
      return
    }
    for (const child of Array.from(el.children)) visit(child)
  }

  visit(section)
  const pageNo = section.getAttribute('data-marpit-pagination')
  if (pageNo) {
    const fs = parseFloat(getComputedStyle(section).fontSize) || 24
    boxes.push({
      kind: 'text',
      rect: { xPx: sizePx.w - 120, yPx: sizePx.h - 44, wPx: 100, hPx: 28 },
      valign: 'bottom',
      paras: [{
        runs: [{ text: pageNo, style: { fontFace: 'sans-serif', sizePt: fs * 0.75 * 0.7, bold: false, italic: false, color: '888888' } }],
        align: 'right',
        lineSpacingPt: fs * 0.75,
      }],
    })
  }
  const scs = getComputedStyle(section)
  const sectionBg = scs.backgroundColor
  const hasBgImage = scs.backgroundImage !== 'none' && scs.backgroundImage !== ''
  const background = hasBgImage
    ? { needsRaster: true }
    : (isTransparent(sectionBg) ? undefined : { fill: hex(sectionBg) })
  return { sizePx, boxes, background }
}
