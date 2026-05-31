// Auto-layout engine — the reusable geometry + fit primitives that templates
// compose from, so any template (built-in, user-authored, or generated)
// auto-sizes its content and stays aligned. This is the "smart template" core:
// uniform regions (grids/distribution) + shrink-to-fit text sizing.
//
// All units are CSS px on the engine's 1280x720 canvas.

export interface Box {
  x: number
  y: number
  w: number
  h: number
}

export const box = (x: number, y: number, w: number, h: number): Box => ({ x, y, w, h })
export const inset = (b: Box, p: number): Box => ({ x: b.x + p, y: b.y + p, w: b.w - 2 * p, h: b.h - 2 * p })

/** Split a box into a uniform rows×cols grid of equal cells (so content lines up). */
export function gridCells(b: Box, rows: number, cols: number, gapX = 0, gapY = 0): Box[] {
  const cw = (b.w - gapX * (cols - 1)) / cols
  const ch = (b.h - gapY * (rows - 1)) / rows
  const out: Box[] = []
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      out.push({ x: b.x + c * (cw + gapX), y: b.y + r * (ch + gapY), w: cw, h: ch })
    }
  }
  return out
}

/** N evenly-spaced center positions along a box's width (endpoints inclusive). */
export function distributeX(b: Box, n: number): number[] {
  if (n <= 1) return [b.x + b.w / 2]
  const step = b.w / (n - 1)
  return Array.from({ length: n }, (_, i) => b.x + i * step)
}

// Estimate how many wrapped lines `text` takes at `sizePt` within `widthPx`.
function wrappedLines(text: string, sizePt: number, widthPx: number): number {
  const fontPx = sizePt / 0.75
  // 0.55em avg advance — pessimistic vs. real metrics so we over-reserve height
  // (under-reserving causes overlap; over-reserving only leaves a harmless gap).
  const charsPerLine = Math.max(1, Math.floor(widthPx / (fontPx * 0.55)))
  // account for explicit words so a long word doesn't overflow silently
  const longest = text.split(/\s+/).reduce((m, w) => Math.max(m, w.length), 0)
  const perLine = Math.max(charsPerLine, longest)
  return Math.max(1, Math.ceil(text.length / perLine))
}

/** px height a text block needs at a given pt within a width. */
export function textHeight(text: string, sizePt: number, widthPx: number): number {
  return Math.ceil(wrappedLines(text, sizePt, widthPx) * (sizePt / 0.75) * 1.3)
}

/**
 * Shrink-to-fit: the largest point size in [minPt, maxPt] whose wrapped text
 * fits within wPx × hPx. This is the "scale to fit" behavior smart templates
 * use so content never overflows its region.
 */
export function fitFontPt(text: string, wPx: number, hPx: number, maxPt: number, minPt = 9): number {
  for (let pt = Math.round(maxPt); pt >= minPt; pt--) {
    if (textHeight(text, pt, wPx) <= hPx) return pt
  }
  return minPt
}
