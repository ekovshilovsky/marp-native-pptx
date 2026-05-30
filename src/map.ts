import { pxToInch, scaleFor } from './units.js'
import type { LayoutBox, PptxModel, PptxShape, SlideLayout } from './types.js'

export function mapToPptx(
  slides: SlideLayout[],
  slideWidthIn: number,
  slideHeightIn: number,
): PptxModel {
  const out: PptxModel = { slideWidthIn, slideHeightIn, slides: [] }
  for (const slide of slides) {
    const s = scaleFor(slide.sizePx, slideWidthIn, slideHeightIn)
    const shapes: PptxShape[] = slide.boxes.map((b) => mapBox(b, s))
    const bg = slide.background
    const background = bg?.imageDataUri
      ? { imageDataUri: bg.imageDataUri }
      : bg?.fill
        ? { fill: bg.fill }
        : undefined
    out.slides.push({ shapes, background })
  }
  return out
}

function mapRun(r: import('./types.js').Run) {
  return {
    text: r.text,
    bold: r.style.bold,
    italic: r.style.italic,
    color: r.style.color,
    fontFace: r.style.fontFace,
    fontSize: r.style.sizePt,
    highlight: r.style.fill,
    underline: r.style.underline,
    strike: r.style.strike,
  }
}

function mapBox(b: LayoutBox, s: { x: number; y: number }): PptxShape {
  const geom = {
    xIn: pxToInch(b.rect.xPx, s.x),
    yIn: pxToInch(b.rect.yPx, s.y),
    wIn: pxToInch(b.rect.wPx, s.x),
    hIn: pxToInch(b.rect.hPx, s.y),
  }
  if (b.kind === 'image') return { kind: 'image', ...geom, src: b.src }
  if (b.kind === 'table') {
    return {
      kind: 'table',
      ...geom,
      colWidthsIn: b.colWidthsPx.map((w) => pxToInch(w, s.x)),
      rows: b.rows.map((row) =>
        row.map((cell) => ({
          header: cell.header,
          fill: cell.fill,
          align: cell.paras[0]?.align ?? 'left',
          runs: cell.paras.flatMap((p) => p.runs.map(mapRun)),
        })),
      ),
    }
  }
  return {
    kind: 'text',
    ...geom,
    valign: b.valign,
    paragraphs: b.paras.map((p) => ({
      align: p.align,
      lineSpacingPt: p.lineSpacingPt,
      bullet: p.bullet,
      runs: p.runs.map(mapRun),
    })),
  }
}
