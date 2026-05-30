import { describe, it, expect } from 'vitest'
import { mapToPptx } from '../src/map.js'
import type { SlideLayout } from '../src/types.js'

const slide: SlideLayout = {
  sizePx: { w: 1280, h: 720 },
  boxes: [
    {
      kind: 'text',
      rect: { xPx: 64, yPx: 36, wPx: 640, hPx: 48 },
      valign: 'top',
      paras: [
        {
          align: 'left',
          lineSpacingPt: 24,
          runs: [
            { text: 'See ', style: { fontFace: 'Arial', sizePt: 16, bold: false, italic: false, color: '111111' } },
            { text: 'code.ts', style: { fontFace: 'Menlo', sizePt: 14, bold: false, italic: false, color: '9333ea', fill: 'f3e8ff' } },
          ],
        },
      ],
    },
  ],
}

describe('mapToPptx', () => {
  it('maps a slide to a model with inches and a highlighted code run', () => {
    const model = mapToPptx([slide], 13.333, 7.5)
    expect(model.slides.length).toBe(1)
    const box = model.slides[0].shapes[0]
    expect(box.kind).toBe('text')
    if (box.kind !== 'text') throw new Error('expected text')
    expect(box.xIn).toBeCloseTo(0.6665, 3)        // 64px * (13.333/1280)
    expect(box.wIn).toBeCloseTo(6.6665, 3)
    const runs = box.paragraphs[0].runs
    expect(runs).toHaveLength(2)
    expect(runs[1].highlight).toBe('f3e8ff')       // chip fill preserved
    expect(runs[1].fontFace).toBe('Menlo')
    expect(runs[1].fontSize).toBe(14)
  })
})
