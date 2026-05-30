import { describe, it, expect } from 'vitest'
import { mapToPptx } from '../src/map.js'
import type { SlideLayout } from '../src/types.js'

const slide: SlideLayout = {
  sizePx: { w: 1280, h: 720 },
  boxes: [
    {
      kind: 'table',
      rect: { xPx: 64, yPx: 100, wPx: 640, hPx: 120 },
      colWidthsPx: [320, 320],
      rows: [
        [
          { header: true, fill: 'eeeeee', paras: [{ align: 'left', lineSpacingPt: 24, runs: [{ text: 'Character', style: { fontFace: 'Arial', sizePt: 14, bold: true, italic: false, color: '111111' } }] }] },
          { header: true, fill: 'eeeeee', paras: [{ align: 'left', lineSpacingPt: 24, runs: [{ text: 'Show', style: { fontFace: 'Arial', sizePt: 14, bold: true, italic: false, color: '111111' } }] }] },
        ],
        [
          { header: false, paras: [{ align: 'left', lineSpacingPt: 24, runs: [{ text: 'Pixel the Robot', style: { fontFace: 'Arial', sizePt: 14, bold: false, italic: false, color: '111111' } }] }] },
          { header: false, paras: [{ align: 'left', lineSpacingPt: 24, runs: [{ text: 'Pixel Park', style: { fontFace: 'Arial', sizePt: 14, bold: false, italic: false, color: '111111' } }] }] },
        ],
      ],
    },
  ],
}

describe('mapToPptx tables', () => {
  it('maps a table box to a PptxTable with inches and cell runs', () => {
    const model = mapToPptx([slide], 13.333, 7.5)
    const shape = model.slides[0].shapes[0]
    expect(shape.kind).toBe('table')
    if (shape.kind !== 'table') throw new Error('expected table')
    expect(shape.colWidthsIn.length).toBe(2)
    expect(shape.colWidthsIn[0]).toBeCloseTo(3.33325, 3)
    expect(shape.rows.length).toBe(2)
    expect(shape.rows[0][0].header).toBe(true)
    expect(shape.rows[0][0].fill).toBe('eeeeee')
    expect(shape.rows[1][0].runs[0].text).toBe('Pixel the Robot')
    expect(shape.rows[0][0].runs[0].bold).toBe(true)
  })
})
