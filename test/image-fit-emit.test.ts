import { describe, expect, it } from 'vitest'
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PNG } from 'pngjs'
import { emit } from '../src/emit.js'
import { slideXml } from './helpers/pptx-xml.js'
import type { ImageFit, PptxModel } from '../src/types.js'

// A real PNG of the given size as a data URI, so emit can decode its aspect.
function pngDataUri(w: number, h: number): string {
  const png = new PNG({ width: w, height: h })
  png.data.fill(0xff)
  return 'data:image/png;base64,' + PNG.sync.write(png).toString('base64')
}

// One image shape, `src` at `fit`, in a 2x2in box on a 10x7.5in slide.
function oneImageDeck(src: string, fit?: ImageFit): PptxModel {
  return {
    slideWidthIn: 10,
    slideHeightIn: 7.5,
    slides: [{ shapes: [{ kind: 'image', xIn: 0, yIn: 0, wIn: 2, hIn: 2, src, fit }] }],
  }
}

async function emitAndRead(model: PptxModel): Promise<string> {
  const out = join(mkdtempSync(join(tmpdir(), 'mnp-fit-')), 'deck.pptx')
  await emit(model, out)
  return (await slideXml(out))[0]
}

// pull the four srcRect crop percentages (l/r/t/b) out of the slide XML, if any
function srcRect(xml: string): { l: number; r: number; t: number; b: number } | null {
  const m = xml.match(/<a:srcRect l="(-?\d+)" r="(-?\d+)" t="(-?\d+)" b="(-?\d+)"\/>/)
  return m ? { l: +m[1], r: +m[2], t: +m[3], b: +m[4] } : null
}

describe('emit image fit', () => {
  it("fill (default) stretches to the box — no crop", async () => {
    const xml = await emitAndRead(oneImageDeck(pngDataUri(4, 2))) // no fit
    expect(srcRect(xml)).toBeNull()
    expect(xml).toContain('<a:stretch><a:fillRect/></a:stretch>')
  })

  it('cover on a wide image into a square box crops the sides, not top/bottom', async () => {
    const xml = await emitAndRead(oneImageDeck(pngDataUri(4, 2), 'cover')) // 2:1 into 1:1
    const rect = srcRect(xml)
    expect(rect, 'a crop was applied').toBeTruthy()
    if (!rect) return
    expect(rect.l).toBeGreaterThan(0) // sides cropped...
    expect(rect.r).toBeGreaterThan(0)
    expect(rect.t).toBe(0) // ...top/bottom untouched
    expect(rect.b).toBe(0)
    // 2:1 covering 1:1 keeps the middle half → ~25% off each side
    expect(rect.l).toBeGreaterThan(20000)
    expect(rect.l).toBeLessThan(30000)
  })

  it('cover on a tall image into a square box crops top/bottom, not sides', async () => {
    const xml = await emitAndRead(oneImageDeck(pngDataUri(2, 4), 'cover')) // 1:2 into 1:1
    const rect = srcRect(xml)
    expect(rect).toBeTruthy()
    if (!rect) return
    expect(rect.t).toBeGreaterThan(0)
    expect(rect.b).toBeGreaterThan(0)
    expect(rect.l).toBe(0)
    expect(rect.r).toBe(0)
  })

  it('contain on a wide image into a square box letterboxes (negative top/bottom inset, no side crop)', async () => {
    const xml = await emitAndRead(oneImageDeck(pngDataUri(4, 2), 'contain'))
    const rect = srcRect(xml)
    expect(rect, 'a srcRect inset was emitted').toBeTruthy()
    if (!rect) return
    // the full width is shown (no horizontal crop)...
    expect(rect.l).toBe(0)
    expect(rect.r).toBe(0)
    // ...and the image is inset vertically (negative = padding, not crop)
    expect(rect.t).toBeLessThan(0)
    expect(rect.b).toBeLessThan(0)
  })

  it('cover falls back to a plain stretch when the image is undecodable', async () => {
    const xml = await emitAndRead(oneImageDeck('data:image/svg+xml;base64,PHN2Zy8+', 'cover'))
    expect(srcRect(xml)).toBeNull()
    expect(xml).toContain('<a:stretch><a:fillRect/></a:stretch>')
  })
})
