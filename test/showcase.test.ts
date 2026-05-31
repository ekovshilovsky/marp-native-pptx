import { describe, expect, it } from 'vitest'
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { mapToPptx } from '../src/map.js'
import { emit } from '../src/emit.js'
import { validatePptx } from '../src/validate.js'
import { buildMatrix, LAYOUT_ORDER } from '../src/showcase-deck.js'
import { themeList } from '../src/themes.js'

describe('showcase: every layout under every theme', () => {
  it('renders the full layout x theme matrix PowerPoint-clean, with zero images', async () => {
    const slides = buildMatrix()
    // every theme x every layout
    expect(slides.length).toBe(themeList().length * LAYOUT_ORDER.length)
    // each slide produced real geometry (no empty/blank layouts)
    expect(slides.every((s) => s.boxes.length > 0)).toBe(true)

    const model = mapToPptx(slides, 13.333, 7.5)
    // fully native: not a single raster image anywhere in the deck
    const images = model.slides.flatMap((s) => s.shapes).filter((sh) => sh.kind === 'image')
    expect(images.length).toBe(0)

    const out = join(mkdtempSync(join(tmpdir(), 'mnp-showcase-')), 'showcase.pptx')
    await emit(model, out)
    expect(await validatePptx(out)).toEqual([]) // 0 findings == PowerPoint-clean
  })
})
