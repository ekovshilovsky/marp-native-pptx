import { describe, expect, it } from 'vitest'
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { mapToPptx } from '../src/map.js'
import { emit } from '../src/emit.js'
import { validatePptx } from '../src/validate.js'
import { buildMatrix, LAYOUT_ORDER } from '../src/showcase-deck.js'
import { blocksToPptx } from '../src/blocks-convert.js'
import { themeList } from '../src/themes.js'
import type { ShowcaseAssets } from '../src/showcase-assets.js'

// a 1x1 transparent PNG — stands in for baked icons/illustration without a browser
const PNG_1x1 =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQAY3Y2wAAAAAElFTkSuQmCC'

describe('showcase: every layout under every theme', () => {
  it('renders the full layout x theme matrix PowerPoint-clean, with zero rasterized text', async () => {
    const slides = buildMatrix() // no assets: pure native, browser-free
    expect(slides.length).toBe(themeList().length * LAYOUT_ORDER.length)
    expect(slides.every((s) => s.boxes.length > 0)).toBe(true)

    const model = mapToPptx(slides, 13.333, 7.5)
    // text/shapes are never rasterized: an asset-free deck has zero images
    const images = model.slides.flatMap((s) => s.shapes).filter((sh) => sh.kind === 'image')
    expect(images.length).toBe(0)

    const out = join(mkdtempSync(join(tmpdir(), 'mnp-showcase-')), 'showcase.pptx')
    await emit(model, out)
    expect(await validatePptx(out)).toEqual([])
  })

  it('embeds icon + hero images and still validates PowerPoint-clean', async () => {
    const fakeIcons = Object.fromEntries(['mic', 'eye', 'clock', 'shield', 'music', 'refresh'].map((k) => [k, PNG_1x1]))
    const assets: ShowcaseAssets = {
      icons: { dark: fakeIcons, light: fakeIcons },
      robot: { dark: PNG_1x1, light: PNG_1x1 },
      art: { x: { cover: PNG_1x1, tile: PNG_1x1 } },
    }
    const out = join(mkdtempSync(join(tmpdir(), 'mnp-showcase-img-')), 'imgs.pptx')
    await blocksToPptx(
      {
        slides: [
          { layout: 'title', blocks: [{ type: 'heading', level: 1, text: 'Hero' }, { type: 'image', src: assets.robot.dark }] },
          {
            layout: 'image-feature',
            blocks: [{ type: 'heading', level: 1, text: 'Feature' }, { type: 'bullets', items: ['a', 'b'] }, { type: 'image', src: assets.robot.dark }],
          },
          {
            layout: 'grid',
            blocks: [
              { type: 'heading', level: 1, text: 'Skills' },
              { type: 'feature', label: 'Voice', desc: 'talks', icon: assets.icons.dark.mic },
              { type: 'feature', label: 'Vision', desc: 'sees', icon: assets.icons.dark.eye },
            ],
          },
        ],
      },
      out,
    )
    expect(await validatePptx(out)).toEqual([]) // images embedded, still clean
  })
})
