import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { render } from '../src/render.js'
import { layout } from '../src/layout.js'

describe('image sizing', () => {
  it('measures an image at its real rendered size (not the unloaded default)', async () => {
    const p = new URL('./fixtures/sized-image.md', import.meta.url).pathname
    const md = readFileSync(p, 'utf8')
    const baseHref = pathToFileURL(resolve(dirname(p))).href + '/'
    const slides = await layout(await render(md, {}), baseHref)
    const img = slides[0].boxes.find((b) => b.kind === 'image')
    expect(img).toBeTruthy()
    if (!img || img.kind !== 'image') throw new Error('no image')
    expect(img.rect.wPx / img.rect.hPx).toBeGreaterThan(2)
  })
})
