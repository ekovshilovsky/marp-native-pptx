import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { render } from '../src/render.js'
import { layout } from '../src/layout.js'

describe('layout', () => {
  it('extracts one text box per block element with geometry', async () => {
    const md = readFileSync(new URL('./fixtures/hello.md', import.meta.url), 'utf8')
    const r = await render(md, {})
    const slides = await layout(r)
    expect(slides.length).toBe(1)
    const texts = slides[0].boxes.filter((b) => b.kind === 'text')
    expect(texts.length).toBeGreaterThanOrEqual(2) // heading + paragraph
    expect(slides[0].sizePx.w).toBeGreaterThan(100)
    for (const b of texts) expect(b.rect.wPx).toBeGreaterThan(0)
  })
})

describe('layout bullets', () => {
  it('extracts one text box per list item', async () => {
    const md = readFileSync(new URL('./fixtures/bullets.md', import.meta.url), 'utf8')
    const slides = await layout(await render(md, {}))
    const items = slides[0].boxes.filter((b) => b.kind === 'text' && b.paras[0]?.bullet)
    expect(items.length).toBe(3)
  })
})
