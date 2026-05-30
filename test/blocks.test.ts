import { describe, it, expect } from 'vitest'
import { mkdtempSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createRequire } from 'module'
import { blocksToPptx, layoutDeck } from '../src/blocks-convert.js'
import { validatePptx } from '../src/validate.js'
import type { Deck } from '../src/blocks.js'

const require = createRequire(import.meta.url)
// eslint-disable-next-line @typescript-eslint/no-explicit-any, @typescript-eslint/no-require-imports
const JSZip = require('jszip') as any

const deck: Deck = {
  slides: [
    {
      layout: 'title',
      blocks: [
        { type: 'kicker', text: 'Quarterly Review' },
        { type: 'heading', level: 1, text: 'Platform Evaluation' },
        { type: 'paragraph', text: 'A short description of the deck.' },
      ],
    },
    {
      layout: 'content',
      blocks: [
        { type: 'heading', level: 2, text: 'What we found' },
        { type: 'bullets', items: ['First finding', 'Second finding', 'Third finding'] },
      ],
    },
    {
      layout: 'two-column',
      blocks: [{ type: 'heading', level: 2, text: 'Context vs plan' }],
      columns: [
        [{ type: 'bullets', items: ['Market', 'Problem', 'Positioning'] }],
        [{ type: 'bullets', items: ['Strategy', 'Budget', 'Next steps'] }],
      ],
    },
  ],
}

describe('block IR -> pptx', () => {
  it('renders a block deck to a PowerPoint-clean pptx (reusing emit + validator)', async () => {
    const out = join(mkdtempSync(join(tmpdir(), 'mnp-blocks-')), 'deck.pptx')
    await blocksToPptx(deck, out)
    expect(await validatePptx(out)).toEqual([]) // 0 findings == PowerPoint-clean
  })

  it('anchors content-slide titles at a consistent top-left (aligned, not centered)', () => {
    const slides = layoutDeck(deck)
    const content = slides[1] // the 'content' slide
    const heading = content.boxes.find((b) => b.kind === 'text')
    expect(heading).toBeTruthy()
    if (!heading || heading.kind !== 'text') throw new Error('no heading box')
    expect(heading.rect.yPx).toBe(70) // TITLE_TOP — identical on every content slide
    expect(heading.rect.xPx).toBe(80) // left margin, not centered
  })

  it('emits a native shape (accent rule) with the theme accent fill', async () => {
    const out = join(mkdtempSync(join(tmpdir(), 'mnp-shape-')), 'deck.pptx')
    await blocksToPptx({ ...deck, theme: { accent: '34d399' } }, out)
    const zip = await JSZip.loadAsync(readFileSync(out))
    const xml: string = await zip.files['ppt/slides/slide1.xml'].async('string')
    expect(xml).toContain('<a:prstGeom prst="rect"') // a real shape, not an image
    expect(xml.toUpperCase()).toContain('34D399') // accent fill present
  })
})
