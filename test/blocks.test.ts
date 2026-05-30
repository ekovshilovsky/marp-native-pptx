import { describe, it, expect } from 'vitest'
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { blocksToPptx, layoutDeck } from '../src/blocks-convert.js'
import { validatePptx } from '../src/validate.js'
import type { Deck } from '../src/blocks.js'

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
})
