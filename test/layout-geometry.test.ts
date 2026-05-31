import { describe, expect, it } from 'vitest'
import { layoutDeck, type Deck } from '../src/blocks.js'
import { buildMatrix } from '../src/showcase-deck.js'
import type { LayoutBox, SlideLayout } from '../src/types.js'

// The engine canvas (CSS px). Geometry is the deterministic source of truth
// that map+emit turn into the .pptx, so asserting on it is a pixel-precise
// contract: if these hold, the rendered slide is laid out correctly.
const W = 1280
const H = 720
const EPS = 2

const center = (b: LayoutBox) => ({ x: b.rect.xPx + b.rect.wPx / 2, y: b.rect.yPx + b.rect.hPx / 2 })
const vOverlap = (a: LayoutBox, b: LayoutBox) =>
  Math.max(a.rect.yPx, b.rect.yPx) < Math.min(a.rect.yPx + a.rect.hPx, b.rect.yPx + b.rect.hPx) - EPS

describe('layout geometry contract', () => {
  it('keeps every box on-canvas with valid dimensions, across all layouts x themes', () => {
    const slides: SlideLayout[] = buildMatrix()
    for (const s of slides) {
      for (const b of s.boxes) {
        expect(b.rect.wPx, 'positive width').toBeGreaterThan(0)
        expect(b.rect.hPx, 'positive height').toBeGreaterThan(0)
        expect(b.rect.xPx).toBeGreaterThanOrEqual(-EPS)
        expect(b.rect.yPx).toBeGreaterThanOrEqual(-EPS)
        expect(b.rect.xPx + b.rect.wPx).toBeLessThanOrEqual(W + EPS)
        expect(b.rect.yPx + b.rect.hPx).toBeLessThanOrEqual(H + EPS)
      }
    }
  })

  it('centers each step number inside its node circle (concentric)', () => {
    const deck: Deck = {
      slides: [
        {
          layout: 'steps',
          blocks: [
            { type: 'heading', level: 1, text: 'Three steps' },
            { type: 'feature', label: 'One', desc: 'first' },
            { type: 'feature', label: 'Two', desc: 'second' },
            { type: 'feature', label: 'Three', desc: 'third' },
          ],
        },
      ],
    }
    const boxes = layoutDeck(deck)[0].boxes
    const circles = boxes.filter((b) => b.kind === 'shape' && (b as Extract<LayoutBox, { kind: 'shape' }>).preset === 'ellipse')
    expect(circles.length).toBe(3)
    // each digit text box must share a center with exactly one node circle
    const digits = boxes.filter((b) => b.kind === 'text' && b.paras[0]?.runs[0] && /^[0-9]+$/.test(b.paras[0].runs[0].text))
    expect(digits.length).toBe(3)
    for (const d of digits) {
      const dc = center(d)
      const match = circles.find((c) => Math.abs(center(c).x - dc.x) < 1 && Math.abs(center(c).y - dc.y) < 1)
      expect(match, `digit at ${dc.x},${dc.y} concentric with a node`).toBeDefined()
    }
  })

  it('never overlaps a centered title with its subtitle (the wrap-height bug)', () => {
    // a deliberately long title that wraps to two lines
    const deck: Deck = {
      slides: [
        {
          layout: 'grid',
          blocks: [
            { type: 'heading', level: 1, text: 'Key Components and Frameworks for Building LLM Agents' },
            { type: 'paragraph', text: 'Building blocks for advanced AI workflows' },
            { type: 'feature', label: 'A', desc: 'x' },
            { type: 'feature', label: 'B', desc: 'y' },
          ],
        },
      ],
    }
    const boxes = layoutDeck(deck)[0].boxes
    const texts = boxes.filter((b) => b.kind === 'text')
    const title = texts[0]
    const subtitle = texts[1]
    expect(vOverlap(title, subtitle)).toBe(false)
  })
})
