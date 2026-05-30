import { describe, it, expect } from 'vitest'
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { markdownToDeck, markdownToPptx } from '../src/blocks-convert.js'
import { validatePptx } from '../src/validate.js'

const md = `---
marp: true
---

# Platform Evaluation

How the product performed this quarter.

---

## What we found

- Usage grew 40%
- Onboarding is the drop-off
- Enterprise drives revenue

---

<!-- layout: two-column -->

## Context

- Market
- Problem

## What we'll do

- Strategy
- Budget
`

describe('markdown -> blocks', () => {
  it('parses front-matter, slide breaks, headings, bullets, and layout directives', () => {
    const deck = markdownToDeck(md)
    expect(deck.slides.length).toBe(3)
    expect(deck.slides[0].layout).toBe('title') // first-slide heuristic
    expect(deck.slides[1].layout).toBe('content')
    const c2 = deck.slides[2]
    expect(c2.layout).toBe('two-column')
    expect(c2.columns?.length).toBe(2) // grouped at the two headings
    const h0 = deck.slides[0].blocks?.[0]
    expect(h0?.type).toBe('heading')
    const bullets = deck.slides[1].blocks?.find((b) => b.type === 'bullets')
    expect(bullets && bullets.type === 'bullets' && bullets.items.length).toBe(3)
  })

  it('renders markdown to a PowerPoint-clean pptx', async () => {
    const out = join(mkdtempSync(join(tmpdir(), 'mnp-md-')), 'md.pptx')
    await markdownToPptx(md, out)
    expect(await validatePptx(out)).toEqual([])
  })
})
