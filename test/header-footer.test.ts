import { describe, it, expect } from 'vitest'
import { render } from '../src/render.js'
import { layout } from '../src/layout.js'

const md = `---
marp: true
paginate: true
header: 'My Header'
footer: 'My Footer'
---

# One

text

---

# Two
`

describe('header/footer/pagination', () => {
  it('captures header, footer, and a page-number text box', async () => {
    const slides = await layout(await render(md, {}))
    const allText = (s: number) =>
      slides[s].boxes.filter((b) => b.kind === 'text').flatMap((b) =>
        b.kind === 'text' ? b.paras.flatMap((p) => p.runs.map((r) => r.text)) : [],
      ).join(' ')
    expect(allText(0)).toContain('My Header')
    expect(allText(0)).toContain('My Footer')
    expect(allText(1)).toMatch(/\b2\b/)
  })
})
