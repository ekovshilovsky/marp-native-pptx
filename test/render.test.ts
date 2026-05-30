import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { render } from '../src/render.js'

describe('render', () => {
  it('returns html containing a section and css', async () => {
    const md = readFileSync(new URL('./fixtures/hello.md', import.meta.url), 'utf8')
    const { html, css } = await render(md, {})
    expect(html).toContain('<section')
    expect(html).toContain('Hello')
    expect(css.length).toBeGreaterThan(0)
  })
})
