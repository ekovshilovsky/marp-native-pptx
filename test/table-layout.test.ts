import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { render } from '../src/render.js'
import { layout } from '../src/layout.js'

describe('table extraction', () => {
  it('extracts a table box with header + body rows and cell text', async () => {
    const md = readFileSync(new URL('./fixtures/table.md', import.meta.url), 'utf8')
    const slides = await layout(await render(md, {}))
    const tables = slides[0].boxes.filter((b) => b.kind === 'table')
    expect(tables.length).toBe(1)
    const t = tables[0]
    if (t.kind !== 'table') throw new Error('expected table')
    expect(t.rows.length).toBe(3)                 // 1 header + 2 body
    expect(t.rows[0][0].header).toBe(true)
    expect(t.rows[0][0].paras[0].runs.map((r) => r.text).join('')).toContain('Character')
    expect(t.rows[1][0].paras[0].runs.map((r) => r.text).join('')).toContain('Pixel the Robot')
    expect(t.colWidthsPx.length).toBe(2)
  })
})
