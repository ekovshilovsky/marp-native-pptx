import { describe, it, expect } from 'vitest'
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { convert } from '../src/convert.js'
import { slideXml } from './helpers/pptx-xml.js'

const fixtureDir = new URL('./fixtures/', import.meta.url).pathname

describe('gradient slide background', () => {
  it('rasterizes a gradient section background as a slide image, keeping text native', async () => {
    const out = join(mkdtempSync(join(tmpdir(), 'mnp-bgg-')), 'bgg.pptx')
    await convert(join(fixtureDir, 'bg-gradient.md'), out)
    const xml = (await slideXml(out))[0]
    expect(xml).toMatch(/<p:bg>[\s\S]*?<a:blip/i)
    expect(xml).toContain('Gradient title')
  })
})
