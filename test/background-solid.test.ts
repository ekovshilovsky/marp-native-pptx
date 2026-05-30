import { describe, it, expect } from 'vitest'
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { convert } from '../src/convert.js'
import { slideXml } from './helpers/pptx-xml.js'

const fixtureDir = new URL('./fixtures/', import.meta.url).pathname

describe('solid slide background', () => {
  it('sets a slide background color from a section background', async () => {
    const out = join(mkdtempSync(join(tmpdir(), 'mnp-bg-')), 'bg.pptx')
    await convert(join(fixtureDir, 'bg-solid.md'), out)
    const xml = (await slideXml(out))[0]
    expect(xml).toMatch(/<p:bg>[\s\S]*?<a:srgbClr val="123456"/i)
  })
})
