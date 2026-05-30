import { describe, it, expect } from 'vitest'
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { convert } from '../src/convert.js'
import { slideXml } from './helpers/pptx-xml.js'

describe('table integration', () => {
  it('emits a native table with cell text from a deck', async () => {
    const inPath = new URL('./fixtures/table.md', import.meta.url).pathname
    const out = join(mkdtempSync(join(tmpdir(), 'mnp-tbl-')), 'table.pptx')
    await convert(inPath, out)
    const xml = (await slideXml(out))[0]
    expect(xml).toContain('<a:tbl>')
    expect(xml).toContain('Character')
    expect(xml).toContain('Pixel the Robot')
    expect(xml).toContain('Pixel Park')
  })
})
