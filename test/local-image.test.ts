import { describe, it, expect } from 'vitest'
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { convert } from '../src/convert.js'
import { slideXml } from './helpers/pptx-xml.js'

describe('local relative images', () => {
  it('resolves a relative image path and embeds it without crashing', async () => {
    const inPath = new URL('./fixtures/local-image.md', import.meta.url).pathname
    const out = join(mkdtempSync(join(tmpdir(), 'mnp-li-')), 'local-image.pptx')
    await convert(inPath, out)
    const xml = (await slideXml(out))[0]
    expect(xml).toContain('<p:pic>')
  })
})
