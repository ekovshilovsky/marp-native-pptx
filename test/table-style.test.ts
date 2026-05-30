import { describe, it, expect } from 'vitest'
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { convert } from '../src/convert.js'
import { slideXml } from './helpers/pptx-xml.js'

describe('table style id', () => {
  it('emits a table style id so PowerPoint does not flag the table', async () => {
    const inPath = new URL('./fixtures/table.md', import.meta.url).pathname
    const out = join(mkdtempSync(join(tmpdir(), 'mnp-ts-')), 'table.pptx')
    await convert(inPath, out)
    const xml = (await slideXml(out))[0]
    expect(xml).toMatch(/<a:tableStyleId>\{?[0-9A-Fa-f-]+\}?<\/a:tableStyleId>/)
  })
})
