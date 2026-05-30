import { describe, it, expect } from 'vitest'
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { convert } from '../src/convert.js'
import { slideXml } from './helpers/pptx-xml.js'

describe('unique shape ids', () => {
  it('emits no duplicate p:cNvPr ids within a slide mixing text, table, image', async () => {
    const inPath = new URL('./fixtures/mixed.md', import.meta.url).pathname
    const out = join(mkdtempSync(join(tmpdir(), 'mnp-ids-')), 'mixed.pptx')
    await convert(inPath, out)
    for (const xml of await slideXml(out)) {
      const ids = [...xml.matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => m[1])
      expect(new Set(ids).size).toBe(ids.length)
    }
  })
})
