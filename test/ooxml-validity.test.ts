import { describe, it, expect } from 'vitest'
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { convert } from '../src/convert.js'
import { slideXml } from './helpers/pptx-xml.js'

describe('ooxml validity (schema-shape regressions)', () => {
  it('<= one <a:pPr> per <a:p>, and every table cell has a paragraph', async () => {
    const inPath = new URL('./fixtures/mixed.md', import.meta.url).pathname
    const out = join(mkdtempSync(join(tmpdir(), 'mnp-val-')), 'mixed.pptx')
    await convert(inPath, out)
    for (const xml of await slideXml(out)) {
      for (const m of xml.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)) {
        expect((m[1].match(/<a:pPr\b/g) || []).length).toBeLessThanOrEqual(1)
      }
      for (const m of xml.matchAll(/<a:tc>([\s\S]*?)<\/a:tc>/g)) {
        expect(/<a:p[>\s/]/.test(m[1])).toBe(true)
      }
    }
  })
})
