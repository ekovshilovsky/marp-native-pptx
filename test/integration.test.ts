import { describe, it, expect } from 'vitest'
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { convert } from '../src/convert.js'
import { slideXml } from './helpers/pptx-xml.js'

describe('integration: no fragmentation', () => {
  it('renders an inline code chip as ONE run with a fill inside ONE paragraph', async () => {
    const inPath = new URL('./fixtures/chip.md', import.meta.url).pathname
    const out = join(mkdtempSync(join(tmpdir(), 'mnp-')), 'chip.pptx')
    await convert(inPath, out)
    const xml = (await slideXml(out))[0]

    // The body paragraph that mentions the domain lives in exactly one shape.
    const shapesWithDomain = (xml.match(/cartoons\.example\.com/g) || []).length
    expect(shapesWithDomain).toBe(1) // not split across multiple frames

    // The chip carries a highlight fill (a:highlight) on its run.
    expect(xml).toMatch(/<a:highlight>[\s\S]*?f3e8ff/i)

    // The domain text and surrounding text share one paragraph (one <a:p> around them):
    // there is no extra paragraph break introduced inside the sentence.
    const sentenceShape = xml.split('<p:sp>').find((s) => s.includes('cartoons.example.com'))!
    const paraCount = (sentenceShape.match(/<a:p>/g) || []).length
    expect(paraCount).toBe(1)
  })

  it('converts a deck with an inline image without crashing and emits a picture', async () => {
    const inPath = new URL('./fixtures/image.md', import.meta.url).pathname
    const out = join(mkdtempSync(join(tmpdir(), 'mnp-img-')), 'image.pptx')
    await convert(inPath, out)
    const xml = (await slideXml(out))[0]
    expect(xml).toContain('<p:pic>')
  })
})
