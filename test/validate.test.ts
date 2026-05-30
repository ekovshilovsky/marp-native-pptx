import { describe, it, expect, beforeAll } from 'vitest'
import { readFileSync, mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createRequire } from 'module'
import { convert } from '../src/convert.js'
import { validatePptx } from '../src/validate.js'

const require = createRequire(import.meta.url)
// eslint-disable-next-line @typescript-eslint/no-explicit-any, @typescript-eslint/no-require-imports
const JSZip = require('jszip') as any

let cleanBuf: Buffer

beforeAll(async () => {
  const inPath = new URL('./fixtures/mixed.md', import.meta.url).pathname
  const out = join(mkdtempSync(join(tmpdir(), 'mnp-validate-')), 'mixed.pptx')
  await convert(inPath, out)
  cleanBuf = readFileSync(out)
})

// Corrupt the clean deck in one specific way and return the bytes.
async function corrupt(mutate: (read: (n: string) => Promise<string>, put: (n: string, s: string) => void) => Promise<void>): Promise<Buffer> {
  const zip = await JSZip.loadAsync(cleanBuf)
  await mutate(
    (n) => zip.files[n].async('string'),
    (n, s) => zip.file(n, s),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

describe('validatePptx — passes clean output', () => {
  it('reports 0 findings for a freshly converted deck', async () => {
    expect(await validatePptx(cleanBuf)).toEqual([])
  })
})

describe('validatePptx — catches violations (proves the validator works)', () => {
  const rulesFor = async (buf: Buffer) => (await validatePptx(buf)).map((f) => f.rule)

  it('empty-text-run', async () => {
    const bad = await corrupt(async (read, put) => {
      const xml = await read('ppt/slides/slide1.xml')
      put('ppt/slides/slide1.xml', xml.replace('</a:p>', '<a:r><a:t></a:t></a:r></a:p>'))
    })
    expect(await rulesFor(bad)).toContain('empty-text-run')
  })

  it('ct-dangling-override', async () => {
    const bad = await corrupt(async (read, put) => {
      const ct = await read('[Content_Types].xml')
      put('[Content_Types].xml', ct.replace('</Types>', '<Override PartName="/ppt/slides/slideZ.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>'))
    })
    expect(await rulesFor(bad)).toContain('ct-dangling-override')
  })

  it('duplicate-shape-id', async () => {
    const bad = await corrupt(async (read, put) => {
      const xml = await read('ppt/slides/slide1.xml')
      put('ppt/slides/slide1.xml', xml.replace(/(<p:cNvPr id=")\d+(")/g, '$12$2'))
    })
    expect(await rulesFor(bad)).toContain('duplicate-shape-id')
  })

  it('shared-master-theme', async () => {
    const bad = await corrupt(async (read, put) => {
      const rels = await read('ppt/notesMasters/_rels/notesMaster1.xml.rels')
      put('ppt/notesMasters/_rels/notesMaster1.xml.rels', rels.replace(/theme\d+\.xml/, 'theme1.xml'))
    })
    expect(await rulesFor(bad)).toContain('shared-master-theme')
  })

  it('dangling-relationship', async () => {
    const bad = await corrupt(async (read, put) => {
      const rels = await read('ppt/slides/_rels/slide1.xml.rels')
      put('ppt/slides/_rels/slide1.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rIdX" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/missing.png"/></Relationships>'))
    })
    expect(await rulesFor(bad)).toContain('dangling-relationship')
  })

  it('multiple-pPr', async () => {
    const bad = await corrupt(async (read, put) => {
      const xml = await read('ppt/slides/slide1.xml')
      // inject a second pPr inside the first paragraph
      put('ppt/slides/slide1.xml', xml.replace(/(<a:p>)/, '$1<a:pPr/><a:pPr/>'))
    })
    expect(await rulesFor(bad)).toContain('multiple-pPr')
  })

  it('malformed-xml', async () => {
    const bad = await corrupt(async (read, put) => {
      const xml = await read('ppt/slides/slide1.xml')
      put('ppt/slides/slide1.xml', xml.replace('</p:sld>', '')) // drop root close tag
    })
    expect(await rulesFor(bad)).toContain('malformed-xml')
  })

  it('invalid-shape-preset', async () => {
    const bad = await corrupt(async (read, put) => {
      const xml = await read('ppt/slides/slide1.xml')
      put('ppt/slides/slide1.xml', xml.replace('prst="rect"', 'prst="oval"'))
    })
    expect(await rulesFor(bad)).toContain('invalid-shape-preset')
  })

  it('table-column-count-mismatch', async () => {
    const bad = await corrupt(async (read, put) => {
      const xml = await read('ppt/slides/slide1.xml')
      // add an extra grid column so rows no longer match the column count
      put('ppt/slides/slide1.xml', xml.replace('</a:tblGrid>', '<a:gridCol w="100"/></a:tblGrid>'))
    })
    expect(await rulesFor(bad)).toContain('table-column-count-mismatch')
  })

  it('table-graphicframe-namespace', async () => {
    const bad = await corrupt(async (read, put) => {
      const xml = await read('ppt/slides/slide1.xml')
      // p:graphicFrame -> a:graphicFrame (open + close); leaves XML well-formed
      put('ppt/slides/slide1.xml', xml.replaceAll('p:graphicFrame', 'a:graphicFrame'))
    })
    expect(await rulesFor(bad)).toContain('table-graphicframe-namespace')
  })

  it('missing-slide-size', async () => {
    const bad = await corrupt(async (read, put) => {
      const p = await read('ppt/presentation.xml')
      put('ppt/presentation.xml', p.replace(/<p:sldSz\b[^>]*\/>/, ''))
    })
    expect(await rulesFor(bad)).toContain('missing-slide-size')
  })

  it('missing-color-map', async () => {
    const bad = await corrupt(async (read, put) => {
      const m = await read('ppt/slideMasters/slideMaster1.xml')
      put('ppt/slideMasters/slideMaster1.xml', m.replace(/<p:clrMap\b[^>]*\/>/, ''))
    })
    expect(await rulesFor(bad)).toContain('missing-color-map')
  })
})
