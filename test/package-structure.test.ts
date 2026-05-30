import { describe, it, expect } from 'vitest'
import { readFileSync, mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createRequire } from 'module'
import { convert } from '../src/convert.js'

const require = createRequire(import.meta.url)
const JSZip = require('jszip') as any

describe('OPC package structure', () => {
  it('[Content_Types].xml is first, no directory entries, no dangling overrides', async () => {
    const inPath = new URL('./fixtures/mixed.md', import.meta.url).pathname
    const out = join(mkdtempSync(join(tmpdir(), 'mnp-pkg-')), 'mixed.pptx')
    await convert(inPath, out)
    const zip = await JSZip.loadAsync(readFileSync(out))
    const names = Object.keys(zip.files) as string[]
    expect(names[0]).toBe('[Content_Types].xml')
    expect(names.filter((n) => n.endsWith('/'))).toHaveLength(0)
    const ct: string = await zip.files['[Content_Types].xml'].async('string')
    const existing = new Set(names.map((n) => '/' + n))
    const overrides = [...ct.matchAll(/<Override PartName="([^"]+)"/g)].map((m) => m[1])
    expect(overrides.filter((p) => !existing.has(p))).toEqual([])
  })

  it('contains no empty-text runs in any part (PowerPoint repairs those)', async () => {
    const inPath = new URL('./fixtures/mixed.md', import.meta.url).pathname
    const out = join(mkdtempSync(join(tmpdir(), 'mnp-empty-')), 'mixed.pptx')
    await convert(inPath, out)
    const zip = await JSZip.loadAsync(readFileSync(out))
    for (const name of Object.keys(zip.files) as string[]) {
      if (!name.endsWith('.xml')) continue
      const xml: string = await zip.files[name].async('string')
      expect(xml.includes('<a:t></a:t>')).toBe(false)
    }
  })

  it('notes master does not share the slide master theme (PowerPoint repairs that)', async () => {
    const inPath = new URL('./fixtures/mixed.md', import.meta.url).pathname
    const out = join(mkdtempSync(join(tmpdir(), 'mnp-theme-')), 'mixed.pptx')
    await convert(inPath, out)
    const zip = await JSZip.loadAsync(readFileSync(out))
    const themeOf = async (rels: string): Promise<string | undefined> => {
      const xml: string = await zip.files[rels].async('string')
      const m = xml.match(/Type="[^"]*\/theme"[^>]*Target="([^"]+)"/) || xml.match(/Target="([^"]+)"[^>]*Type="[^"]*\/theme"/)
      return m?.[1].split('/').pop()
    }
    const names = Object.keys(zip.files) as string[]
    const sm = names.find((n) => /slideMasters\/_rels\/.+\.rels$/.test(n))!
    const nm = names.find((n) => /notesMasters\/_rels\/.+\.rels$/.test(n))!
    const smTheme = await themeOf(sm)
    const nmTheme = await themeOf(nm)
    expect(smTheme).toBeTruthy()
    expect(nmTheme).toBeTruthy()
    expect(nmTheme).not.toBe(smTheme) // each master owns a distinct theme part
  })
})
