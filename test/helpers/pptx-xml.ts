import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'

const require = createRequire(import.meta.url)
// eslint-disable-next-line @typescript-eslint/no-require-imports, @typescript-eslint/no-explicit-any
const JSZip = require('jszip') as any

/** Return the XML string of every slide in a .pptx, in order. */
export async function slideXml(pptxPath: string): Promise<string[]> {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-call
  const zip = await JSZip.loadAsync(readFileSync(pptxPath))
  // eslint-disable-next-line @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access
  const names = (Object.keys(zip.files) as string[])
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)/)![1], 10)
      const nb = parseInt(b.match(/slide(\d+)/)![1], 10)
      return na - nb
    })
  // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-unsafe-return
  return Promise.all(names.map((n) => zip.files[n].async('string')))
}
