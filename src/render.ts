import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { Marp } from '@marp-team/marp-core'
import type { RenderOptions, RenderResult } from './types.js'

export async function render(markdown: string, opts: RenderOptions): Promise<RenderResult> {
  const marp = new Marp({ html: true })
  for (const dir of opts.themeSet ?? []) {
    for (const file of readdirSync(dir)) {
      if (file.endsWith('.css')) marp.themeSet.add(readFileSync(join(dir, file), 'utf8'))
    }
  }
  const { html, css } = marp.render(markdown)
  return { html, css }
}
