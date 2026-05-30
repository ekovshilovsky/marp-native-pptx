import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { render } from './render.js'
import { layout } from './layout.js'
import { mapToPptx } from './map.js'
import { emit } from './emit.js'
import type { RenderOptions } from './types.js'

export async function convert(inputPath: string, outPath: string, opts: RenderOptions = {}): Promise<void> {
  const md = readFileSync(inputPath, 'utf8')
  const rendered = await render(md, opts)
  const baseHref = pathToFileURL(resolve(dirname(inputPath))).href + '/'
  const slides = await layout(rendered, baseHref)
  const model = mapToPptx(slides, 13.333, 7.5)
  await emit(model, outPath)
}
