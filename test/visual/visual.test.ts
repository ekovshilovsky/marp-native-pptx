import { beforeAll, describe, expect, it } from 'vitest'
import { execFileSync } from 'node:child_process'
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readdirSync } from 'node:fs'
import { readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import pixelmatch from 'pixelmatch'
import { PNG } from 'pngjs'
import { mapToPptx } from '../../src/map.js'
import { emit } from '../../src/emit.js'
import type { SlideLayout } from '../../src/types.js'
import { buildThemeSlides, LAYOUT_ORDER } from '../../src/showcase-deck.js'
import { assetsExist, fileAssets, writeBakedAssets } from './assets.js'

const HERE = dirname(fileURLToPath(import.meta.url))
const ASSET_DIR = join(HERE, 'assets')
const GOLDEN_DIR = join(HERE, 'golden')
const THEMES = ['vega', 'dusk'] // one light, one dark
const UPDATE = process.env.UPDATE_GOLDENS === '1'
// 96 DPI on a 13.333x7.5in slide => exactly 1280x720 px
const DPI = 96
// allow a tiny fraction of pixels to differ (font anti-aliasing noise)
const MAX_DIFF_RATIO = 0.004

// Emit `slides` to a .pptx and rasterize each page to PNG via LibreOffice.
async function renderSlidesToPngs(slides: SlideLayout[], outDir: string): Promise<string[]> {
  mkdirSync(outDir, { recursive: true })
  const pptx = join(outDir, 'deck.pptx')
  await emit(mapToPptx(slides, 13.333, 7.5), pptx)
  execFileSync('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', outDir, pptx], { stdio: 'ignore' })
  const pdf = join(outDir, 'deck.pdf')
  execFileSync('pdftoppm', ['-png', '-r', String(DPI), pdf, join(outDir, 'p')], { stdio: 'ignore' })
  return readdirSync(outDir)
    .filter((f) => f.startsWith('p') && f.endsWith('.png'))
    .sort()
    .map((f) => join(outDir, f))
}

function readPng(path: string): PNG {
  return PNG.sync.read(readFileSync(path))
}

const rendered: Record<string, string[]> = {}

beforeAll(async () => {
  if (UPDATE || !assetsExist(ASSET_DIR)) await writeBakedAssets(ASSET_DIR) // bakes icons/robot (Chromium)
  mkdirSync(GOLDEN_DIR, { recursive: true })
  const work = mkdtempSync(join(tmpdir(), 'mnp-visual-'))
  for (const themeId of THEMES) {
    const slides = buildThemeSlides(themeId, fileAssets(ASSET_DIR))
    rendered[themeId] = await renderSlidesToPngs(slides, join(work, themeId))
  }
}, 180000)

describe('visual regression: layouts render pixel-stable', () => {
  for (const themeId of THEMES) {
    for (let i = 0; i < LAYOUT_ORDER.length; i++) {
      const layout = LAYOUT_ORDER[i]
      it(`${themeId} / ${layout}`, () => {
        const actualPath = rendered[themeId][i]
        expect(actualPath, 'rendered page exists').toBeTruthy()
        const golden = join(GOLDEN_DIR, `${themeId}-${layout}.png`)

        if (UPDATE || !existsSync(golden)) {
          copyFileSync(actualPath, golden)
          return
        }

        const a = readPng(actualPath)
        const g = readPng(golden)
        expect(`${a.width}x${a.height}`, 'dimensions match golden').toBe(`${g.width}x${g.height}`)
        const diff = new PNG({ width: g.width, height: g.height })
        const n = pixelmatch(g.data, a.data, diff.data, g.width, g.height, { threshold: 0.1 })
        const ratio = n / (g.width * g.height)
        if (ratio > MAX_DIFF_RATIO) {
          const out = join(tmpdir(), `visual-diff-${themeId}-${layout}.png`)
          writeFileSync(out, PNG.sync.write(diff))
          throw new Error(`${themeId}/${layout}: ${n} px differ (${(ratio * 100).toFixed(2)}%) > ${(MAX_DIFF_RATIO * 100).toFixed(2)}%. Diff: ${out}`)
        }
      })
    }
  }
})
