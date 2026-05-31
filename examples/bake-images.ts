// Bake real per-theme showcase imagery with nano banana (Gemini 2.5 Flash
// Image), then emit a showcase deck that runs those photos through the ordinary
// image + cover-fit path — proving the layouts support real image content.
//
//   npx tsx examples/bake-images.ts --one            # just Vega (2 images) — cheap pipeline check
//   npx tsx examples/bake-images.ts                  # all themes (cover + tile each)
//   npx tsx examples/bake-images.ts --themes vega,plum
//
// The key is read from .env.local (GEMINI_API_KEY=...), which is gitignored.
import { execFileSync } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { generateImage } from '../src/gemini-image.js'
import { scrimWrapSvg } from '../src/showcase-art.js'
import { svgsToPngDataUris } from '../src/rasterize.js'
import { bakeShowcaseAssets, type ShowcaseAssets } from '../src/showcase-assets.js'
import { buildGallery } from '../src/showcase-deck.js'
import { mapToPptx } from '../src/map.js'
import { emit } from '../src/emit.js'
import { validatePptx } from '../src/validate.js'
import { themeList, type ThemePreset } from '../src/themes.js'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const GEN_DIR = join(ROOT, 'examples', 'generated')
const OUT = join(ROOT, 'showcase-real.pptx')

// Pull GEMINI_API_KEY from .env.local (falling back to the process env).
function loadKey(): string {
  const fromEnv = process.env.GEMINI_API_KEY
  if (fromEnv) return fromEnv
  const envPath = join(ROOT, '.env.local')
  if (existsSync(envPath)) {
    for (const line of readFileSync(envPath, 'utf8').split('\n')) {
      const m = line.match(/^\s*GEMINI_API_KEY\s*=\s*(.+?)\s*$/)
      if (m && m[1]) return m[1].replace(/^["']|["']$/g, '')
    }
  }
  return ''
}

// Mood adjectives drawn from the theme's category so each image matches the
// preset's personality, tinted by its accent color for a cohesive set.
const MOOD: Record<string, string> = {
  Light: 'airy, bright, clean, sunlit',
  Dark: 'moody, low-key, dramatic, cinematic night',
  Minimal: 'calm, restrained, generous negative space',
  Bold: 'vivid, high-energy, saturated, confident',
  Professional: 'refined, trustworthy, understated, corporate',
  Playful: 'friendly, warm, rounded, optimistic',
  Editorial: 'sophisticated, magazine-quality, art-directed',
}

function coverPrompt(t: ThemePreset): string {
  const mood = MOOD[t.category] ?? 'modern, elegant'
  return (
    `Premium abstract brand photograph for a friendly smart-home assistant. ` +
    `${mood}. Dominant accent color #${t.accent}. Soft cinematic lighting, shallow depth of field, ` +
    `minimal and elegant, lots of room on the left for overlaid text. No text, no logos, no watermarks, no people. Wide 16:9.`
  )
}

function tilePrompt(t: ThemePreset): string {
  const mood = MOOD[t.category] ?? 'modern, elegant'
  return (
    `Clean product-style vignette evoking a calm connected home. ${mood}. ` +
    `Accent color #${t.accent}. Studio lighting, minimal background, single subject, centered. ` +
    `No text, no logos, no watermarks, no people. Square 1:1.`
  )
}

async function mapLimited<T, R>(items: T[], limit: number, fn: (item: T, i: number) => Promise<R>): Promise<R[]> {
  const out: R[] = new Array(items.length)
  let next = 0
  await Promise.all(
    Array.from({ length: Math.min(limit, items.length) }, async () => {
      while (next < items.length) {
        const i = next++
        out[i] = await fn(items[i], i)
      }
    }),
  )
  return out
}

async function main(): Promise<void> {
  const args = process.argv.slice(2)
  const rescrim = args.includes('--rescrim') // re-bake the scrim from saved raw — no API calls, free
  const only = args.includes('--one')
  const themesArg = args.find((a) => a.startsWith('--themes='))?.split('=')[1]
  let themes = themeList()
  if (only) themes = themes.slice(0, 1)
  else if (themesArg) themes = themes.filter((t) => themesArg.split(',').includes(t.id))

  mkdirSync(GEN_DIR, { recursive: true })
  // Icons + robot stay vector-baked; only the per-theme photo art is generated.
  const base = await bakeShowcaseAssets()
  const art: ShowcaseAssets['art'] = { ...base.art }
  const dataUri = (b: Buffer) => `data:image/png;base64,${b.toString('base64')}`

  // 1) obtain raw cover + tile per theme — either from the API or from disk.
  let raw: { id: string; cover: Buffer; tile: Buffer }[]
  if (rescrim) {
    console.log('Re-baking scrim from saved raw images (no API calls)…')
    raw = themeList()
      .filter((t) => existsSync(join(GEN_DIR, `cover-raw-${t.id}.png`)))
      .map((t) => ({ id: t.id, cover: readFileSync(join(GEN_DIR, `cover-raw-${t.id}.png`)), tile: readFileSync(join(GEN_DIR, `tile-${t.id}.png`)) }))
    if (raw.length === 0) {
      console.error('No raw covers found in examples/generated/. Run a full bake first.')
      process.exit(1)
    }
  } else {
    const key = loadKey()
    if (!key) {
      console.error('No GEMINI_API_KEY found. Paste a key into .env.local (see the file) and re-run.')
      process.exit(1)
    }
    console.log(`Generating ${themes.length * 2} images (${themes.length} themes × cover+tile)…`)
    raw = await mapLimited(themes, 3, async (t) => {
      const [cover, tile] = await Promise.all([
        generateImage(coverPrompt(t), { apiKey: key, aspectRatio: '16:9' }),
        generateImage(tilePrompt(t), { apiKey: key, aspectRatio: '1:1' }),
      ])
      console.log(`  ✓ ${t.name}`)
      return { id: t.id, cover, tile }
    })
    // keep the un-scrimmed originals so the scrim can be re-tuned for free
    raw.forEach((r) => writeFileSync(join(GEN_DIR, `cover-raw-${r.id}.png`), r.cover))
  }

  // 2) bake the legibility scrim onto every cover in one browser session
  const scrimmed = await svgsToPngDataUris(
    raw.map((r) => scrimWrapSvg(dataUri(r.cover))),
    { width: 1600, height: 900, scale: 1 },
  )

  // 3) write files (scrimmed cover) + assemble the art map
  raw.forEach((r, i) => {
    writeFileSync(join(GEN_DIR, `cover-${r.id}.png`), Buffer.from(scrimmed[i].split(',')[1], 'base64'))
    writeFileSync(join(GEN_DIR, `tile-${r.id}.png`), r.tile)
    art[r.id] = { cover: scrimmed[i], tile: dataUri(r.tile) }
  })

  const assets: ShowcaseAssets = { ...base, art }
  const slides = buildGallery(assets)
  await emit(mapToPptx(slides, 13.333, 7.5), OUT)
  const findings = await validatePptx(OUT)
  console.log(`\nimages saved to: ${GEN_DIR}`)
  console.log(`deck:            ${OUT}`)
  console.log(`validator:       ${findings.length === 0 ? 'PowerPoint-clean (0 findings)' : JSON.stringify(findings, null, 2)}`)

  // best-effort PNG preview — render into the gitignored generated/ dir so the
  // repo root stays clean.
  try {
    execFileSync('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', GEN_DIR, OUT], { stdio: 'ignore' })
    execFileSync('pdftoppm', ['-png', '-r', '110', '-f', '1', '-l', '6', join(GEN_DIR, 'showcase-real.pdf'), join(GEN_DIR, 'preview')], { stdio: 'ignore' })
    console.log(`preview pages:   ${join(GEN_DIR, 'preview-*.png')}`)
  } catch {
    /* libreoffice/poppler optional */
  }
}

void main()
