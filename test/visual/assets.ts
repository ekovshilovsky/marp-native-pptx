// Image assets for the visual-regression fixture.
//
// Icons/illustration are baked once (Chromium) and committed as PNG files under
// test/visual/assets/, so the regression render is deterministic and needs no
// browser at test time — the fixture just points `image` blocks at file:// URLs.
import { mkdirSync, writeFileSync, existsSync } from 'node:fs'
import { join } from 'node:path'
import { bakeShowcaseAssets, ICON_KEYS, type ShowcaseAssets } from '../../src/showcase-assets.js'
import { themeList } from '../../src/themes.js'

const dataUriToBuffer = (uri: string): Buffer => Buffer.from(uri.replace(/^data:image\/png;base64,/, ''), 'base64')

/** Bake the asset set and write each PNG into `dir` (used when updating goldens). */
export async function writeBakedAssets(dir: string): Promise<void> {
  mkdirSync(dir, { recursive: true })
  const a = await bakeShowcaseAssets()
  for (const key of ICON_KEYS) {
    writeFileSync(join(dir, `icon-${key}-dark.png`), dataUriToBuffer(a.icons.dark[key]))
    writeFileSync(join(dir, `icon-${key}-light.png`), dataUriToBuffer(a.icons.light[key]))
  }
  writeFileSync(join(dir, 'robot-dark.png'), dataUriToBuffer(a.robot.dark))
  writeFileSync(join(dir, 'robot-light.png'), dataUriToBuffer(a.robot.light))
  for (const t of themeList()) {
    writeFileSync(join(dir, `cover-${t.id}.png`), dataUriToBuffer(a.art[t.id].cover))
    writeFileSync(join(dir, `tile-${t.id}.png`), dataUriToBuffer(a.art[t.id].tile))
  }
}

/** Build a ShowcaseAssets that points at the committed PNG files via file:// URLs. */
export function fileAssets(dir: string): ShowcaseAssets {
  const url = (name: string) => `file://${join(dir, name)}`
  const icons = (variant: 'dark' | 'light') =>
    Object.fromEntries(ICON_KEYS.map((k) => [k, url(`icon-${k}-${variant}.png`)]))
  const art = Object.fromEntries(themeList().map((t) => [t.id, { cover: url(`cover-${t.id}.png`), tile: url(`tile-${t.id}.png`) }]))
  return {
    icons: { dark: icons('dark'), light: icons('light') },
    robot: { dark: url('robot-dark.png'), light: url('robot-light.png') },
    art,
  }
}

/** True if the committed asset PNGs are present. */
export function assetsExist(dir: string): boolean {
  return existsSync(join(dir, 'robot-dark.png'))
}
