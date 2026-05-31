// Vector art for the showcase: a small monoline icon set + a cartoon robot,
// baked to PNG data URIs so they embed as ordinary images in any PowerPoint.
//
// These live with the showcase (not the core library) — the library just
// places whatever image `src` it is handed. Baking happens once, in the
// example/test, via the Chromium rasterizer.
import { svgsToPngDataUris } from './rasterize.js'
import { coverSvg, tileSvg } from './showcase-art.js'
import { themeList } from './themes.js'

const wrap = (inner: string, stroke: string): string =>
  `<svg viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg" fill="none" stroke="#${stroke}" ` +
  `stroke-width="4.2" stroke-linecap="round" stroke-linejoin="round">${inner}</svg>`

// 6 icons keyed by name. Each is monoline so it reads at any size/color.
const ICONS: Record<string, string> = {
  mic: '<rect x="25" y="11" width="14" height="27" rx="7"/><path d="M18 31a14 14 0 0 0 28 0"/><line x1="32" y1="45" x2="32" y2="52"/><line x1="23" y1="52" x2="41" y2="52"/>',
  eye: '<path d="M7 32C17 18 47 18 57 32 47 46 17 46 7 32Z"/><circle cx="32" cy="32" r="7"/>',
  clock: '<circle cx="32" cy="32" r="20"/><path d="M32 21v11l8 5"/>',
  shield: '<path d="M32 9l19 8v13c0 13-9 19-19 23-10-4-19-10-19-23V17z"/><path d="M24 32l6 6 12-13"/>',
  music: '<path d="M26 45V19l21-5v25"/><circle cx="20.5" cy="45.5" r="6"/><circle cx="41.5" cy="39.5" r="6"/>',
  refresh: '<path d="M50 24a20 20 0 1 0 4 14"/><path d="M50 10v14H36"/>',
}

export const ICON_KEYS = Object.keys(ICONS)

/** A friendly cartoon robot, line-art in one ink color (transparent fill). */
function robotSvg(stroke: string): string {
  const s = `#${stroke}`
  return (
    `<svg viewBox="0 0 200 200" xmlns="http://www.w3.org/2000/svg" fill="none" stroke="${s}" ` +
    `stroke-width="5" stroke-linecap="round" stroke-linejoin="round">` +
    `<line x1="100" y1="40" x2="100" y2="24"/><circle cx="100" cy="18" r="6" fill="${s}"/>` +
    `<rect x="54" y="40" width="92" height="68" rx="18"/>` +
    `<circle cx="82" cy="72" r="8" fill="${s}"/><circle cx="118" cy="72" r="8" fill="${s}"/>` +
    `<path d="M82 90q18 13 36 0"/>` +
    `<rect x="68" y="118" width="64" height="50" rx="12"/>` +
    `<line x1="68" y1="132" x2="48" y2="150"/><line x1="132" y1="132" x2="152" y2="150"/>` +
    `<circle cx="100" cy="143" r="8" fill="${s}"/>` +
    `<line x1="86" y1="168" x2="86" y2="180"/><line x1="114" y1="168" x2="114" y2="180"/>` +
    `</svg>`
  )
}

export interface ShowcaseAssets {
  // icons baked in a dark ink (for light chips) and a light ink (for dark chips)
  icons: { dark: Record<string, string>; light: Record<string, string> }
  robot: { dark: string; light: string }
  // per-theme generated imagery: a 16:9 cover (with scrim) + a square photo tile
  art: Record<string, { cover: string; tile: string }>
}

// ink colors used for the two contrast variants
const DARK_INK = '2b2622'
const LIGHT_INK = 'f2f2f0'

/** Rasterize the whole asset set once (single browser session). */
export async function bakeShowcaseAssets(): Promise<ShowcaseAssets> {
  const keys = ICON_KEYS
  const darkSvgs = keys.map((k) => wrap(ICONS[k], DARK_INK))
  const lightSvgs = keys.map((k) => wrap(ICONS[k], LIGHT_INK))
  // rasterize icons (64px) and robots (200px) in two same-size batches
  const iconUris = await svgsToPngDataUris([...darkSvgs, ...lightSvgs], { width: 64, height: 64 })
  const robotUris = await svgsToPngDataUris([robotSvg(DARK_INK), robotSvg(LIGHT_INK)], { width: 200, height: 200 })
  const dark: Record<string, string> = {}
  const light: Record<string, string> = {}
  keys.forEach((k, i) => {
    dark[k] = iconUris[i]
    light[k] = iconUris[keys.length + i]
  })
  // per-theme cover (16:9) + tile (square), gradients from each theme's palette
  const themes = themeList()
  const coverUris = await svgsToPngDataUris(themes.map((t) => coverSvg(t.palette)), { width: 1600, height: 900 })
  const tileUris = await svgsToPngDataUris(themes.map((t) => tileSvg(t.palette, 0)), { width: 560, height: 560 })
  const art: Record<string, { cover: string; tile: string }> = {}
  themes.forEach((t, i) => {
    art[t.id] = { cover: coverUris[i], tile: tileUris[i] }
  })
  return { icons: { dark, light }, robot: { dark: robotUris[0], light: robotUris[1] }, art }
}
