// Render a block-IR Deck to an editable .pptx, reusing the engine's existing
// map -> emit pipeline (and therefore the PowerPoint-clean output guarantees).
// No browser is involved: layoutDeck() produces the geometric SlideLayout
// directly, which mapToPptx + emit already know how to turn into a clean .pptx.
import { mapToPptx } from './map.js'
import { emit } from './emit.js'
import { layoutDeck, type Deck } from './blocks.js'
import { markdownToDeck } from './md-blocks.js'

export async function blocksToPptx(deck: Deck, outPath: string): Promise<void> {
  const slides = layoutDeck(deck)
  const model = mapToPptx(slides, 13.333, 7.5)
  await emit(model, outPath)
}

/** Markdown → block IR → editable .pptx (reuses the same clean emit pipeline). */
export async function markdownToPptx(md: string, outPath: string): Promise<void> {
  await blocksToPptx(markdownToDeck(md), outPath)
}

export { layoutDeck, DEFAULT_THEME } from './blocks.js'
export { markdownToDeck } from './md-blocks.js'
export type { Deck, Slide, Block, InlineRun, Theme, LayoutKind } from './blocks.js'
export {
  THEMES,
  themeList,
  getTheme,
  themesByCategory,
  themesByTag,
  allTags,
} from './themes.js'
export type { ThemePreset, ThemeCategory } from './themes.js'
