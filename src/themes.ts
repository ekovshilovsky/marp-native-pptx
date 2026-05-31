// Curated theme presets — the library's own visual identity.
//
// These palettes are original to this project, not borrowed from any slide
// tool. The house signature is "Vega": an electric indigo-violet on cool
// paper — crisp, modern, and a little bold, with a designed (not rainbow)
// supporting palette. Around it sits a deliberately broad set so the library
// never feels like one brand: cool and warm, light and dark, minimal and
// editorial. The point is range with taste, not a single mood.
//
// Each preset is a full `Theme` (colors + fonts) plus discovery metadata: a
// stable `id`, a human `name`, a primary `category`, and free-form `tags`.
// The categorization/tagging is deliberate — it lets a UI (or an AI picking a
// theme) filter by intent ("a dark keynote theme", "something playful") rather
// than guessing from raw hex. New presets only need adding here: every layout
// template renders from the `Theme` fields, so a new theme works everywhere.
import type { Theme } from './blocks.js'

// Broad intent buckets. A theme may read as more than one; `category` is its
// primary bucket and `tags` capture the rest.
export type ThemeCategory =
  | 'Light'
  | 'Dark'
  | 'Minimal'
  | 'Bold'
  | 'Professional'
  | 'Playful'
  | 'Editorial'

export interface ThemePreset extends Theme {
  id: string
  name: string
  category: ThemeCategory
  tags: string[]
}

const def = (p: ThemePreset): ThemePreset => p

// 12 presets spanning the categories. Colors are 6-hex (no '#'). Palettes are
// hand-tuned to harmonize within each theme (a warm lead + supporting hues).
export const THEMES: Record<string, ThemePreset> = {
  vega: def({
    id: 'vega',
    name: 'Vega',
    category: 'Bold',
    tags: ['signature', 'default', 'modern', 'indigo', 'crisp'],
    accent: '5346e3', // electric indigo-violet — the house signature
    ink: '1a1a2e',
    muted: '7e7e99',
    bg: 'f7f7fb', // cool near-white
    surface: 'ececf5',
    palette: ['5346e3', '2d9cdb', '12b5a8', 'f0654e', '9b5de5', 'f4a52e'],
    headingFont: 'Arial',
    bodyFont: 'Arial',
  }),
  clay: def({
    id: 'clay',
    name: 'Clay',
    category: 'Editorial',
    tags: ['warm', 'elegant', 'editorial', 'terracotta'],
    accent: 'c2603f', // warm terracotta
    ink: '2b2622',
    muted: '8c8178',
    bg: 'f7f3ee', // warm paper
    surface: 'efe7dd',
    palette: ['c2603f', 'd99873', '6b8f71', 'c9a13b', '7d6f9c', '4f7a86'],
    headingFont: 'Arial',
    bodyFont: 'Arial',
  }),
  linen: def({
    id: 'linen',
    name: 'Linen',
    category: 'Minimal',
    tags: ['minimal', 'neutral', 'calm', 'soft'],
    accent: '5b6472', // soft slate
    ink: '23282f',
    muted: '8a8f98',
    bg: 'faf8f5',
    surface: 'eeebe5',
    palette: ['5b6472', '8a8f98', 'b08d57', '6f8a86', 'a3a8b0', 'b5654a'],
    headingFont: 'Arial',
    bodyFont: 'Arial',
  }),
  harbor: def({
    id: 'harbor',
    name: 'Harbor',
    category: 'Professional',
    tags: ['professional', 'calm', 'teal', 'business'],
    accent: '2f6f7e', // deep teal
    ink: '17242b',
    muted: '6a7c84',
    bg: 'ffffff',
    surface: 'eef3f4',
    palette: ['2f6f7e', '3f7d8c', 'c08a4e', '5b6f8a', '87a4ad', 'b5654a'],
    headingFont: 'Arial',
    bodyFont: 'Arial',
  }),
  cobalt: def({
    id: 'cobalt',
    name: 'Cobalt',
    category: 'Bold',
    tags: ['bold', 'professional', 'blue', 'confident'],
    accent: '2d4ea0', // rich cobalt
    ink: '121a2b',
    muted: '5f6b82',
    bg: 'ffffff',
    surface: 'eef1f8',
    palette: ['2d4ea0', '3f6fb5', 'c08a4e', '5b8a8f', '7a6fb0', 'b5654a'],
    headingFont: 'Arial',
    bodyFont: 'Arial',
  }),
  sage: def({
    id: 'sage',
    name: 'Sage',
    category: 'Professional',
    tags: ['nature', 'calm', 'organic', 'green'],
    accent: '5f7a5b', // soft forest
    ink: '222b22',
    muted: '7d8a78',
    bg: 'f5f7f1',
    surface: 'e6ece0',
    palette: ['5f7a5b', '7b9476', 'b08d57', '4f6f6a', '9caa6e', 'c2703f'],
    headingFont: 'Arial',
    bodyFont: 'Arial',
  }),
  mist: def({
    id: 'mist',
    name: 'Mist',
    category: 'Light',
    tags: ['airy', 'calm', 'cool', 'data'],
    accent: '40697a', // deep slate-teal — reads clearly on the near-white bg
    ink: '24313a',
    muted: '79868f',
    bg: 'fbfcfd',
    surface: 'e9eef1',
    palette: ['40697a', '5f8ba0', 'a8894f', '5d6f8a', '4f8a76', 'b06a6a'],
    headingFont: 'Arial',
    bodyFont: 'Arial',
  }),
  marigold: def({
    id: 'marigold',
    name: 'Marigold',
    category: 'Bold',
    tags: ['warm', 'marketing', 'energetic', 'gold'],
    accent: 'd98324', // marigold
    ink: '2b231a',
    muted: '9a8a76',
    bg: 'fffaf2',
    surface: 'fbeed7',
    palette: ['d98324', 'c2603f', 'c9a13b', '6b8f71', 'b5654a', '8a6d4f'],
    headingFont: 'Arial',
    bodyFont: 'Arial',
  }),
  plum: def({
    id: 'plum',
    name: 'Plum',
    category: 'Editorial',
    tags: ['editorial', 'elegant', 'wine', 'refined'],
    accent: '7a3b5e', // wine
    ink: '241a22',
    muted: '8b7d85',
    bg: 'faf6f7',
    surface: 'f0e6ea',
    palette: ['7a3b5e', 'a85d7d', 'b07a3f', '5e6f8a', 'c98fa8', '6b4a6b'],
    headingFont: 'Arial',
    bodyFont: 'Arial',
  }),
  paper: def({
    id: 'paper',
    name: 'Paper',
    category: 'Editorial',
    tags: ['editorial', 'classic', 'print', 'serif', 'cream'],
    accent: '9a3b2e', // brick red
    ink: '1c1916',
    muted: '746c64',
    bg: 'f6f1e7', // cream
    surface: 'efe7d8',
    palette: ['9a3b2e', '1c1916', '9c7a3c', '6b6f5c', '7d5a4f', '8a6d4f'],
    headingFont: 'Georgia',
    bodyFont: 'Georgia',
  }),
  dusk: def({
    id: 'dusk',
    name: 'Dusk',
    category: 'Dark',
    tags: ['dark', 'elegant', 'keynote', 'gold'],
    accent: 'e0a458', // warm gold
    ink: 'ede7df',
    muted: '9aa0ad',
    bg: '1c1f29', // deep blue-charcoal
    surface: '262b38',
    palette: ['e0a458', 'd97757', '7fa8c9', '8fbf9f', 'c98fb0', 'd8c27d'],
    headingFont: 'Arial',
    bodyFont: 'Arial',
  }),
  ember: def({
    id: 'ember',
    name: 'Ember',
    category: 'Bold',
    tags: ['dark', 'bold', 'warm', 'ember'],
    accent: 'd9774e', // ember orange
    ink: 'f0e8e0',
    muted: 'a89c92',
    bg: '1f1a17', // warm charcoal
    surface: '2a2420',
    palette: ['d9774e', 'e0a458', 'c2603f', '8a9a6b', 'b08fae', '6fa1a8'],
    headingFont: 'Arial',
    bodyFont: 'Arial',
  }),
  obsidian: def({
    id: 'obsidian',
    name: 'Obsidian',
    category: 'Minimal',
    tags: ['dark', 'minimal', 'mono', 'stone'],
    accent: 'c8b8a1', // warm stone
    ink: 'f2f2f0',
    muted: '9b9b97',
    bg: '121212',
    surface: '1e1e1e',
    palette: ['c8b8a1', 'a89f93', 'd9cbb8', '8a8278', 'bcae9b', '766f64'],
    headingFont: 'Arial',
    bodyFont: 'Arial',
  }),
}

/** All presets as an array (registration order; Vega is first / the default). */
export const themeList = (): ThemePreset[] => Object.values(THEMES)

/** Look up a preset by id; falls back to the default (Vega) if unknown. */
export const getTheme = (id: string): ThemePreset => THEMES[id] ?? THEMES.vega

/** Presets in a given category. */
export const themesByCategory = (c: ThemeCategory): ThemePreset[] =>
  themeList().filter((t) => t.category === c)

/** Presets carrying a given tag (case-insensitive). */
export const themesByTag = (tag: string): ThemePreset[] => {
  const t = tag.toLowerCase()
  return themeList().filter((p) => p.tags.some((x) => x.toLowerCase() === t))
}

/** Every distinct tag across all presets, sorted. */
export const allTags = (): string[] =>
  [...new Set(themeList().flatMap((t) => t.tags))].sort()
