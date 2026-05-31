// Showcase deck builder — demonstrates that the engine generalizes: every
// layout template renders correctly under every theme preset, as fully native
// (editable, zero-image) PowerPoint, and validates PowerPoint-clean.
//
// Content is deliberately generic cartoon copy ("Pixel the Robot") so the deck
// is shareable and carries no real-world identifiers. Each slide is badged with
// its theme + layout so a reviewer can audit the matrix at a glance.
//
// Two builders share one content library:
//   buildGallery() — a curated, human-reviewable deck (each theme shown with a
//                    rotating set of layouts; covers all themes and all layouts)
//   buildMatrix()  — the exhaustive layout × theme matrix (for validation)
import type { SlideLayout } from './types.js'
import { layoutDeck, type Block, type LayoutKind, type Slide } from './blocks.js'
import { getTheme, themeList, type ThemePreset } from './themes.js'
import { ICON_KEYS, type ShowcaseAssets } from './showcase-assets.js'

export const LAYOUT_ORDER: LayoutKind[] = [
  'title',
  'section',
  'content',
  'two-column',
  'image-feature',
  'grid',
  'metrics',
  'timeline',
  'steps',
  'quote',
  'bento',
  'comparison',
  'chart',
  'stat-hero',
  'closing',
  'full-bleed',
]

// On a light theme the icon chips are light tints (use the dark icon variant);
// on a dark theme they are dark tints (use the light variant).
const variantFor = (t: ThemePreset): 'dark' | 'light' =>
  parseInt(t.bg.slice(0, 2), 16) * 0.299 + parseInt(t.bg.slice(2, 4), 16) * 0.587 + parseInt(t.bg.slice(4, 6), 16) * 0.114 > 150 ? 'dark' : 'light'

// Build the block content for one layout, badged with `badge` (e.g. the theme
// name + layout) so each slide is self-identifying in a review. When `assets`
// are supplied, image-bearing layouts get real icons / illustration.
function content(layout: LayoutKind, badge: string, t: ThemePreset, assets?: ShowcaseAssets): Slide {
  const v = variantFor(t)
  const robot = assets?.robot[v]
  const icon = (i: number): string | undefined => assets?.icons[v][ICON_KEYS[i % ICON_KEYS.length]]
  const cover = assets?.art[t.id]?.cover
  const tile = assets?.art[t.id]?.tile
  switch (layout) {
    case 'title':
      return {
        layout,
        blocks: [
          { type: 'kicker', text: badge },
          { type: 'heading', level: 1, text: 'Meet Pixel the Robot' },
          { type: 'paragraph', text: 'Your friendly companion for the connected home.' },
          ...(robot ? [{ type: 'image', src: robot } as Block] : []),
        ],
      }
    case 'image-feature':
      return {
        layout,
        blocks: [
          { type: 'kicker', text: badge },
          { type: 'heading', level: 1, text: 'Say hello to Pixel' },
          {
            type: 'bullets',
            items: ['Wakes on your voice', 'Learns your routines', 'Looks after the home'],
          },
          ...(tile ? [{ type: 'image', src: tile } as Block] : []),
        ],
      }
    case 'full-bleed':
      return {
        layout,
        blocks: [
          { type: 'kicker', text: badge },
          { type: 'heading', level: 1, text: 'A calmer, smarter home' },
          { type: 'paragraph', text: 'Pixel fades into the background until the moment you need it.' },
          ...(cover ? [{ type: 'image', src: cover } as Block] : []),
        ],
      }
    case 'section':
      return {
        layout,
        blocks: [
          { type: 'kicker', text: badge },
          { type: 'heading', level: 1, text: 'Why Pixel?' },
          { type: 'paragraph', text: 'A home assistant that actually understands you.' },
        ],
      }
    case 'content':
      return {
        layout,
        blocks: [
          { type: 'kicker', text: badge },
          { type: 'heading', level: 1, text: 'What Pixel does' },
          {
            type: 'bullets',
            items: [
              'Listens and responds in plain language',
              'Controls 200+ smart devices',
              'Learns your daily routines',
              'Keeps your data on the device',
            ],
          },
        ],
      }
    case 'two-column':
      return {
        layout,
        blocks: [
          { type: 'kicker', text: badge },
          { type: 'heading', level: 1, text: 'Pixel at a glance' },
        ],
        columns: [
          [
            { type: 'heading', level: 2, text: 'Hardware' },
            { type: 'bullets', items: ['Quad far-field mics', 'Ambient light ring', 'On-device chip'] },
          ],
          [
            { type: 'heading', level: 2, text: 'Software' },
            { type: 'bullets', items: ['Offline wake word', 'Routine engine', 'Weekly updates'] },
          ],
        ],
      }
    case 'grid': {
      const feat = (i: number, label: string, desc: string): Block => ({ type: 'feature', label, desc, icon: icon(i) })
      return {
        layout,
        blocks: [
          { type: 'heading', level: 1, text: 'Built-in skills' },
          { type: 'paragraph', text: badge },
          feat(0, 'Voice', 'Natural back-and-forth conversation.'),
          feat(1, 'Vision', 'Recognizes faces and gestures.'),
          feat(2, 'Routines', 'Automates your mornings.'),
          feat(3, 'Security', 'Watches the door while you are out.'),
          feat(4, 'Music', 'Room-filling adaptive sound.'),
          feat(5, 'Updates', 'Gets smarter every week.'),
        ],
      }
    }
    case 'metrics': {
      const stat = (value: string, label: string): Block => ({ type: 'stat', value, label })
      return {
        layout,
        blocks: [
          { type: 'heading', level: 1, text: 'Pixel by the numbers' },
          { type: 'paragraph', text: badge },
          stat('4.9★', 'Average rating'),
          stat('200+', 'Devices supported'),
          stat('50ms', 'Response time'),
          stat('1M+', 'Homes served'),
        ],
      }
    }
    case 'timeline': {
      const ev = (date: string, text: string): Block => ({ type: 'event', date, text })
      return {
        layout,
        blocks: [
          { type: 'heading', level: 1, text: 'Road to launch' },
          { type: 'paragraph', text: badge },
          ev('Jan', 'Concept sketches'),
          ev('Mar', 'First prototype'),
          ev('Jun', 'Closed beta'),
          ev('Sep', 'Public launch'),
          ev('Dec', 'Pixel v2'),
        ],
      }
    }
    case 'steps': {
      const step = (label: string, desc: string): Block => ({ type: 'feature', label, desc })
      return {
        layout,
        blocks: [
          { type: 'heading', level: 1, text: 'Set up in three steps' },
          { type: 'paragraph', text: badge },
          step('Unbox', 'Place Pixel anywhere with power.'),
          step('Connect', 'Join your home network in the app.'),
          step('Say hi', 'Pixel greets you and starts learning.'),
        ],
      }
    }
    case 'quote':
      return {
        layout,
        blocks: [
          {
            type: 'quote',
            text: 'Pixel feels less like a gadget and more like a member of the family.',
            cite: `Early beta tester · ${badge}`,
          },
        ],
      }
    case 'bento': {
      const feat = (i: number, label: string, desc: string): Block => ({ type: 'feature', label, desc, icon: icon(i) })
      return {
        layout,
        blocks: [
          { type: 'heading', level: 1, text: 'Why people love Pixel' },
          { type: 'paragraph', text: badge },
          feat(0, 'Always listening, never creepy', 'On-device wake word, so audio never leaves your home until you ask.'),
          feat(1, 'Sees', 'Recognizes faces and gestures.'),
          feat(2, 'Learns', 'Adapts to your routines.'),
          feat(3, 'Guards the home while you are away', 'Watches the door and pings your phone.'),
        ],
      }
    }
    case 'comparison':
      return {
        layout,
        blocks: [{ type: 'heading', level: 1, text: 'Pixel vs. the old way' }, { type: 'paragraph', text: badge }],
        columns: [
          [
            { type: 'heading', level: 2, text: 'Before Pixel' },
            { type: 'bullets', items: ['Five different apps', 'Cloud-only, always online', 'Forgets your preferences', 'Setup takes an afternoon'] },
          ],
          [
            { type: 'heading', level: 2, text: 'With Pixel' },
            { type: 'bullets', items: ['One voice for everything', 'Works offline, on-device', 'Learns and remembers', 'Set up in three minutes'] },
          ],
        ],
      }
    case 'chart': {
      const bar = (label: string, value: number): Block => ({ type: 'bar', label, value })
      return {
        layout,
        blocks: [
          { type: 'heading', level: 1, text: 'Homes served per quarter' },
          { type: 'paragraph', text: badge },
          bar('Q1', 120),
          bar('Q2', 280),
          bar('Q3', 540),
          bar('Q4', 910),
        ],
      }
    }
    case 'stat-hero':
      return {
        layout,
        blocks: [
          { type: 'kicker', text: badge },
          { type: 'stat', value: '1,000,000+', label: 'homes now wake up with Pixel' },
          { type: 'paragraph', text: 'From a single prototype to a million households in under a year.' },
        ],
      }
    case 'closing':
      return {
        layout,
        blocks: [
          { type: 'heading', level: 1, text: 'Say hi to Pixel' },
          { type: 'paragraph', text: badge },
          { type: 'bullets', items: ['hello@pixel.example', 'pixel.example/preorder'] },
        ],
      }
  }
}

const badgeFor = (t: ThemePreset, layout: LayoutKind): string => `${t.name} · ${t.category} · ${layout}`

/**
 * Curated review deck: each theme shown with a rotating window of layouts.
 * Across the 12 presets this covers every layout and every theme, in ~36
 * slides — compact enough to scroll through and audit by eye.
 */
export function buildGallery(assets?: ShowcaseAssets, perTheme = 4): SlideLayout[] {
  const out: SlideLayout[] = []
  themeList().forEach((t, ti) => {
    const layouts = Array.from(
      { length: perTheme },
      (_, k) => LAYOUT_ORDER[(ti * perTheme + k) % LAYOUT_ORDER.length],
    )
    out.push(...layoutDeck({ theme: t, slides: layouts.map((l) => content(l, badgeFor(t, l), t, assets)) }))
  })
  return out
}

/** All layouts (in LAYOUT_ORDER) rendered under a single theme. */
export function buildThemeSlides(themeId: string, assets?: ShowcaseAssets): SlideLayout[] {
  const t = getTheme(themeId)
  return layoutDeck({ theme: t, slides: LAYOUT_ORDER.map((l) => content(l, badgeFor(t, l), t, assets)) })
}

/** Exhaustive layout × theme matrix (every combination), grouped by theme. */
export function buildMatrix(assets?: ShowcaseAssets): SlideLayout[] {
  const out: SlideLayout[] = []
  themeList().forEach((t) => {
    out.push(...layoutDeck({ theme: t, slides: LAYOUT_ORDER.map((l) => content(l, badgeFor(t, l), t, assets)) }))
  })
  return out
}
