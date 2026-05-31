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
import { themeList, type ThemePreset } from './themes.js'

export const LAYOUT_ORDER: LayoutKind[] = [
  'title',
  'section',
  'content',
  'two-column',
  'grid',
  'metrics',
  'timeline',
  'steps',
  'quote',
]

// Build the block content for one layout, badged with `badge` (e.g. the theme
// name + layout) so each slide is self-identifying in a review.
function content(layout: LayoutKind, badge: string): Slide {
  switch (layout) {
    case 'title':
      return {
        layout,
        blocks: [
          { type: 'kicker', text: badge },
          { type: 'heading', level: 1, text: 'Meet Pixel the Robot' },
          { type: 'paragraph', text: 'Your friendly companion for the connected home.' },
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
      const feat = (label: string, desc: string): Block => ({ type: 'feature', label, desc })
      return {
        layout,
        blocks: [
          { type: 'heading', level: 1, text: 'Built-in skills' },
          { type: 'paragraph', text: badge },
          feat('Voice', 'Natural back-and-forth conversation.'),
          feat('Vision', 'Recognizes faces and gestures.'),
          feat('Routines', 'Automates your mornings.'),
          feat('Security', 'Watches the door while you are out.'),
          feat('Music', 'Room-filling adaptive sound.'),
          feat('Updates', 'Gets smarter every week.'),
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
  }
}

const badgeFor = (t: ThemePreset, layout: LayoutKind): string => `${t.name} · ${t.category} · ${layout}`

/**
 * Curated review deck: each theme shown with a rotating window of layouts.
 * Across the 12 presets this covers every layout and every theme, in ~36
 * slides — compact enough to scroll through and audit by eye.
 */
export function buildGallery(perTheme = 3): SlideLayout[] {
  const out: SlideLayout[] = []
  themeList().forEach((t, ti) => {
    const layouts = Array.from(
      { length: perTheme },
      (_, k) => LAYOUT_ORDER[(ti * perTheme + k) % LAYOUT_ORDER.length],
    )
    out.push(...layoutDeck({ theme: t, slides: layouts.map((l) => content(l, badgeFor(t, l))) }))
  })
  return out
}

/** Exhaustive layout × theme matrix (every combination), grouped by theme. */
export function buildMatrix(): SlideLayout[] {
  const out: SlideLayout[] = []
  themeList().forEach((t) => {
    out.push(...layoutDeck({ theme: t, slides: LAYOUT_ORDER.map((l) => content(l, badgeFor(t, l))) }))
  })
  return out
}
