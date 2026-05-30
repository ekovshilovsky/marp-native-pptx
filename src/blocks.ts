// Block IR — a semantic, layout-agnostic deck model.
//
// This is the "structured IR" layer: a deck is slides of typed blocks (kicker,
// heading, paragraph, bullets, image), with a per-slide *layout template* that
// decides WHERE each block goes. Templates place blocks in fixed, aligned
// regions (title anchored at a consistent top, content in a grid) — the
// "everything lines up" look of polished decks, instead of content floating
// vertically centered.
//
// The engine already separates concerns: layoutDeck() produces the SAME
// geometric `SlideLayout` the browser path produces, so it reuses `mapToPptx`
// -> `emit` -> `validatePptx` unchanged. The hard part (PowerPoint-clean OOXML)
// is already solved; this file is just a layout engine on top of it.
import type { LayoutBox, Paragraph, Run, SlideLayout } from './types.js'

export interface InlineRun {
  text: string
  bold?: boolean
  italic?: boolean
  code?: boolean
  color?: string // 6-hex without '#'
}
export type Inline = string | InlineRun[]

export type Block =
  | { type: 'kicker'; text: string } // small eyebrow label above a title
  | { type: 'heading'; level?: 1 | 2 | 3; text: Inline }
  | { type: 'paragraph'; text: Inline }
  | { type: 'bullets'; items: Inline[] }
  | { type: 'image'; src: string }

export type LayoutKind = 'title' | 'content' | 'two-column'

export interface Slide {
  layout: LayoutKind
  blocks?: Block[] // used by 'title' and 'content'
  columns?: Block[][] // used by 'two-column' (each column is a block list)
}

export interface Theme {
  accent: string // hex (reserved for future shape primitives)
  ink: string // body text hex
  muted: string // kicker / secondary hex
  headingFont: string
  bodyFont: string
}

export interface Deck {
  slides: Slide[]
  theme?: Partial<Theme>
}

const DEFAULT_THEME: Theme = {
  accent: '34d399',
  ink: '1f2933',
  muted: '8a94a6',
  headingFont: 'Arial',
  bodyFont: 'Arial',
}

// Slide canvas in CSS px (matches the engine's 16:9 viewport).
const W = 1280
const H = 720
const M = 80 // consistent outer margin
const TITLE_TOP = 70 // FIXED title top — every content slide's title aligns here

// pt sizes per role
const SIZE = { kicker: 14, h1: 40, h2: 27, body: 18 }

const toRuns = (inline: Inline, theme: Theme, sizePt: number, defaultColor: string): Run[] => {
  const arr: InlineRun[] = typeof inline === 'string' ? [{ text: inline }] : inline
  return arr.map((r) => ({
    text: r.text,
    style: {
      fontFace: r.code ? 'Consolas' : theme.bodyFont,
      sizePt,
      bold: !!r.bold,
      italic: !!r.italic,
      color: r.color ?? defaultColor,
      fill: r.code ? 'f3e8ff' : undefined,
    },
  }))
}

const paragraph = (
  runs: Run[],
  align: Paragraph['align'],
  lineSpacingPt: number,
  bullet?: boolean,
): Paragraph => ({
  runs,
  align,
  lineSpacingPt,
  bullet: bullet ? { type: 'bullet', indentLevel: 0 } : undefined,
})

// rough text-height estimate (px) for stacking content blocks
const estimateHeight = (text: string, sizePt: number, widthPx: number): number => {
  const fontPx = sizePt / 0.75
  const charsPerLine = Math.max(1, Math.floor(widthPx / (fontPx * 0.5)))
  const lines = Math.max(1, Math.ceil(text.length / charsPerLine))
  return Math.ceil(lines * fontPx * 1.3)
}
const inlineText = (inline: Inline): string =>
  typeof inline === 'string' ? inline : inline.map((r) => r.text).join('')

// Render a vertical stack of blocks into boxes within [x, y0, w] and return them.
function stack(blocks: Block[], x: number, y0: number, w: number, theme: Theme): LayoutBox[] {
  const boxes: LayoutBox[] = []
  let y = y0
  for (const b of blocks) {
    if (b.type === 'heading') {
      const pt = b.level === 1 ? SIZE.h1 : SIZE.h2
      const h = estimateHeight(inlineText(b.text), pt, w) + 8
      boxes.push({
        kind: 'text',
        rect: { xPx: x, yPx: y, wPx: w, hPx: h },
        valign: 'top',
        paras: [paragraph(toRuns(b.text, theme, pt, theme.ink).map((r) => ({ ...r, style: { ...r.style, bold: true } })), 'left', pt * 0.75 * 1.15)],
      })
      y += h + 14
    } else if (b.type === 'kicker') {
      const h = estimateHeight(b.text, SIZE.kicker, w) + 4
      boxes.push({
        kind: 'text',
        rect: { xPx: x, yPx: y, wPx: w, hPx: h },
        valign: 'top',
        paras: [paragraph(toRuns(b.text.toUpperCase(), theme, SIZE.kicker, theme.muted).map((r) => ({ ...r, style: { ...r.style, bold: true } })), 'left', SIZE.kicker * 0.75 * 1.2)],
      })
      y += h + 10
    } else if (b.type === 'paragraph') {
      const h = estimateHeight(inlineText(b.text), SIZE.body, w) + 6
      boxes.push({
        kind: 'text',
        rect: { xPx: x, yPx: y, wPx: w, hPx: h },
        valign: 'top',
        paras: [paragraph(toRuns(b.text, theme, SIZE.body, theme.ink), 'left', SIZE.body * 0.75 * 1.3)],
      })
      y += h + 12
    } else if (b.type === 'bullets') {
      const paras = b.items.map((it) =>
        paragraph(toRuns(it, theme, SIZE.body, theme.ink), 'left', SIZE.body * 0.75 * 1.3, true),
      )
      const h = b.items.reduce((acc, it) => acc + estimateHeight(inlineText(it), SIZE.body, w) + 14, 8)
      boxes.push({ kind: 'text', rect: { xPx: x, yPx: y, wPx: w, hPx: h }, valign: 'top', paras })
      y += h + 12
    } else if (b.type === 'image') {
      const h = 260
      boxes.push({ kind: 'image', rect: { xPx: x, yPx: y, wPx: w, hPx: h }, src: b.src })
      y += h + 12
    }
  }
  return boxes
}

/** Lay a block-IR deck out into the engine's geometric SlideLayout[]. */
export function layoutDeck(deck: Deck): SlideLayout[] {
  const theme: Theme = { ...DEFAULT_THEME, ...(deck.theme ?? {}) }
  return deck.slides.map((slide): SlideLayout => {
    const boxes: LayoutBox[] = []
    if (slide.layout === 'title') {
      // Left-anchored title block (kicker / big title / description), like a
      // polished cover slide — all left-aligned at a fixed left margin.
      boxes.push(...stack(slide.blocks ?? [], M, 230, W - 2 * M, theme))
    } else if (slide.layout === 'two-column') {
      const cols = slide.columns ?? []
      const colW = (W - 2 * M - 60) / 2
      // Optional leading heading block shared across the top.
      let top = TITLE_TOP
      const lead = (slide.blocks ?? []).filter((b) => b.type === 'heading' || b.type === 'kicker')
      if (lead.length) {
        boxes.push(...stack(lead, M, TITLE_TOP, W - 2 * M, theme))
        top = TITLE_TOP + 120
      }
      cols.forEach((col, i) => {
        boxes.push(...stack(col, M + i * (colW + 60), top, colW, theme))
      })
    } else {
      // 'content': title anchored at a CONSISTENT top, content in a region below.
      boxes.push(...stack(slide.blocks ?? [], M, TITLE_TOP, W - 2 * M, theme))
    }
    return { sizePx: { w: W, h: H }, boxes }
  })
}
