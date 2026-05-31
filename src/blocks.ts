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
import { box, distributeX, fitFontPt, gridCells, textHeight } from './layout-engine.js'

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
  | { type: 'feature'; color?: string; label: Inline; desc?: Inline } // icon-dot + label + description
  | { type: 'event'; date: string; text: Inline } // a timeline event

export type LayoutKind = 'title' | 'content' | 'two-column' | 'grid' | 'timeline'

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

// text-height estimate (px) for stacking content blocks — delegates to the
// engine's measurement so the whole system shares one (pessimistic) model.
const estimateHeight = (text: string, sizePt: number, widthPx: number): number =>
  textHeight(text, sizePt, widthPx)
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
      y += h + 6
      // accent rule under the heading (consistent design element, like the
      // colored underline in polished templates)
      const ruleH = b.level === 1 ? 6 : 4
      boxes.push({ kind: 'shape', rect: { xPx: x, yPx: y, wPx: b.level === 1 ? 90 : 60, hPx: ruleH }, fill: theme.accent })
      y += ruleH + 16
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

// Centered title + subtitle at the top (shrink-to-fit title). Returns the boxes
// and the y where content below should begin. Used by grid/timeline templates.
function centeredHeader(blocks: Block[], theme: Theme, startY: number): { boxes: LayoutBox[]; endY: number } {
  const out: LayoutBox[] = []
  let y = startY
  for (const b of blocks) {
    if (b.type === 'heading') {
      const pt = fitFontPt(inlineText(b.text), W - 2 * M, 130, SIZE.h1, 24)
      const h = estimateHeight(inlineText(b.text), pt, W - 2 * M) + 8
      out.push({
        kind: 'text',
        rect: { xPx: M, yPx: y, wPx: W - 2 * M, hPx: h },
        valign: 'top',
        paras: [paragraph(toRuns(b.text, theme, pt, theme.ink).map((r) => ({ ...r, style: { ...r.style, bold: true } })), 'center', pt * 0.75 * 1.1)],
      })
      y += h + 12
    } else if (b.type === 'paragraph') {
      const h = estimateHeight(inlineText(b.text), SIZE.body, W - 2 * M) + 6
      out.push({
        kind: 'text',
        rect: { xPx: M, yPx: y, wPx: W - 2 * M, hPx: h },
        valign: 'top',
        paras: [paragraph(toRuns(b.text, theme, SIZE.body, theme.muted), 'center', SIZE.body * 0.75 * 1.3)],
      })
      y += h + 10
    }
  }
  return { boxes: out, endY: y }
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
    } else if (slide.layout === 'grid') {
      // Centered title + subtitle, then an icon/feature grid (dot + label + desc)
      // in UNIFORM cells (rows line up) with shrink-to-fit text (never overflows).
      const blocks = slide.blocks ?? []
      const head = centeredHeader(blocks.filter((b) => b.type === 'heading' || b.type === 'paragraph'), theme, 56)
      boxes.push(...head.boxes)
      const feats = blocks.filter((b): b is Extract<Block, { type: 'feature' }> => b.type === 'feature')
      const palette = ['34d399', '38bdf8', '818cf8', 'f472b6', 'fbbf24', '22d3ee']
      const cols = feats.length <= 4 ? 2 : 3
      const rows = Math.max(1, Math.ceil(feats.length / cols))
      const gridTop = Math.max(head.endY + 36, 230)
      const cells = gridCells(box(M, gridTop, W - 2 * M, H - gridTop - M), rows, cols, 50, 24)
      const dot = 46
      const labelRegion = 46 // fixed top region per cell so rows align
      feats.forEach((f, i) => {
        const cell = cells[i]
        boxes.push({ kind: 'shape', preset: 'ellipse', rect: { xPx: cell.x, yPx: cell.y, wPx: dot, hPx: dot }, fill: f.color ?? palette[i % palette.length] })
        const tx = cell.x + dot + 16
        const tw = cell.w - dot - 16
        const labelPt = fitFontPt(inlineText(f.label), tw, labelRegion, 18, 12)
        boxes.push({
          kind: 'text',
          rect: { xPx: tx, yPx: cell.y - 2, wPx: tw, hPx: labelRegion },
          valign: 'top',
          paras: [paragraph(toRuns(f.label, theme, labelPt, theme.ink).map((r) => ({ ...r, style: { ...r.style, bold: true } })), 'left', labelPt * 0.75 * 1.1)],
        })
        if (f.desc) {
          const dy = cell.y - 2 + labelRegion
          const dh = cell.y + cell.h - dy - 6
          const descPt = fitFontPt(inlineText(f.desc), tw, dh, 15, 10)
          boxes.push({
            kind: 'text',
            rect: { xPx: tx, yPx: dy, wPx: tw, hPx: dh },
            valign: 'top',
            paras: [paragraph(toRuns(f.desc, theme, descPt, theme.muted), 'left', descPt * 0.75 * 1.3)],
          })
        }
      })
    } else if (slide.layout === 'timeline') {
      // Horizontal timeline: a centered title, an axis, evenly-distributed event
      // nodes alternating above/below with stems and shrink-to-fit labels.
      const blocks = slide.blocks ?? []
      const head = centeredHeader(blocks.filter((b) => b.type === 'heading' || b.type === 'paragraph'), theme, 44)
      boxes.push(...head.boxes)
      const events = blocks.filter((b): b is Extract<Block, { type: 'event' }> => b.type === 'event')
      const n = events.length
      if (n) {
        const axisY = 372
        const pad = 90
        const track = box(M + pad, axisY, W - 2 * M - 2 * pad, 0)
        const xs = distributeX(track, n)
        boxes.push({ kind: 'shape', rect: { xPx: track.x, yPx: axisY - 2, wPx: track.w, hPx: 4 }, fill: theme.muted })
        const node = 18
        const stemH = 64
        const colW = Math.min(220, (W - 2 * M) / n)
        const labelH = 100
        events.forEach((e, i) => {
          const cx = xs[i]
          const above = i % 2 === 0
          boxes.push({ kind: 'shape', preset: 'ellipse', rect: { xPx: cx - node / 2, yPx: axisY - node / 2, wPx: node, hPx: node }, fill: theme.accent })
          boxes.push({ kind: 'shape', rect: { xPx: cx - 1, yPx: above ? axisY - stemH : axisY, wPx: 2, hPx: stemH }, fill: theme.muted })
          const lx = cx - colW / 2
          const ly = above ? axisY - stemH - labelH : axisY + stemH + 6
          boxes.push({
            kind: 'text',
            rect: { xPx: lx, yPx: ly, wPx: colW, hPx: 26 },
            valign: 'top',
            paras: [paragraph(toRuns(e.date, theme, 16, theme.ink).map((r) => ({ ...r, style: { ...r.style, bold: true } })), 'center', 16 * 0.75 * 1.1)],
          })
          const tpt = fitFontPt(inlineText(e.text), colW, labelH - 30, 14, 9)
          boxes.push({
            kind: 'text',
            rect: { xPx: lx, yPx: ly + 28, wPx: colW, hPx: labelH - 30 },
            valign: 'top',
            paras: [paragraph(toRuns(e.text, theme, tpt, theme.muted), 'center', tpt * 0.75 * 1.25)],
          })
        })
      }
    } else {
      // 'content': title anchored at a CONSISTENT top, content in a region below.
      boxes.push(...stack(slide.blocks ?? [], M, TITLE_TOP, W - 2 * M, theme))
    }
    return { sizePx: { w: W, h: H }, boxes }
  })
}
