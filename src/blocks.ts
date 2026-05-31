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
  | { type: 'feature'; color?: string; label: Inline; desc?: Inline } // icon-dot + label + description (also a numbered step)
  | { type: 'event'; date: string; text: Inline } // a timeline event
  | { type: 'stat'; value: string; label: Inline } // a big-number metric
  | { type: 'quote'; text: Inline; cite?: string } // a pull quote + attribution

export type LayoutKind =
  | 'title'
  | 'content'
  | 'two-column'
  | 'grid'
  | 'timeline'
  | 'section'
  | 'metrics'
  | 'quote'
  | 'steps'

export interface Slide {
  layout: LayoutKind
  blocks?: Block[] // used by 'title' and 'content'
  columns?: Block[][] // used by 'two-column' (each column is a block list)
}

export interface Theme {
  accent: string // primary brand hex (rules, nodes, numbers)
  ink: string // body/heading text hex
  muted: string // kicker / secondary text hex
  bg: string // slide background hex
  surface: string // card / panel fill hex
  palette: string[] // accent palette cycled for feature dots / nodes / stats
  headingFont: string
  bodyFont: string
}

export interface Deck {
  slides: Slide[]
  theme?: Partial<Theme>
}

// The built-in fallback — the house signature "Vega": electric indigo-violet
// on cool paper. The full curated, tagged preset set lives in ./themes.ts;
// this is the safe default that any partial deck.theme is merged onto. (Kept
// in sync with THEMES.vega there.)
export const DEFAULT_THEME: Theme = {
  accent: '5346e3',
  ink: '1a1a2e',
  muted: '7e7e99',
  bg: 'f7f7fb',
  surface: 'ececf5',
  palette: ['5346e3', '2d9cdb', '12b5a8', 'f0654e', '9b5de5', 'f4a52e'],
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

// --- color helpers: keep generated text legible on any theme color ---
const hexToRgb = (h: string): [number, number, number] => {
  const s = h.replace('#', '')
  return [parseInt(s.slice(0, 2), 16), parseInt(s.slice(2, 4), 16), parseInt(s.slice(4, 6), 16)]
}
const toHex = (n: number): string => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0')
// relative luminance (perceptual weighting) → pick dark or light ink for contrast
const readableOn = (bgHex: string): string => {
  const [r, g, b] = hexToRgb(bgHex)
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255 > 0.6 ? '1f2933' : 'ffffff'
}
// linear blend of two hex colors, t in [0,1] toward `b`
const mix = (a: string, b: string, t: number): string => {
  const [ar, ag, ab] = hexToRgb(a)
  const [br, bg, bb] = hexToRgb(b)
  return toHex(ar + (br - ar) * t) + toHex(ag + (bg - ag) * t) + toHex(ab + (bb - ab) * t)
}

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
    let bgFill = theme.bg // section overrides this with a full-bleed accent
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
      const palette = theme.palette
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
    } else if (slide.layout === 'section') {
      // Full-bleed divider: accent background, big left-anchored number + title
      // + optional one-line description, all in inverted (background-colored) text.
      const blocks = slide.blocks ?? []
      bgFill = theme.accent
      const onAccent = readableOn(theme.accent)
      const dim = mix(theme.accent, onAccent, 0.78) // a muted version of the inverted ink
      let y = 250
      const kicker = blocks.find((b) => b.type === 'kicker') as Extract<Block, { type: 'kicker' }> | undefined
      if (kicker) {
        boxes.push({
          kind: 'text',
          rect: { xPx: M, yPx: y, wPx: W - 2 * M, hPx: 40 },
          valign: 'top',
          paras: [paragraph(toRuns(kicker.text.toUpperCase(), theme, 20, dim).map((r) => ({ ...r, style: { ...r.style, bold: true } })), 'left', 20 * 0.75 * 1.2)],
        })
        y += 52
      }
      const heading = blocks.find((b) => b.type === 'heading') as Extract<Block, { type: 'heading' }> | undefined
      if (heading) {
        const pt = fitFontPt(inlineText(heading.text), W - 2 * M, 150, 54, 30)
        const h = estimateHeight(inlineText(heading.text), pt, W - 2 * M) + 8
        boxes.push({
          kind: 'text',
          rect: { xPx: M, yPx: y, wPx: W - 2 * M, hPx: h },
          valign: 'top',
          paras: [paragraph(toRuns(heading.text, theme, pt, onAccent).map((r) => ({ ...r, style: { ...r.style, bold: true } })), 'left', pt * 0.75 * 1.1)],
        })
        y += h + 14
      }
      const desc = blocks.find((b) => b.type === 'paragraph') as Extract<Block, { type: 'paragraph' }> | undefined
      if (desc) {
        boxes.push({
          kind: 'text',
          rect: { xPx: M, yPx: y, wPx: Math.min(760, W - 2 * M), hPx: 80 },
          valign: 'top',
          paras: [paragraph(toRuns(desc.text, theme, SIZE.body, dim), 'left', SIZE.body * 0.75 * 1.3)],
        })
      }
    } else if (slide.layout === 'metrics') {
      // Centered header + a row of big-number stat cards (number in accent,
      // label below) distributed evenly across the width.
      const blocks = slide.blocks ?? []
      const head = centeredHeader(blocks.filter((b) => b.type === 'heading' || b.type === 'paragraph'), theme, 60)
      boxes.push(...head.boxes)
      const stats = blocks.filter((b): b is Extract<Block, { type: 'stat' }> => b.type === 'stat')
      const n = stats.length
      if (n) {
        const top = Math.max(head.endY + 50, 280)
        const cardH = 200
        const cells = gridCells(box(M, top, W - 2 * M, cardH), 1, n, 36, 0)
        stats.forEach((st, i) => {
          const cell = cells[i]
          boxes.push({ kind: 'shape', preset: 'roundRect', radiusPx: 18, rect: { xPx: cell.x, yPx: cell.y, wPx: cell.w, hPx: cardH }, fill: theme.surface })
          const numPt = fitFontPt(st.value, cell.w - 24, 90, 60, 28)
          boxes.push({
            kind: 'text',
            rect: { xPx: cell.x + 12, yPx: cell.y + 34, wPx: cell.w - 24, hPx: 90 },
            valign: 'top',
            paras: [paragraph(toRuns(st.value, theme, numPt, theme.palette[i % theme.palette.length]).map((r) => ({ ...r, style: { ...r.style, bold: true } })), 'center', numPt * 0.75 * 1.05)],
          })
          const labPt = fitFontPt(inlineText(st.label), cell.w - 24, 56, SIZE.body, 11)
          boxes.push({
            kind: 'text',
            rect: { xPx: cell.x + 12, yPx: cell.y + cardH - 70, wPx: cell.w - 24, hPx: 56 },
            valign: 'top',
            paras: [paragraph(toRuns(st.label, theme, labPt, theme.muted), 'center', labPt * 0.75 * 1.25)],
          })
        })
      }
    } else if (slide.layout === 'quote') {
      // Big centered pull quote with an accent mark, plus attribution.
      const blocks = slide.blocks ?? []
      const q = blocks.find((b): b is Extract<Block, { type: 'quote' }> => b.type === 'quote')
      if (q) {
        boxes.push({ kind: 'shape', rect: { xPx: W / 2 - 36, yPx: 150, wPx: 72, hPx: 8 }, fill: theme.accent })
        const qpt = fitFontPt(inlineText(q.text), W - 2 * M - 120, 300, 34, 18)
        const qh = estimateHeight(inlineText(q.text), qpt, W - 2 * M - 120) + 20
        boxes.push({
          kind: 'text',
          rect: { xPx: M + 60, yPx: 210, wPx: W - 2 * M - 120, hPx: qh },
          valign: 'top',
          paras: [paragraph(toRuns(q.text, theme, qpt, theme.ink).map((r) => ({ ...r, style: { ...r.style, italic: true } })), 'center', qpt * 0.75 * 1.3)],
        })
        if (q.cite) {
          boxes.push({
            kind: 'text',
            rect: { xPx: M + 60, yPx: 210 + qh + 18, wPx: W - 2 * M - 120, hPx: 36 },
            valign: 'top',
            paras: [paragraph(toRuns(`— ${q.cite}`, theme, SIZE.body, theme.muted).map((r) => ({ ...r, style: { ...r.style, bold: true } })), 'center', SIZE.body * 0.75 * 1.2)],
          })
        }
      }
    } else if (slide.layout === 'steps') {
      // Numbered horizontal process: a centered header, then evenly-distributed
      // numbered nodes with a connecting rail, each with a title + fit-to-width desc.
      const blocks = slide.blocks ?? []
      const head = centeredHeader(blocks.filter((b) => b.type === 'heading' || b.type === 'paragraph'), theme, 56)
      boxes.push(...head.boxes)
      const steps = blocks.filter((b): b is Extract<Block, { type: 'feature' }> => b.type === 'feature')
      const n = steps.length
      if (n) {
        const railY = Math.max(head.endY + 70, 320)
        const pad = 110
        const rail = box(M + pad, railY, W - 2 * M - 2 * pad, 0)
        const xs = distributeX(rail, n)
        if (n > 1) boxes.push({ kind: 'shape', rect: { xPx: rail.x, yPx: railY - 1, wPx: rail.w, hPx: 3 }, fill: theme.muted })
        const node = 56
        const colW = Math.min(240, (W - 2 * M) / n - 10)
        const onAccent = readableOn(theme.palette[0])
        steps.forEach((st, i) => {
          const cx = xs[i]
          const c = theme.palette[i % theme.palette.length]
          boxes.push({ kind: 'shape', preset: 'ellipse', rect: { xPx: cx - node / 2, yPx: railY - node / 2, wPx: node, hPx: node }, fill: c })
          boxes.push({
            kind: 'text',
            rect: { xPx: cx - node / 2, yPx: railY - node / 2, wPx: node, hPx: node },
            valign: 'middle',
            paras: [paragraph(toRuns(String(i + 1), theme, 24, readableOn(c)).map((r) => ({ ...r, style: { ...r.style, bold: true } })), 'center', 24 * 0.75)],
          })
          const lx = cx - colW / 2
          const ly = railY + node / 2 + 18
          boxes.push({
            kind: 'text',
            rect: { xPx: lx, yPx: ly, wPx: colW, hPx: 30 },
            valign: 'top',
            paras: [paragraph(toRuns(st.label, theme, 17, theme.ink).map((r) => ({ ...r, style: { ...r.style, bold: true } })), 'center', 17 * 0.75 * 1.1)],
          })
          if (st.desc) {
            const dpt = fitFontPt(inlineText(st.desc), colW, 110, 14, 9)
            boxes.push({
              kind: 'text',
              rect: { xPx: lx, yPx: ly + 32, wPx: colW, hPx: 110 },
              valign: 'top',
              paras: [paragraph(toRuns(st.desc, theme, dpt, theme.muted), 'center', dpt * 0.75 * 1.25)],
            })
          }
        })
        void onAccent
      }
    } else {
      // 'content': title anchored at a CONSISTENT top, content in a region below.
      boxes.push(...stack(slide.blocks ?? [], M, TITLE_TOP, W - 2 * M, theme))
    }
    return { sizePx: { w: W, h: H }, boxes, background: bgFill ? { fill: bgFill } : undefined }
  })
}
