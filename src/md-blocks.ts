// Markdown → block-IR. A thin input adapter: parse (Marp-style) markdown with
// remark, walk the AST, and emit the semantic block model from blocks.ts. The
// block IR then renders through layout templates → map → emit → validate.
//
// Conventions:
//   - `---` (thematic break) separates slides
//   - leading YAML front-matter is stripped (Marp directives)
//   - `<!-- layout: title|content|two-column -->` sets a slide's template
//   - first slide defaults to `title` if it's only heading/paragraph
//   - in a `two-column` slide, blocks are grouped into columns at each heading
import { unified } from 'unified'
import remarkParse from 'remark-parse'
import type { Block, Deck, InlineRun, Slide } from './blocks.js'

/* eslint-disable @typescript-eslint/no-explicit-any */
function inlineRuns(node: any): InlineRun[] {
  const out: InlineRun[] = []
  const walk = (n: any, ctx: Partial<InlineRun>): void => {
    if (!n) return
    if (n.type === 'text') out.push({ text: n.value as string, ...ctx })
    else if (n.type === 'inlineCode') out.push({ text: n.value as string, code: true, ...ctx })
    else if (n.type === 'strong') (n.children ?? []).forEach((c: any) => walk(c, { ...ctx, bold: true }))
    else if (n.type === 'emphasis') (n.children ?? []).forEach((c: any) => walk(c, { ...ctx, italic: true }))
    else if (n.children) (n.children as any[]).forEach((c) => walk(c, ctx))
  }
  ;(node.children ?? []).forEach((c: any) => walk(c, {}))
  return out.length ? out : [{ text: '' }]
}

function listItems(list: any): InlineRun[][] {
  return (list.children ?? []).map((li: any) => {
    const p = (li.children ?? []).find((c: any) => c.type === 'paragraph')
    return inlineRuns(p ?? li)
  })
}

function guessLayout(blocks: Block[], index: number): Slide['layout'] {
  if (index === 0 && blocks.every((b) => b.type === 'heading' || b.type === 'paragraph' || b.type === 'kicker')) {
    return 'title'
  }
  return 'content'
}

// Group a two-column slide's blocks into columns, starting a new column at each heading.
function toColumns(blocks: Block[]): Block[][] {
  const cols: Block[][] = []
  for (const b of blocks) {
    if (b.type === 'heading' || cols.length === 0) cols.push([b])
    else cols[cols.length - 1].push(b)
  }
  return cols
}

/** Parse a (Marp-style) markdown string into a block-IR Deck. */
export function markdownToDeck(md: string): Deck {
  const body = md.replace(/^﻿?---\r?\n[\s\S]*?\r?\n---\r?\n/, '') // strip YAML front-matter
  const tree: any = unified().use(remarkParse).parse(body)
  const slides: Slide[] = []
  let blocks: Block[] = []
  let hint: Slide['layout'] | undefined

  const flush = (): void => {
    if (blocks.length) {
      const layout = hint ?? guessLayout(blocks, slides.length)
      slides.push(layout === 'two-column' ? { layout, columns: toColumns(blocks) } : { layout, blocks })
    }
    blocks = []
    hint = undefined
  }

  for (const node of (tree.children ?? []) as any[]) {
    if (node.type === 'thematicBreak') {
      flush()
    } else if (node.type === 'html') {
      const m = /<!--\s*layout:\s*([\w-]+)\s*-->/.exec(node.value ?? '')
      if (m && (m[1] === 'title' || m[1] === 'content' || m[1] === 'two-column')) hint = m[1] as Slide['layout']
    } else if (node.type === 'heading') {
      blocks.push({ type: 'heading', level: Math.min(3, node.depth) as 1 | 2 | 3, text: inlineRuns(node) })
    } else if (node.type === 'list') {
      blocks.push({ type: 'bullets', items: listItems(node) })
    } else if (node.type === 'paragraph') {
      const img = (node.children ?? []).find((c: any) => c.type === 'image')
      if (img && (node.children ?? []).length === 1) blocks.push({ type: 'image', src: img.url as string })
      else blocks.push({ type: 'paragraph', text: inlineRuns(node) })
    }
  }
  flush()
  return { slides }
}
