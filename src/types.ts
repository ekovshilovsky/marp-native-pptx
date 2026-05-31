// Geometry, in CSS pixels relative to the slide's top-left origin.
export interface Rect { xPx: number; yPx: number; wPx: number; hPx: number }

// One maximal text span sharing a single computed style.
export interface RunStyle {
  fontFace: string
  sizePt: number
  bold: boolean
  italic: boolean
  color: string        // 6-hex, no leading '#'
  fill?: string        // run background (e.g. code-chip), 6-hex or undefined
  underline?: boolean
  strike?: boolean
}
export interface Run { text: string; style: RunStyle }

export interface BulletSpec { type: 'bullet' | 'number'; indentLevel: number }

export interface Paragraph {
  runs: Run[]
  align: 'left' | 'center' | 'right' | 'justify'
  lineSpacingPt: number
  bullet?: BulletSpec
}

export interface TableCell {
  paras: Paragraph[]
  fill?: string        // cell background, 6-hex or undefined
  header: boolean      // <th> vs <td>
}

export interface LineSpec { color: string; widthPt: number }

export type LayoutBox =
  | { kind: 'text'; rect: Rect; paras: Paragraph[]; valign: 'top' | 'middle' | 'bottom' }
  | { kind: 'image'; rect: Rect; src: string; fit?: ImageFit }
  | { kind: 'table'; rect: Rect; rows: TableCell[][]; colWidthsPx: number[] }
  | { kind: 'shape'; rect: Rect; fill?: string; line?: LineSpec; radiusPx?: number; preset?: 'rect' | 'roundRect' | 'ellipse' }

export interface SlideLayout { sizePx: { w: number; h: number }; boxes: LayoutBox[]; background?: { fill?: string; imageDataUri?: string } }

// ---- PPTX model (units already converted to inches/points) ----
export interface PptxRunOpts {
  text: string
  bold: boolean
  italic: boolean
  color: string
  fontFace: string
  fontSize: number          // points
  highlight?: string        // run fill, hex
  underline?: boolean
  strike?: boolean
}
export interface PptxTextBox {
  kind: 'text'
  xIn: number; yIn: number; wIn: number; hIn: number
  paragraphs: { runs: PptxRunOpts[]; align: string; lineSpacingPt: number; bullet?: BulletSpec }[]
  valign: 'top' | 'middle' | 'bottom'
}
// How a raster fills its box. 'fill' stretches to the box (the default — right
// for icons/illustrations authored at the target aspect). 'cover' scales to fill
// and center-crops; 'contain' fits the whole image inside, letterboxed. cover/
// contain require the image's intrinsic aspect, decoded at emit time.
export type ImageFit = 'fill' | 'cover' | 'contain'
export interface PptxImage {
  kind: 'image'
  xIn: number; yIn: number; wIn: number; hIn: number
  src: string
  fit?: ImageFit
}
export interface PptxTableCell {
  runs: PptxRunOpts[]
  align: string
  fill?: string
  header: boolean
}
export interface PptxTable {
  kind: 'table'
  xIn: number; yIn: number; wIn: number; hIn: number
  rows: PptxTableCell[][]
  colWidthsIn: number[]
}
export interface PptxRect {
  kind: 'shape'
  xIn: number; yIn: number; wIn: number; hIn: number
  fill?: string
  line?: LineSpec
  radiusIn?: number
  preset?: 'rect' | 'roundRect' | 'ellipse'
}
export type PptxShape = PptxTextBox | PptxImage | PptxTable | PptxRect
export interface PptxSlide { shapes: PptxShape[]; background?: { fill?: string; imageDataUri?: string } }
export interface PptxModel { slideWidthIn: number; slideHeightIn: number; slides: PptxSlide[] }

export interface RenderOptions { themeSet?: string[]; allowLocalFiles?: boolean }
export interface RenderResult { html: string; css: string }
