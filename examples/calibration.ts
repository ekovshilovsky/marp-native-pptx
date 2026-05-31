// Image-fit calibration: render one non-square "test card" into a square box
// under each fit mode, side by side, so distortion is visible at a glance.
//
//   fill    — stretches to the box: the circle becomes an ellipse (the old bug)
//   cover   — scales to fill + center-crops: the circle stays round
//   contain — fits the whole image, letterboxed: round, with surface margins
//
//   npx tsx examples/calibration.ts        # -> /tmp/calibration.pptx (+ .png if soffice present)
import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { calibrationSvg } from '../src/showcase-art.js'
import { svgToPngDataUri } from '../src/rasterize.js'
import { mapToPptx } from '../src/map.js'
import { emit } from '../src/emit.js'
import type { ImageFit, LayoutBox, Run, SlideLayout } from '../src/types.js'

const W = 1280
const H = 720

function run(text: string, sizePt: number, color = '1a1a2e'): Run {
  return { text, style: { bold: true, italic: false, color, fontFace: 'Arial', sizePt } }
}
function label(text: string, xPx: number, yPx: number, wPx: number): LayoutBox {
  return {
    kind: 'text',
    rect: { xPx, yPx, wPx, hPx: 40 },
    valign: 'middle',
    paras: [{ runs: [run(text, 18)], align: 'center', lineSpacingPt: 22 }],
  }
}

async function main(): Promise<void> {
  // A wide source (8:3) and a tall source (2:3) — both non-square, so a square
  // box must crop or letterbox; a stretch would visibly distort the rings.
  const wide = await svgToPngDataUri(calibrationSvg(1600, 600, '8:3 source'), { width: 1600, height: 600 })
  const tall = await svgToPngDataUri(calibrationSvg(600, 900, '2:3 source'), { width: 600, height: 900 })

  const side = 300
  const top = 150
  const col = [80, 380, 680, 980]
  const cells: { src: string; fit: ImageFit; cap: string }[] = [
    { src: wide, fit: 'fill', cap: 'fill (8:3 → square)' },
    { src: wide, fit: 'cover', cap: 'cover (8:3 → square)' },
    { src: wide, fit: 'contain', cap: 'contain (8:3 → square)' },
    { src: tall, fit: 'cover', cap: 'cover (2:3 → square)' },
  ]

  const boxes: LayoutBox[] = [
    { kind: 'text', rect: { xPx: 0, yPx: 40, wPx: W, hPx: 50 }, valign: 'middle', paras: [{ runs: [run('Image-fit calibration — rings stay round only when fit is correct', 24)], align: 'center', lineSpacingPt: 28 }] },
  ]
  cells.forEach((c, i) => {
    const x = col[i]
    boxes.push({ kind: 'shape', rect: { xPx: x, yPx: top, wPx: side, hPx: side }, fill: 'ececf5', radiusPx: 12 })
    boxes.push({ kind: 'image', rect: { xPx: x, yPx: top, wPx: side, hPx: side }, src: c.src, fit: c.fit })
    boxes.push(label(c.cap, x - 10, top + side + 12, side + 20))
  })

  const slide: SlideLayout = { sizePx: { w: W, h: H }, boxes, background: { fill: 'f7f7fb' } }
  const out = join(mkdtempSync(join(tmpdir(), 'mnp-calib-')), 'calibration.pptx')
  await emit(mapToPptx([slide], 13.333, 7.5), out)
  console.log('pptx:', out)

  // Best-effort PNG so the result is viewable without PowerPoint.
  try {
    const dir = join(out, '..')
    execFileSync('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', dir, out], { stdio: 'ignore' })
    const pdf = join(dir, 'calibration.pdf')
    if (existsSync(pdf)) {
      execFileSync('pdftoppm', ['-png', '-r', '110', pdf, join(dir, 'calibration')], { stdio: 'ignore' })
      console.log('png: ', join(dir, 'calibration-1.png'))
    }
  } catch {
    console.log('(install libreoffice + poppler for an auto PNG preview)')
  }
}

void main()
