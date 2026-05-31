// Generate the showcase deck and validate it.
//
//   npx tsx examples/showcase.ts            # curated gallery -> showcase.pptx
//   npx tsx examples/showcase.ts --matrix   # full layout x theme matrix
//
// The output is fully native (editable, zero-image) PowerPoint and is checked
// with the bundled PowerPoint-clean validator before it is written.
import { mapToPptx } from '../src/map.js'
import { emit } from '../src/emit.js'
import { validatePptx } from '../src/validate.js'
import { buildGallery, buildMatrix, LAYOUT_ORDER } from '../src/showcase-deck.js'
import { bakeShowcaseAssets } from '../src/showcase-assets.js'
import { themeList } from '../src/themes.js'

async function main() {
  const matrix = process.argv.includes('--matrix')
  console.log('baking icon + illustration assets...')
  const assets = await bakeShowcaseAssets()
  const slides = matrix ? buildMatrix(assets) : buildGallery(assets)
  const out = matrix ? 'showcase-matrix.pptx' : 'showcase.pptx'

  const model = mapToPptx(slides, 13.333, 7.5)
  const images = model.slides.flatMap((s) => s.shapes).filter((sh) => sh.kind === 'image').length
  await emit(model, out)

  const findings = await validatePptx(out)
  console.log(`themes:   ${themeList().length}  (${themeList().map((t) => t.name).join(', ')})`)
  console.log(`layouts:  ${LAYOUT_ORDER.length}  (${LAYOUT_ORDER.join(', ')})`)
  console.log(`slides:   ${slides.length}`)
  console.log(`images:   ${images}  (0 = fully native / editable)`)
  console.log(`validator: ${findings.length === 0 ? 'PowerPoint-clean (0 findings)' : `${findings.length} finding(s)`}`)
  for (const f of findings) console.log(`  - [${f.severity}] ${f.rule}: ${f.message}`)
  console.log(`wrote ${out}`)
}

main().catch((e) => {
  console.error(e)
  process.exit(1)
})
