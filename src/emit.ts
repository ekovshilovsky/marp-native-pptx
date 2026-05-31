// pptxgenjs ships as CJS with a UMD type declaration (export as namespace PptxGenJS).
// With NodeNext + esModuleInterop the default import resolves to the namespace object,
// which has no TS construct signature. We import via require-style interop instead.
import { createRequire } from 'module'
import { readFileSync, writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import type { PptxModel } from './types.js'

const require = createRequire(import.meta.url)
// eslint-disable-next-line @typescript-eslint/no-require-imports, @typescript-eslint/no-explicit-any
const PptxGenJS = require('pptxgenjs') as any
// eslint-disable-next-line @typescript-eslint/no-require-imports, @typescript-eslint/no-explicit-any
const JSZip = require('jszip') as any

/**
 * Post-process the written .pptx zip to fix two PowerPoint "repair" triggers
 * that originate in pptxgenjs 3.12, neither of which has an API option:
 *
 *  1. Empty `<a:tblPr/>` (hardcoded at line 5176 of pptxgen.cjs.js): tables
 *     without a tableStyleId get flagged. We inject the built-in "Table Grid"
 *     style GUID.
 *  2. Duplicate `<p:cNvPr id="...">` within a slide: when a slide mixes text,
 *     table, and image shapes, pptxgenjs's per-slide id counter can collide,
 *     producing duplicate shape ids — which makes PowerPoint repair the file.
 *     We renumber every shape's id sequentially within each slide to guarantee
 *     uniqueness (shape ids are not referenced elsewhere in our output, so
 *     renumbering is safe).
 */
const TABLE_STYLE_ID = '{5940675A-B579-460E-94D1-54222C63F5DA}'
const TBLPR_WITH_STYLE = `<a:tblPr><a:tableStyleId>${TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>`

// Remove runs with empty text: <a:r>[<a:rPr.../>]<a:t></a:t></a:r>. The schema
// permits an empty <a:t> (so the SDK validator accepts it), but PowerPoint flags
// such runs and "repairs" the file. pptxgenjs emits them in auto-generated notes
// bodies and in empty table cells. Stripping them leaves a valid (possibly empty)
// paragraph — exactly what PowerPoint's own repair produces.
const stripEmptyRuns = (xml: string): string =>
  xml.replace(/<a:r>(?:<a:rPr\b[^>]*\/>|<a:rPr\b[\s\S]*?<\/a:rPr>)?<a:t><\/a:t><\/a:r>/g, '')

/* eslint-disable @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-unsafe-argument */
async function postProcess(pptxPath: string): Promise<void> {
  const src = await JSZip.loadAsync(readFileSync(pptxPath))
  // File entries in their current order, skipping zip directory entries.
  const fileNames = (Object.keys(src.files) as string[]).filter((n) => !src.files[n].dir)
  const existing = new Set(fileNames.map((n) => '/' + n))

  const parts = new Map<string, string | Buffer>()
  for (const name of fileNames) {
    if (/^ppt\/slides\/slide\d+\.xml$/.test(name)) {
      let xml: string = await src.files[name].async('string')
      // (a) give tables a valid style reference (pptxgenjs emits empty <a:tblPr/>)
      xml = xml.replace(/<a:tblPr\/>/g, TBLPR_WITH_STYLE)
      // (b) renumber shape ids sequentially so they are unique within the slide
      let idCounter = 0
      xml = xml.replace(/(<p:cNvPr id=")\d+(")/g, (_m, pre: string, post: string) => `${pre}${++idCounter}${post}`)
      // (c) keep only the FIRST <a:pPr> per <a:p> (pptxgenjs emits one per run;
      //     OOXML allows at most one, as the first child). The first carries the
      //     real paragraph props (set on run 0 in emit).
      xml = xml.replace(/<a:p>[\s\S]*?<\/a:p>/g, (para) => {
        let seen = false
        return para.replace(/<a:pPr\b[\s\S]*?<\/a:pPr>/g, (pp) => {
          if (seen) return ''
          seen = true
          return pp
        })
      })
      // (d) drop empty-text runs (PowerPoint repairs them; e.g. empty table cells)
      xml = stripEmptyRuns(xml)
      parts.set(name, xml)
    } else if (/^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name)) {
      // pptxgenjs auto-generates a notes body containing an empty run, which
      // PowerPoint flags → strip it.
      parts.set(name, stripEmptyRuns(await src.files[name].async('string')))
    } else if (name === 'ppt/presentation.xml') {
      // (d) move <p:notesMasterIdLst> before <p:sldIdLst> (schema child order)
      let xml: string = await src.files[name].async('string')
      const nm = xml.match(/<p:notesMasterIdLst>[\s\S]*?<\/p:notesMasterIdLst>/)
      if (nm && xml.indexOf('<p:sldIdLst') < xml.indexOf(nm[0])) {
        xml = xml.replace(nm[0], '').replace('<p:sldIdLst', `${nm[0]}<p:sldIdLst`)
      }
      parts.set(name, xml)
    } else if (name === '[Content_Types].xml') {
      // (e) drop <Override> entries for parts that don't exist. pptxgenjs declares
      //     one slideMaster override per slide but writes only slideMaster1 — the
      //     dangling overrides make PowerPoint repair the file.
      let ct: string = await src.files[name].async('string')
      ct = ct.replace(/<Override PartName="([^"]+)"[^>]*\/>/g, (m, part: string) =>
        existing.has(part) ? m : '',
      )
      // Drop <Default> entries for extensions no part actually uses (pptxgenjs
      // declares many; PowerPoint trims them). Always keep xml + rels.
      const usedExt = new Set(fileNames.map((n) => n.split('.').pop()!.toLowerCase()))
      usedExt.add('xml')
      usedExt.add('rels')
      ct = ct.replace(/<Default Extension="([^"]+)"[^>]*\/>/g, (m, ext: string) =>
        usedExt.has(ext.toLowerCase()) ? m : '',
      )
      parts.set(name, ct)
    } else if (name.endsWith('.rels') || name.endsWith('.xml')) {
      // keep all XML/rels as strings so later steps (e.g. theme-clone) can edit them
      parts.set(name, await src.files[name].async('string'))
    } else {
      parts.set(name, await src.files[name].async('nodebuffer'))
    }
  }

  // (f) A notes master must not SHARE the slide master's theme part. pptxgenjs
  //     points the notes master's theme relationship at the slide master's
  //     theme1.xml; PowerPoint treats that as corruption and "repairs" it by
  //     cloning the theme. Replicate that: give each notes master its own theme.
  const extraParts: string[] = []
  const themeTarget = (rels: string): string | null => {
    const m =
      rels.match(/<Relationship[^>]*Type="[^"]*\/theme"[^>]*Target="([^"]+)"/) ||
      rels.match(/<Relationship[^>]*Target="([^"]+)"[^>]*Type="[^"]*\/theme"/)
    return m ? m[1] : null
  }
  const resolve = (relsName: string, target: string): string => {
    const segs = (relsName.replace(/\/_rels\/[^/]+$/, '') + '/' + target).split('/')
    const stack: string[] = []
    for (const s of segs) {
      if (s === '..') stack.pop()
      else if (s !== '.' && s !== '') stack.push(s)
    }
    return stack.join('/')
  }
  const slideMasterThemes = new Set<string>()
  for (const n of fileNames) {
    if (/^ppt\/slideMasters\/_rels\/.+\.rels$/.test(n)) {
      const t = themeTarget(parts.get(n) as string)
      if (t) slideMasterThemes.add(resolve(n, t))
    }
  }
  let themeMax = Math.max(
    0,
    ...fileNames.map((n) => {
      const m = n.match(/^ppt\/theme\/theme(\d+)\.xml$/)
      return m ? Number(m[1]) : 0
    }),
  )
  let ctThemeAdds = ''
  for (const n of fileNames) {
    if (!/^ppt\/notesMasters\/_rels\/.+\.rels$/.test(n)) continue
    const rels = parts.get(n) as string
    const t = themeTarget(rels)
    if (!t) continue
    if (!slideMasterThemes.has(resolve(n, t))) continue // already owns its theme
    const newName = `ppt/theme/theme${++themeMax}.xml`
    const sharedPath = resolve(n, t)
    parts.set(newName, parts.get(sharedPath) as string)
    extraParts.push(newName)
    ctThemeAdds += `<Override PartName="/${newName}" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`
    parts.set(n, rels.replace(t, t.replace(/theme\d+\.xml$/, `theme${themeMax}.xml`)))
  }
  if (ctThemeAdds) {
    const ct = parts.get('[Content_Types].xml') as string
    parts.set('[Content_Types].xml', ct.replace('</Types>', `${ctThemeAdds}</Types>`))
  }

  // (g) rebuild a fresh package with OPC-correct structure: [Content_Types].xml
  //     FIRST and NO directory entries (our prior re-zip buried content-types and
  //     emitted folder entries, which strict readers / PowerPoint reject).
  const out = new JSZip()
  out.file('[Content_Types].xml', parts.get('[Content_Types].xml') as string, { createFolders: false })
  for (const name of [...fileNames, ...extraParts]) {
    if (name === '[Content_Types].xml') continue
    out.file(name, parts.get(name)!, { createFolders: false })
  }
  const buf: Buffer = await out.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  writeFileSync(pptxPath, buf)
}
/* eslint-enable @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-unsafe-argument */

export async function emit(model: PptxModel, outPath: string): Promise<void> {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-call
  const pptx = new PptxGenJS()
  pptx.defineLayout({ name: 'CUSTOM', width: model.slideWidthIn, height: model.slideHeightIn })
  pptx.layout = 'CUSTOM'

  for (const slide of model.slides) {
    const s = pptx.addSlide()
    if (slide.background?.fill) s.background = { color: slide.background.fill }
    else if (slide.background?.imageDataUri) s.background = { data: slide.background.imageDataUri }
    for (const shape of slide.shapes) {
      if (shape.kind === 'shape') {
        const prst = shape.preset ?? (shape.radiusIn && shape.radiusIn > 0 ? 'roundRect' : 'rect')
        s.addShape(prst, {
          x: shape.xIn, y: shape.yIn, w: shape.wIn, h: shape.hIn,
          fill: shape.fill ? { color: shape.fill } : { type: 'none' },
          line: shape.line ? { color: shape.line.color, width: shape.line.widthPt } : { type: 'none' },
          ...(shape.radiusIn ? { rectRadius: shape.radiusIn } : {}),
        })
        continue
      }
      if (shape.kind === 'image') {
        const geom = { x: shape.xIn, y: shape.yIn, w: shape.wIn, h: shape.hIn }
        if (shape.src.startsWith('data:')) {
          s.addImage({ data: shape.src, ...geom })
        } else if (shape.src.startsWith('file://')) {
          s.addImage({ path: fileURLToPath(shape.src), ...geom })
        } else {
          s.addImage({ path: shape.src, ...geom })
        }
        continue
      }
      if (shape.kind === 'table') {
        const rows = shape.rows.map((row) =>
          row.map((cell) => ({
            // An empty cell must still yield a paragraph — a table cell whose
            // <a:txBody> has no <a:p> is invalid (PowerPoint repair). Use '' so
            // pptxgenjs emits a single empty paragraph.
            text: cell.runs.length
              ? cell.runs.map((r) => ({
                  text: r.text,
                  options: {
                    bold: r.bold, italic: r.italic, color: r.color,
                    fontFace: r.fontFace, fontSize: r.fontSize,
                    highlight: r.highlight,
                  },
                }))
              : '',
            options: {
              fill: cell.fill ? { color: cell.fill } : undefined,
              align: cell.align as any,
              valign: 'top' as const,
              bold: cell.header,
            },
          })),
        )
        s.addTable(rows as any, {
          x: shape.xIn, y: shape.yIn, w: shape.wIn, h: shape.hIn,
          colW: shape.colWidthsIn, border: { type: 'solid', pt: 0.5, color: 'D0D0D0' },
          autoPage: false,
        })
        continue
      }
      const textRuns = shape.paragraphs.flatMap((p, pi) =>
        p.runs.map((r, ri) => ({
          text: r.text,
          options: {
            bold: r.bold,
            italic: r.italic,
            color: r.color,
            fontFace: r.fontFace,
            fontSize: r.fontSize,
            highlight: r.highlight,
            underline: r.underline ? { style: 'sng' as const } : undefined,
            strike: r.strike,
            // Paragraph-level props go ONLY on the first run of each paragraph.
            // pptxgenjs emits an <a:pPr> for every fragment that carries them,
            // and OOXML allows at most one <a:pPr> per <a:p> (as the first child)
            // — putting them on every run produced duplicate/misplaced pPr that
            // made PowerPoint flag the file for repair.
            ...(ri === 0
              ? {
                  align: p.align as any,
                  bullet: p.bullet ? { indent: p.bullet.indentLevel * 18 } : undefined,
                  lineSpacing: p.lineSpacingPt,
                }
              : {}),
            // break after the last run of a paragraph, except the final paragraph
            breakLine: ri === p.runs.length - 1 && pi < shape.paragraphs.length - 1,
          },
        })),
      )
      s.addText(textRuns as any, {
        x: shape.xIn, y: shape.yIn, w: shape.wIn, h: shape.hIn,
        margin: 0, valign: shape.valign, wrap: true, autoFit: false, isTextBox: true,
      })
    }
  }
  await pptx.writeFile({ fileName: outPath })
  await postProcess(outPath)
}
