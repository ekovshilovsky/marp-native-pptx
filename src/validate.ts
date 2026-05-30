// PowerPoint-clean validator. The Open XML SDK schema validator accepts files
// that PowerPoint still "repairs" — because PowerPoint enforces packaging and
// relationship-ownership rules the schema doesn't express. This module encodes
// those rules (learned by diffing PowerPoint's own repaired output) so we can
// assert a .pptx will open cleanly, and validate arbitrary files.
//
// For full OOXML *schema* validation, additionally run the bundled Open XML SDK
// validator (tools/ooxml-schema) via schema-validate.ts — that layer is optional
// and requires .NET; this layer is pure Node and always available.
import { createRequire } from 'module'
import { readFileSync } from 'node:fs'
import { XMLValidator } from 'fast-xml-parser'

const require = createRequire(import.meta.url)
// eslint-disable-next-line @typescript-eslint/no-require-imports, @typescript-eslint/no-explicit-any
const JSZip = require('jszip') as any

export interface Finding {
  rule: string
  part?: string
  detail: string
}

/** Resolve an OPC relationship Target (relative to the part's directory). */
function resolveTarget(partDir: string, target: string): string {
  const stack: string[] = partDir ? partDir.split('/') : []
  for (const seg of target.split('/')) {
    if (seg === '..') stack.pop()
    else if (seg !== '.' && seg !== '') stack.push(seg)
  }
  return stack.join('/')
}

/** First theme part referenced by a .rels file, resolved to a package path. */
function themeTargetPath(relsName: string, relsXml: string): string | undefined {
  const m =
    relsXml.match(/<Relationship[^>]*Type="[^"]*\/theme"[^>]*Target="([^"]+)"/) ||
    relsXml.match(/<Relationship[^>]*Target="([^"]+)"[^>]*Type="[^"]*\/theme"/)
  if (!m) return undefined
  const partDir = relsName.replace(/\/_rels\/[^/]+$/, '')
  return resolveTarget(partDir, m[1])
}

/**
 * Validate a .pptx against the "PowerPoint-clean" rules. Returns a list of
 * findings; an empty list means the file passes every rule. Accepts a path or
 * a buffer.
 */
export async function validatePptx(pathOrBuffer: string | Buffer): Promise<Finding[]> {
  const buf = typeof pathOrBuffer === 'string' ? readFileSync(pathOrBuffer) : pathOrBuffer
  // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access
  const zip = await JSZip.loadAsync(buf)
  const findings: Finding[] = []
  // eslint-disable-next-line @typescript-eslint/no-unsafe-argument, @typescript-eslint/no-unsafe-member-access
  const names: string[] = Object.keys(zip.files)
  // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
  const fileNames = names.filter((n) => !zip.files[n].dir)
  // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-return
  const read = (n: string): Promise<string> => zip.files[n].async('string')
  const existing = new Set(fileNames.map((n) => '/' + n))

  // Rule: [Content_Types].xml must be the first zip entry.
  if (names[0] !== '[Content_Types].xml') {
    findings.push({ rule: 'opc-content-types-first', detail: `first zip entry is '${names[0]}', expected [Content_Types].xml` })
  }
  // Rule: no zip directory entries.
  const dirs = names.filter((n) => n.endsWith('/'))
  if (dirs.length) {
    findings.push({ rule: 'opc-no-directory-entries', detail: `${dirs.length} directory entries present (e.g. '${dirs[0]}')` })
  }

  // Content types: dangling overrides + part coverage.
  if (existing.has('/[Content_Types].xml')) {
    const ct = await read('[Content_Types].xml')
    const overrides = new Set([...ct.matchAll(/<Override PartName="([^"]+)"/g)].map((m) => m[1]))
    const defaults = new Set([...ct.matchAll(/<Default Extension="([^"]+)"/g)].map((m) => m[1].toLowerCase()))
    for (const p of overrides) {
      if (!existing.has(p)) findings.push({ rule: 'ct-dangling-override', part: '[Content_Types].xml', detail: `Override for missing part ${p}` })
    }
    for (const n of fileNames) {
      if (n === '[Content_Types].xml') continue
      const ext = n.split('.').pop()!.toLowerCase()
      if (!overrides.has('/' + n) && !defaults.has(ext)) {
        findings.push({ rule: 'ct-uncovered-part', part: n, detail: `part has no content type (extension .${ext})` })
      }
    }
  }

  // Per-XML-part rules.
  for (const n of fileNames) {
    if (!n.endsWith('.xml') && !n.endsWith('.rels')) continue
    const xml = await read(n)
    // Well-formed XML.
    const wf = XMLValidator.validate(xml)
    if (wf !== true) {
      findings.push({ rule: 'malformed-xml', part: n, detail: typeof wf === 'object' ? wf.err.msg : 'not well-formed' })
      continue // other rules assume parseable XML
    }
    if (!n.startsWith('ppt/')) continue
    // Empty-text runs.
    if (xml.includes('<a:t></a:t>')) {
      findings.push({ rule: 'empty-text-run', part: n, detail: 'contains an empty <a:t></a:t> run (PowerPoint repairs these)' })
    }
    // <= 1 <a:pPr> per <a:p>.
    for (const m of xml.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)) {
      const c = (m[1].match(/<a:pPr\b/g) || []).length
      if (c > 1) {
        findings.push({ rule: 'multiple-pPr', part: n, detail: `a paragraph has ${c} <a:pPr> elements (max 1, as first child)` })
        break
      }
    }
    // Every table cell has a paragraph.
    for (const m of xml.matchAll(/<a:tc>([\s\S]*?)<\/a:tc>/g)) {
      if (!/<a:p[>\s/]/.test(m[1])) {
        findings.push({ rule: 'empty-table-cell', part: n, detail: 'a table cell <a:tc> has no <a:p>' })
        break
      }
    }
    // Invalid prstGeom preset aliases — PowerPoint drops the shape on repair.
    // (Pulled in from prior art: gitbrent/PptxGenJS#1449, MIT. Valid presets are an enum.)
    const BAD_PRESET: Record<string, string> = {
      oval: 'ellipse',
      roundedRectangle: 'roundRect',
      roundedRect: 'roundRect',
      rectangle: 'rect',
    }
    for (const m of xml.matchAll(/<a:prstGeom\b[^>]*\bprst="([^"]+)"/g)) {
      if (BAD_PRESET[m[1]]) {
        findings.push({ rule: 'invalid-shape-preset', part: n, detail: `prstGeom prst="${m[1]}" is not a valid preset (use "${BAD_PRESET[m[1]]}")` })
      }
    }
    // Tables: the graphicFrame must be in the p: namespace, and every row's
    // cell count must equal the grid's column count.
    // (Pulled in from prior art: VeeamHub/veeam-healthcheck fix-notes, MIT.)
    if (xml.includes('<a:tbl>')) {
      if (/<a:graphicFrame\b/.test(xml)) {
        findings.push({ rule: 'table-graphicframe-namespace', part: n, detail: 'table <graphicFrame> is in the a: namespace; PowerPoint requires <p:graphicFrame>' })
      }
      for (const t of xml.match(/<a:tbl>[\s\S]*?<\/a:tbl>/g) ?? []) {
        const cols = (t.match(/<a:gridCol\b/g) ?? []).length
        if (!cols) continue
        for (const row of t.match(/<a:tr\b[\s\S]*?<\/a:tr>/g) ?? []) {
          const cells = (row.match(/<a:tc[ />]/g) ?? []).length
          if (cells !== cols) {
            findings.push({ rule: 'table-column-count-mismatch', part: n, detail: `a table row has ${cells} cell(s) but the grid defines ${cols} column(s)` })
            break
          }
        }
      }
    }
    // Unique shape ids within a slide.
    if (/^ppt\/slides\/slide\d+\.xml$/.test(n)) {
      const ids = [...xml.matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => m[1])
      if (new Set(ids).size !== ids.length) {
        findings.push({ rule: 'duplicate-shape-id', part: n, detail: 'duplicate <p:cNvPr id> values within the slide' })
      }
    }
  }

  // presentation.xml: notesMasterIdLst must precede sldIdLst.
  if (existing.has('/ppt/presentation.xml')) {
    const p = await read('ppt/presentation.xml')
    const iN = p.indexOf('<p:notesMasterIdLst')
    const iS = p.indexOf('<p:sldIdLst')
    if (iN >= 0 && iS >= 0 && iN > iS) {
      findings.push({ rule: 'presentation-child-order', part: 'ppt/presentation.xml', detail: 'notesMasterIdLst appears after sldIdLst (schema requires it before)' })
    }
    // A presentation must declare a slide size.
    // (Pulled in from prior art: VeeamHub/veeam-healthcheck fix-notes, MIT.)
    if (!/<p:sldSz\b/.test(p)) {
      findings.push({ rule: 'missing-slide-size', part: 'ppt/presentation.xml', detail: 'no <p:sldSz> (slide size) — PowerPoint repairs presentations without one' })
    }
  }
  // Every slide master must declare a color map (<p:clrMap>).
  // (Pulled in from prior art: VeeamHub/veeam-healthcheck fix-notes, MIT.)
  for (const n of fileNames.filter((m) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(m))) {
    if (!/<p:clrMap\b/.test(await read(n))) {
      findings.push({ rule: 'missing-color-map', part: n, detail: 'slide master has no <p:clrMap> color mapping' })
    }
  }

  // Relationship integrity + theme ownership.
  const slideMasterThemes = new Set<string>()
  const notesMasterRels: string[] = []
  for (const n of fileNames) {
    if (!n.endsWith('.rels')) continue
    const rels = await read(n)
    const partDir = n.replace(/\/?_rels\/[^/]+\.rels$/, '')
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      const tag = m[0]
      if (/TargetMode="External"/.test(tag)) continue
      const t = tag.match(/Target="([^"]+)"/)
      if (!t) continue
      const resolved = resolveTarget(partDir === n ? '' : partDir, t[1])
      if (!existing.has('/' + resolved)) {
        findings.push({ rule: 'dangling-relationship', part: n, detail: `Target '${t[1]}' resolves to missing part '${resolved}'` })
      }
    }
    if (/^ppt\/slideMasters\/_rels\/.+\.rels$/.test(n)) {
      const tp = themeTargetPath(n, rels)
      if (tp) slideMasterThemes.add(tp)
    }
    if (/^ppt\/notesMasters\/_rels\/.+\.rels$/.test(n)) notesMasterRels.push(n)
  }
  for (const n of notesMasterRels) {
    const tp = themeTargetPath(n, await read(n))
    if (tp && slideMasterThemes.has(tp)) {
      findings.push({ rule: 'shared-master-theme', part: n, detail: `notes master shares a slide master's theme part '${tp}' (each master must own a distinct theme)` })
    }
  }

  return findings
}
