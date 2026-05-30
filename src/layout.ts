import { writeFileSync, unlinkSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { pathToFileURL } from 'node:url'
import { launch } from './browser.js'
import { extractSlide } from './extract.js'
import type { RenderResult, SlideLayout } from './types.js'

export async function layout(r: RenderResult, baseHref?: string): Promise<SlideLayout[]> {
  const base = baseHref ? `<base href="${baseHref}">` : ''
  const doc = `<!doctype html><html><head><meta charset="utf-8">${base}<style>${r.css}</style></head><body>${r.html}</body></html>`
  const browser = await launch()
  let tmpFile: string | undefined
  try {
    const page = await browser.newPage()
    await page.setViewport({ width: 1280, height: 720, deviceScaleFactor: 1 })
    if (baseHref) {
      // Navigate to a real file:// page so file:// images referenced by the deck
      // actually load. Chromium blocks file:// subresources from the about:blank
      // origin that setContent uses (even with --allow-file-access-from-files),
      // so an <img> would render at its broken/unloaded default size and be
      // mis-measured. A real file:// page + --allow-file-access-from-files (set
      // in browser.ts) lets the (possibly cross-directory) image loads succeed.
      tmpFile = join(tmpdir(), `mnp-${process.pid}-${Date.now()}.html`)
      writeFileSync(tmpFile, doc, 'utf8')
      await page.goto(pathToFileURL(tmpFile).href, { waitUntil: 'networkidle0' })
    } else {
      await page.setContent(doc, { waitUntil: 'networkidle0' })
    }
    // wait for fonts without returning a non-serializable value
    await page.evaluate(async () => { await (document as any).fonts?.ready })
    // Wait for every <img> to finish loading/decoding so getBoundingClientRect
    // reports the real rendered size (not the unloaded ~96px default).
    await page.evaluate(async () => {
      const imgs = Array.from(document.images)
      await Promise.all(imgs.map((img) => {
        if (img.complete && img.naturalWidth > 0) return Promise.resolve()
        return new Promise<void>((res) => {
          img.addEventListener('load', () => res(), { once: true })
          img.addEventListener('error', () => res(), { once: true })
        })
      }))
    })
    // Inject extractSlide by binding it to a local name (a bare named function
    // expression is NOT callable by its name in the eval scope), then map it
    // over every slide <section>.
    const expr = `(() => {
      // Provide a no-op __name shim in case esbuild/tsx injected __name() calls
      // into the serialised function (dev-mode only; tsc output is unaffected).
      const __name = (fn) => fn;
      const fn = ${extractSlide.toString()};
      return Array.from(document.querySelectorAll('section')).map((s) => fn(s));
    })()`
    const slides = (await page.evaluate(expr)) as SlideLayout[]
    // Resolve slide backgrounds. Sections flagged needsRaster (gradient/image
    // background that can't be expressed natively) are screenshotted with their
    // content hidden, yielding a decoration-only background image; text boxes
    // stay native on top. Solid fills pass through; others get no background.
    const sectionHandles = await page.$$('section')
    for (let i = 0; i < slides.length; i++) {
      const bg = (slides[i] as any).background as { needsRaster?: boolean; fill?: string } | undefined
      if (!bg?.needsRaster) {
        slides[i].background = bg?.fill ? { fill: bg.fill } : undefined
        continue
      }
      await page.evaluate((idx) => {
        const s = document.querySelectorAll('section')[idx] as HTMLElement
        ;(s as any).__hidden = Array.from(s.children).map((c) => [c, (c as HTMLElement).style.visibility])
        for (const c of Array.from(s.children)) (c as HTMLElement).style.visibility = 'hidden'
      }, i)
      const b64 = (await sectionHandles[i].screenshot({ encoding: 'base64', type: 'png' })) as string
      await page.evaluate((idx) => {
        const s = document.querySelectorAll('section')[idx] as HTMLElement
        for (const [c, v] of (s as any).__hidden as [HTMLElement, string][]) c.style.visibility = v
      }, i)
      slides[i].background = { imageDataUri: `data:image/png;base64,${b64}` }
    }
    return slides
  } finally {
    await browser.close()
    if (tmpFile) { try { unlinkSync(tmpFile) } catch { /* best-effort cleanup */ } }
  }
}
