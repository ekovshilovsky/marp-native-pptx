// Rasterize SVG markup to a PNG data URI using the bundled Chromium.
//
// PowerPoint's SVG support is inconsistent across versions, so we bake vector
// art (icons, illustrations) down to PNG — which every PowerPoint renders
// identically. This is intentionally NOT part of the core block path (that
// path stays browser-free); it's a helper for preparing image assets that are
// then embedded as ordinary `image` blocks.
import { launch } from './browser.js'

export interface RasterOpts {
  width: number
  height: number
  scale?: number // device pixel ratio for crispness (default 2)
}

/** Render one SVG string to a `data:image/png;base64,...` URI at width×height. */
export async function svgToPngDataUri(svg: string, opts: RasterOpts): Promise<string> {
  return (await svgsToPngDataUris([svg], opts))[0]
}

/** Batch variant — reuses a single browser for many icons (much faster). */
export async function svgsToPngDataUris(svgs: string[], opts: RasterOpts): Promise<string[]> {
  const { width, height, scale = 2 } = opts
  const browser = await launch()
  try {
    const page = await browser.newPage()
    await page.setViewport({ width, height, deviceScaleFactor: scale })
    const out: string[] = []
    for (const svg of svgs) {
      // transparent page; the SVG fills the viewport exactly
      const doc = `<!doctype html><html><head><style>
        *{margin:0;padding:0}html,body{width:${width}px;height:${height}px;background:transparent}
        svg{display:block;width:${width}px;height:${height}px}
      </style></head><body>${svg}</body></html>`
      await page.setContent(doc, { waitUntil: 'domcontentloaded' })
      const b64 = (await page.screenshot({ encoding: 'base64', type: 'png', omitBackground: true })) as string
      out.push(`data:image/png;base64,${b64}`)
    }
    return out
  } finally {
    await browser.close()
  }
}
