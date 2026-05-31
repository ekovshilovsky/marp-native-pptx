# Marpoint

> **Native Marp → editable PowerPoint** — no LibreOffice, no fragmentation.
> npm package: [`marp-native-pptx`](https://www.npmjs.com/package/marp-native-pptx)

Native Marp → **editable** PowerPoint generator. Renders your deck with the same
engine Marp uses (`@marp-team/marp-core`), reads the rendered layout in headless
Chromium, and emits **native** PowerPoint shapes with `pptxgenjs`.

No PDF round-trip, no LibreOffice, **no text fragmentation** — an inline `` `code` ``
chip becomes one run inside one text box, exactly where the browser placed it.

## Install
```bash
npm install -g marp-native-pptx   # requires a Chrome/Chromium on PATH or CHROME_PATH
```

## Usage
```bash
marp-native-pptx deck.md -o deck.pptx --theme-set themes/ --allow-local-files
```

## Validating a .pptx is "PowerPoint-clean"

A file can be **schema-valid yet still trigger PowerPoint's "repair" prompt** —
PowerPoint enforces packaging and relationship-ownership rules the OOXML schema
doesn't express. The bundled validator checks both layers, so you can verify any
`.pptx` (not just ones this tool generates) will open cleanly.

```bash
marp-native-pptx-validate deck.pptx            # both layers
marp-native-pptx-validate deck.pptx --no-schema # PowerPoint-clean rules only
```
or programmatically:
```js
import { validatePptx } from 'marp-native-pptx/validate'
const findings = await validatePptx('deck.pptx') // [] === clean
```

**Layer 1 — PowerPoint-clean rules** (pure Node, always runs):

| Rule | Checks |
|---|---|
| `opc-content-types-first` | `[Content_Types].xml` is the first zip entry |
| `opc-no-directory-entries` | no zip directory entries |
| `ct-dangling-override` | no `<Override>` for a non-existent part |
| `ct-uncovered-part` | every part has a content type |
| `empty-text-run` | no empty `<a:t></a:t>` runs |
| `multiple-pPr` | ≤1 `<a:pPr>` per `<a:p>` |
| `empty-table-cell` | every `<a:tc>` has a paragraph |
| `table-column-count-mismatch` | each table row's cell count == grid column count |
| `table-graphicframe-namespace` | tables use `<p:graphicFrame>`, not `<a:graphicFrame>` |
| `invalid-shape-preset` | `prstGeom` uses valid preset names (not `oval`/`roundedRectangle`/…) |
| `duplicate-shape-id` | unique `<p:cNvPr id>` per slide |
| `shared-master-theme` | notes master owns a theme distinct from the slide master |
| `presentation-child-order` | `notesMasterIdLst` precedes `sldIdLst` |
| `missing-slide-size` | `presentation.xml` declares `<p:sldSz>` |
| `missing-color-map` | every slide master has a `<p:clrMap>` |
| `dangling-relationship` | every relationship `Target` resolves |
| `malformed-xml` | every XML part is well-formed |

**Layer 2 — OOXML schema** (optional): the Open XML SDK `OpenXmlValidator`, bundled
at `tools/ooxml-schema`. Requires .NET; if `dotnet` isn't on `PATH` the CLI prints a
skip note and still runs Layer 1. Run it directly with
`dotnet run --project tools/ooxml-schema -- deck.pptx`.

### Provenance

Most Layer-1 rules are **ours** — discovered first-hand by diffing PowerPoint's own
repaired output against the generated file.

Five rules were **pulled in from prior art** (all MIT-licensed), for causes our own
decks don't exercise:
- `table-column-count-mismatch`, `table-graphicframe-namespace`, `missing-slide-size`,
  `missing-color-map` — from VeeamHub/veeam-healthcheck's PowerPoint fix notes
- `invalid-shape-preset` — from the PptxGenJS repair-causes discussion (gitbrent/PptxGenJS#1449)

Several of our first-hand rules were later **independently corroborated** by that prior
art and by python-openxml/opc-diag (the OPC package-diff approach) — corroboration, not
origin.

## Block layouts & markdown (experimental)

Beyond the Marp/DOM path, a semantic **block IR** renders decks through layout
*templates* that anchor titles and place content in aligned regions/grids — the
"everything lines up" look — instead of vertically-centered flow. Author it as
JSON blocks, or straight from markdown:

```ts
import { markdownToPptx } from 'marp-native-pptx/blocks'
await markdownToPptx(md, 'deck.pptx')
//  ---  separates slides;  <!-- layout: title|content|two-column -->  picks a template
```

Both the Marp/DOM path and the block path feed the **same `emit` + validator**, so
the output is PowerPoint-clean either way. No browser is involved on the block path.

### Layout templates

`title` (optional hero image), `content`, `two-column`, `image-feature` (image
beside copy), `section` (full-bleed accent divider), `grid` (icon-chip feature
grid), `metrics` (big-number stat cards), `timeline` (horizontal axis with
alternating event labels), `steps` (numbered process), `quote`, `bento`
(mixed-size tile mosaic), `comparison` (side-by-side cards), `chart` (native bar
chart — bars from shapes, never an image), `stat-hero` (one giant statistic),
`closing` (sign-off / CTA), and `full-bleed` (image fills the slide with text
overlaid on a baked-in scrim). Image-driven layouts (`full-bleed`,
`image-feature`, `title` hero) embed real raster images; the showcase generates
cohesive per-theme cover/tile artwork sized to each region (`showcase-art.ts`),
swappable 1:1 with photographs. Raster images are placed with aspect-aware
`fit` — `cover` (fill + center-crop), `contain` (letterbox), or `fill` — so a
photo whose aspect differs from its region is cropped, never distorted
(`raster-dims.ts` decodes the intrinsic size; see `examples/calibration.ts` for
a visual proof). For a photo-real showcase, `examples/bake-images.ts` optionally
generates per-theme imagery with Gemini 2.5 Flash Image (set `GEMINI_API_KEY` in
`.env.local`) and runs it through the same `cover` path. Each layout is built from a small **auto-layout
engine** (`layout-engine.ts`: uniform grids,
even distribution, and shrink-to-fit text sizing) so content stays aligned and
never overflows its region — the same fit/scale behavior whether a template is
built-in, hand-authored, or generated. Text and chrome are native editable
shapes; `image` blocks (icons, photos, illustrations) embed as real images —
only the genuinely-raster parts are raster, never the text.

Layout correctness is locked by two test layers. A **geometry contract**
(`test/layout-geometry.test.ts`) asserts on the computed px geometry — the
deterministic source of truth that becomes the OOXML — that every box stays
on-canvas, step numbers sit concentric in their circles, and titles never
overlap their subtitles. On top of that, **golden-image visual regression**
(`test/visual/`, `npm run test:visual`) renders every layout under every theme
(all presets × all layouts) to PNG via LibreOffice and pixel-diffs each against
a committed baseline with `pixelmatch`, failing if more than 0.4% of pixels move. It's opt-in (the
fast `npm test` suite excludes it) and baselines are refreshed with
`npm run test:visual:update`.

### Themes

A curated, **tagged** preset set (`themes.ts`) — so a UI or an AI can pick by
intent rather than raw hex:

```ts
import { themeList, themesByTag, getTheme } from 'marp-native-pptx/blocks'
themesByTag('dark')          // -> Midnight-style presets
getTheme('vega')             // the house signature: electric indigo on cool paper
await blocksToPptx({ theme: getTheme('plum'), slides }, 'deck.pptx')
```

13 presets across `Light`/`Dark`/`Minimal`/`Bold`/`Professional`/`Playful`/`Editorial`,
each a full palette (accent, ink, muted, bg, surface, 6-color accent palette,
fonts). Every layout renders under every theme; generated text auto-picks a
legible ink for whatever color it lands on. `npx tsx examples/showcase.ts`
emits a deck of all themes × layouts and runs it through the validator.

## Honest limitations
- **Editable ≠ pixel-identical.** PowerPoint reflows text with its own engine, so a
  nearly-full line may wrap slightly differently than in Chrome. We eliminate
  fragmentation and overlap and match positions/styles closely; for pixel-identical
  output use an image export.
- **Fonts are not embedded** (a pptxgenjs limit). Fonts not installed in PowerPoint
  will substitute. We set explicit font faces from the rendered CSS.
- MVP scope: headings, paragraphs, list items, inline runs (bold/italic/link/code),
  and images. Tables, gradient/`html:true` rasterization, and self-check are on the
  roadmap (see `docs/design.md`).
