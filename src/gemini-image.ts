// Text-to-image via Gemini 2.5 Flash Image ("nano banana").
//
// Used only to bake showcase artwork (examples/bake-images.ts) — never on the
// core deck path, which stays network-free. The returned PNG bytes flow through
// the ordinary `image` block + cover-fit, identically to any other raster.
//
// Endpoint/auth/response shape per ai.google.dev/gemini-api/docs/image-generation.
export interface GenImageOpts {
  apiKey: string
  /** override the model id (default gemini-2.5-flash-image) */
  model?: string
  /** '1:1' | '3:4' | '4:3' | '16:9' | '9:16' — requested, then cover-fit anyway */
  aspectRatio?: string
  signal?: AbortSignal
}

interface InlinePart {
  inlineData?: { data?: string; mimeType?: string }
  inline_data?: { data?: string; mime_type?: string }
  text?: string
}

const API_BASE = 'https://generativelanguage.googleapis.com/v1beta/models'

/** Generate one image from a text prompt; resolves to raw PNG bytes. */
export async function generateImage(prompt: string, opts: GenImageOpts): Promise<Buffer> {
  const model = opts.model ?? 'gemini-2.5-flash-image'
  const body = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseModalities: ['IMAGE'],
      ...(opts.aspectRatio ? { imageConfig: { aspectRatio: opts.aspectRatio } } : {}),
    },
  }
  const res = await fetch(`${API_BASE}/${model}:generateContent`, {
    method: 'POST',
    headers: { 'content-type': 'application/json', 'x-goog-api-key': opts.apiKey },
    body: JSON.stringify(body),
    signal: opts.signal,
  })
  if (!res.ok) {
    const detail = await res.text().catch(() => '')
    throw new Error(`Gemini ${res.status}: ${detail.slice(0, 400)}`)
  }
  const json = (await res.json()) as { candidates?: { content?: { parts?: InlinePart[] } }[] }
  const parts = json.candidates?.[0]?.content?.parts ?? []
  for (const p of parts) {
    const data = p.inlineData?.data ?? p.inline_data?.data
    if (data) return Buffer.from(data, 'base64')
  }
  throw new Error(`Gemini response contained no image part: ${JSON.stringify(json).slice(0, 400)}`)
}
