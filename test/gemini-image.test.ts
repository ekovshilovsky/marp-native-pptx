import { afterEach, describe, expect, it, vi } from 'vitest'
import { generateImage } from '../src/gemini-image.js'

const PNG_B64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQAY3Y2wAAAAAElFTkSuQmCC'

function mockFetch(payload: unknown, ok = true, status = 200): typeof fetch {
  return vi.fn(async () => ({
    ok,
    status,
    json: async () => payload,
    text: async () => JSON.stringify(payload),
  })) as unknown as typeof fetch
}

afterEach(() => {
  vi.restoreAllMocks()
})

describe('generateImage (Gemini 2.5 Flash Image)', () => {
  it('POSTs the prompt with the key in the x-goog-api-key header and returns PNG bytes', async () => {
    const fetchMock = mockFetch({ candidates: [{ content: { parts: [{ inlineData: { mimeType: 'image/png', data: PNG_B64 } }] } }] })
    vi.stubGlobal('fetch', fetchMock)

    const buf = await generateImage('a calm smart home', { apiKey: 'AIzaTEST', aspectRatio: '16:9' })

    expect(buf).toBeInstanceOf(Buffer)
    expect(buf.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))).toBe(true)
    // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
    const [url, init] = (fetchMock as unknown as { mock: { calls: [string, RequestInit][] } }).mock.calls[0]
    expect(url).toContain('gemini-2.5-flash-image:generateContent')
    expect((init.headers as Record<string, string>)['x-goog-api-key']).toBe('AIzaTEST')
    const body = JSON.parse(init.body as string)
    expect(body.contents[0].parts[0].text).toBe('a calm smart home')
  })

  it('accepts snake_case inline_data too (REST variants differ)', async () => {
    vi.stubGlobal('fetch', mockFetch({ candidates: [{ content: { parts: [{ text: 'desc' }, { inline_data: { mime_type: 'image/png', data: PNG_B64 } }] } }] }))
    const buf = await generateImage('x', { apiKey: 'k' })
    expect(buf.length).toBeGreaterThan(0)
  })

  it('throws with the server message on a non-2xx response', async () => {
    vi.stubGlobal('fetch', mockFetch({ error: { message: 'API key not valid' } }, false, 400))
    await expect(generateImage('x', { apiKey: 'bad' })).rejects.toThrow(/400/)
  })

  it('throws when the response carries no image part', async () => {
    vi.stubGlobal('fetch', mockFetch({ candidates: [{ content: { parts: [{ text: 'no image here' }] } }] }))
    await expect(generateImage('x', { apiKey: 'k' })).rejects.toThrow(/no image/i)
  })
})
