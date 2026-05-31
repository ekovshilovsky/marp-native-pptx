import { describe, expect, it } from 'vitest'
import { PNG } from 'pngjs'
import { rasterDims } from '../src/raster-dims.js'

// Build a real PNG of known dimensions and return it as a data URI.
function pngDataUri(w: number, h: number): string {
  const png = new PNG({ width: w, height: h })
  png.data.fill(0xff)
  return 'data:image/png;base64,' + PNG.sync.write(png).toString('base64')
}

// A minimal baseline-JPEG (SOF0) header carrying known dimensions. Only the
// bytes rasterDims walks need to be valid: SOI, then an SOF0 segment with h/w.
function jpegHeader(w: number, h: number): Buffer {
  const sof = Buffer.alloc(10)
  sof.writeUInt16BE(0xffc0, 0) // SOF0 marker
  sof.writeUInt16BE(8, 2) // segment length
  sof.writeUInt8(8, 4) // precision
  sof.writeUInt16BE(h, 5) // height
  sof.writeUInt16BE(w, 7) // width
  return Buffer.concat([Buffer.from([0xff, 0xd8]), sof]) // SOI + SOF0
}

describe('rasterDims: intrinsic dimensions from raw bytes', () => {
  it('reads width/height from a PNG data URI', () => {
    expect(rasterDims(pngDataUri(4, 2))).toEqual({ w: 4, h: 2 })
    expect(rasterDims(pngDataUri(640, 360))).toEqual({ w: 640, h: 360 })
  })

  it('reads width/height from a raw PNG buffer', () => {
    const png = new PNG({ width: 13, height: 7 })
    png.data.fill(0)
    expect(rasterDims(PNG.sync.write(png))).toEqual({ w: 13, h: 7 })
  })

  it('reads width/height from a baseline JPEG', () => {
    expect(rasterDims(jpegHeader(800, 600))).toEqual({ w: 800, h: 600 })
  })

  it('walks intervening JPEG segments to reach the SOF marker', () => {
    // SOI, then an APP0/JFIF segment of length 16, then SOF0 with the real size.
    const app0 = Buffer.concat([
      Buffer.from([0xff, 0xe0]),
      (() => {
        const b = Buffer.alloc(2)
        b.writeUInt16BE(16, 0)
        return b
      })(),
      Buffer.alloc(14),
    ])
    const sof = Buffer.alloc(10)
    sof.writeUInt16BE(0xffc2, 0) // progressive SOF2 — still a frame header
    sof.writeUInt16BE(8, 2)
    sof.writeUInt8(8, 4)
    sof.writeUInt16BE(123, 5)
    sof.writeUInt16BE(456, 7)
    const buf = Buffer.concat([Buffer.from([0xff, 0xd8]), app0, sof])
    expect(rasterDims(buf)).toEqual({ w: 456, h: 123 })
  })

  it('returns null for content it cannot decode', () => {
    expect(rasterDims(Buffer.from('not an image at all, just text'))).toBeNull()
    expect(rasterDims('data:image/gif;base64,R0lGODlhAQABAAAAACwAAAAAAQABAAA=')).toBeNull()
  })
})
