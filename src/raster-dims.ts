// Intrinsic pixel dimensions from raw raster bytes.
//
// pptxgenjs never measures the images it embeds (see its own `FIXME: Measure
// actual image`), yet its cover/contain crop math needs the image's true aspect
// ratio — otherwise the crop computes to zero and silently degrades to a stretch.
// So when we want a region to *cover*-fit a photo, we decode the header here and
// hand pptxgenjs the real aspect. PNG and baseline/progressive JPEG cover every
// raster this library embeds (Chromium screenshots are PNG; photos are JPEG/PNG).
export interface RasterDims {
  w: number
  h: number
}

const PNG_MAGIC = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])

/** Decode a data: URI or raw image buffer to its intrinsic px size, or null. */
export function rasterDims(input: string | Buffer): RasterDims | null {
  const buf = toBuffer(input)
  if (!buf) return null
  return pngDims(buf) ?? jpegDims(buf)
}

function toBuffer(input: string | Buffer): Buffer | null {
  if (Buffer.isBuffer(input)) return input
  if (input.startsWith('data:')) {
    const comma = input.indexOf(',')
    if (comma < 0) return null
    const isB64 = /;base64/i.test(input.slice(0, comma))
    return isB64 ? Buffer.from(input.slice(comma + 1), 'base64') : Buffer.from(decodeURIComponent(input.slice(comma + 1)))
  }
  return null
}

// PNG: 8-byte signature, then the IHDR chunk whose data begins at byte 16 with
// width then height as big-endian uint32s.
function pngDims(buf: Buffer): RasterDims | null {
  if (buf.length < 24 || !buf.subarray(0, 8).equals(PNG_MAGIC)) return null
  return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) }
}

// JPEG: SOI (FFD8), then a chain of marker segments. A frame header (SOF0..SOFF,
// excluding the non-frame DHT/JPG/DAC markers) carries height then width. Walk
// segment lengths until we hit one.
function jpegDims(buf: Buffer): RasterDims | null {
  if (buf.length < 4 || buf[0] !== 0xff || buf[1] !== 0xd8) return null
  let off = 2
  while (off + 9 < buf.length) {
    if (buf[off] !== 0xff) return null // not aligned on a marker — give up
    const marker = buf[off + 1]
    // SOF markers C0..CF are frame headers, except C4 (DHT), C8 (JPG), CC (DAC).
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      return { h: buf.readUInt16BE(off + 5), w: buf.readUInt16BE(off + 7) }
    }
    const segLen = buf.readUInt16BE(off + 2)
    if (segLen < 2) return null
    off += 2 + segLen
  }
  return null
}
