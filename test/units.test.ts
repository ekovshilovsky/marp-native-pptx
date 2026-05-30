import { describe, it, expect } from 'vitest'
import { pxToPt, pxToInch, scaleFor } from '../src/units.js'

describe('units', () => {
  it('converts px to pt at 96dpi (x0.75)', () => {
    expect(pxToPt(21.333)).toBeCloseTo(16, 1)
    expect(pxToPt(24)).toBe(18)
  })
  it('scales px to inches using a slide scale', () => {
    const s = scaleFor({ w: 1280, h: 720 }, 13.333, 7.5)
    expect(pxToInch(640, s.x)).toBeCloseTo(6.6665, 3) // half width
    expect(pxToInch(360, s.y)).toBeCloseTo(3.75, 3)   // half height
  })
})
