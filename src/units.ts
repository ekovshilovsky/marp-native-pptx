export function pxToPt(px: number): number {
  return px * 0.75
}
export function scaleFor(
  sizePx: { w: number; h: number },
  slideWidthIn: number,
  slideHeightIn: number,
): { x: number; y: number } {
  return { x: slideWidthIn / sizePx.w, y: slideHeightIn / sizePx.h }
}
export function pxToInch(px: number, scaleAxis: number): number {
  return px * scaleAxis
}
