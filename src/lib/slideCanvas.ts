/** 演示幻灯片与 PPTX 共用的固定 1080p 设计画布。 */
export const SLIDE_CANVAS_WIDTH = 1920
export const SLIDE_CANVAS_HEIGHT = 1080

/** 1920 CSS px 对应 PowerPoint 的 13.333in，即 1px = 0.5pt。 */
export const SLIDE_CANVAS_PX_PER_IN = 144

export function slideCanvasScale(viewportWidth: number, viewportHeight: number): number {
  const width = Number.isFinite(viewportWidth) && viewportWidth > 0 ? viewportWidth : SLIDE_CANVAS_WIDTH
  const height = Number.isFinite(viewportHeight) && viewportHeight > 0 ? viewportHeight : SLIDE_CANVAS_HEIGHT
  return Math.min(width / SLIDE_CANVAS_WIDTH, height / SLIDE_CANVAS_HEIGHT)
}

export function slideCanvasPxToPt(px: number): number {
  return px * 72 / SLIDE_CANVAS_PX_PER_IN
}
