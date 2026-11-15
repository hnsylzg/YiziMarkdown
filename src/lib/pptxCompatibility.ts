/**
 * pptxCompatibility.ts — PPTX 导出的兼容性边界。
 *
 * custom 固定使用 1920×1080 设计画布，正文 30px 按 144 px/in 映射到
 * PowerPoint 后为 15pt。作者 1280 基线的 14pt 下限在这里只允许缩小 6.7%，
 * 无法处理长内容，因此 custom 使用 10pt 作为最低正文。
 */
import { SLIDE_CANVAS_PX_PER_IN } from './slideCanvas.ts'

export const PPTX_BODY_PX = 30
export const PX_PER_IN = SLIDE_CANVAS_PX_PER_IN
export const PPTX_BODY_PT = PPTX_BODY_PX * 72 / PX_PER_IN
/** 正文允许的最小字号（pt） */
export const PPTX_MIN_BODY_PT = 10
/** 压缩系数下限（防止极端设置导致无意义档位） */
export const PPTX_MIN_SCALE_FLOOR = 0.4

/** 基础字号系数（设置里的 slideFontScale）对应的最小压缩系数 */
export function pptxMinFontScale(baseFontScale = 1): number {
  if (!(baseFontScale > 0)) return 1
  const scale = PPTX_MIN_BODY_PT / (PPTX_BODY_PT * baseFontScale)
  return Math.min(1, Math.max(PPTX_MIN_SCALE_FLOOR, scale))
}

/**
 * 可用的压缩档位：从 0.94 逐级向下，过滤掉低于可读下限的档位。
 * 如果列表里没有恰好等于 minScale 的档位，会补上这一档，确保“尽力缩到下限”。
 */
export function pptxShrinkFactors(minScale: number): number[] {
  if (minScale >= 1 - 1e-9) return []
  const base = [0.94, 0.88, 0.82, 0.76, 0.70, 0.64, 0.60]
  const out = base.filter((f) => f >= minScale - 1e-9)
  const last = out[out.length - 1]
  if (last === undefined || last > minScale + 1e-9) out.push(Math.max(PPTX_MIN_SCALE_FLOOR, minScale))
  return out
}
