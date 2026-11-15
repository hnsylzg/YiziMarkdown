import test from 'node:test'
import assert from 'node:assert/strict'

import {
  PPTX_BODY_PT,
  PPTX_MIN_BODY_PT,
  pptxMinFontScale,
  pptxShrinkFactors,
} from '../src/lib/pptxCompatibility.ts'

test('PPTX 兼容性：设计正文 15pt，默认最小压缩系数对应 10pt', () => {
  assert.ok(Math.abs(PPTX_BODY_PT - 15) < 1e-9)
  assert.equal(PPTX_MIN_BODY_PT, 10)
  const min = pptxMinFontScale(1)
  assert.ok(Math.abs(min - 10 / 15) < 1e-9)
  assert.ok(Math.abs(PPTX_BODY_PT * min - 10) < 1e-9)
})

test('PPTX 兼容性：用户字号系数越小，可压缩空间越小', () => {
  const normal = pptxMinFontScale(1.2)
  assert.ok(Math.abs(normal - 10 / (15 * 1.2)) < 1e-9)
  assert.ok(normal < pptxMinFontScale(1))

  // 基础字号已经低于 10pt 时不再允许压缩
  assert.equal(pptxMinFontScale(0.6), 1)
})

test('PPTX 兼容性：压缩档位按可读下限过滤，并补上下限档', () => {
  const min = pptxMinFontScale(1)
  const factors = pptxShrinkFactors(min)
  assert.deepEqual(factors.slice(0, 3), [0.94, 0.88, 0.82])
  assert.ok(Math.abs(factors[factors.length - 1] - min) < 1e-9)
  assert.ok(factors.every((f) => f >= min - 1e-9))

  assert.deepEqual(pptxShrinkFactors(1), [])
  assert.deepEqual(pptxShrinkFactors(0.9), [0.94, 0.9])
})
