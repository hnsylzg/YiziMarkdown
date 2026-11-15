/**
 * pptxOverflow.test.ts — 场景节点的超页包围盒回归。
 *
 * 导出自动压缩依赖 `measureSceneOverflow`，它必须同时覆盖顶部/底部/左右四边；
 * 以前的 scrollHeight 实现漏掉居中版式的顶部溢出，这里用纯节点坐标锁定行为。
 */
import assert from 'node:assert/strict'
import fs from 'node:fs'
import test from 'node:test'

import {
  applyCodeFlowExpansion,
  fitCodeRowCards,
  measureSceneOverflow,
  measureSceneTightPx,
  sceneReserveIn,
} from '../src/lib/pptxScene.ts'
import {
  SLIDE_CANVAS_PX_PER_IN,
  slideCanvasPxToPt,
  slideCanvasScale,
} from '../src/lib/slideCanvas.ts'
import type { CardNode, RectNode, SceneNode } from '../src/lib/pptxTypes.ts'

const frameSource = fs.readFileSync(new URL('../src/lib/pptxFrame.ts', import.meta.url), 'utf8')

const W = 13.333
const H = 7.5

test('PPTX 导出：代码卡片不再依赖固定底部内边距', () => {
  assert.doesNotMatch(frameSource, /\.ys-col[^\n]*padding-bottom/)
})

test('PPTX export balances code frame space inside every card', () => {
  const cards: CardNode[] = [
    { t: 'card', x: 1, y: 1, w: 4, h: 3 },
    { t: 'card', x: 5, y: 1, w: 4, h: 3 },
  ]
  const frames: RectNode[][] = [
    [{ t: 'rect', x: 1.2, y: 1.4, w: 3.6, h: 2.8 }],
    [],
  ]
  fitCodeRowCards(cards, frames)
  assert.deepEqual(cards.map((card) => card.h), [3.6, 3.6])
  assert.deepEqual(cards.map((card) => [card.x, card.y, card.w]), [[1, 1, 4], [5, 1, 4]])
  fitCodeRowCards(cards, [[], []])
  assert.deepEqual(cards.map((card) => card.h), [3.6, 3.6])
})

function text(x: number, y: number, w: number, h: number): SceneNode {
  return {
    t: 'text', x, y, w, h,
    align: 'left', valign: 'top', lineSpacingPt: 10,
    paragraphs: [], baseColor: '000000', baseSizePt: 10,
    latin: 'Arial', ea: 'Arial',
  }
}

test('PPTX code expansion moves only following nodes in the same column', () => {
  const above = text(1, 1.5, 4, 0.3)
  const following = text(1, 2.1, 4, 0.3)
  const otherColumn = text(5, 2.1, 4, 0.3)
  applyCodeFlowExpansion(
    [above, following, otherColumn],
    [{ x: 1, w: 4, bottom: 2, extra: 0.3 }],
  )
  assert.equal(above.y, 1.5)
  assert.equal(following.y, 2.4)
  assert.equal(otherColumn.y, 2.1)
})

test('PPTX stacked code expansions accumulate in the same column', () => {
  const following = text(1, 3.2, 4, 0.3)
  applyCodeFlowExpansion(
    [following],
    [
      { x: 1, w: 4, bottom: 2, extra: 0.3 },
      { x: 1, w: 4, bottom: 3, extra: 0.2 },
    ],
  )
  assert.equal(following.y, 3.7)
})

test('PPTX 超页：全部节点在页内时四边都是 0', () => {
  assert.deepEqual(measureSceneOverflow([text(1, 1, 4, 1)], W, H), {
    top: 0, right: 0, bottom: 0, left: 0,
  })
})

test('PPTX 超页：顶部负坐标按真实溢出量计入（居中版式的典型场景）', () => {
  const o = measureSceneOverflow([text(1, -0.25, 4, 1)], W, H)
  assert.equal(o.top, Math.round(0.25 * SLIDE_CANVAS_PX_PER_IN))
  assert.equal(o.bottom, 0)
})

test('PPTX 超页：底部与右侧同时溢出时分别报告', () => {
  const o = measureSceneOverflow([text(12.9, 7.0, 0.6, 0.6)], W, H)
  assert.equal(o.right, Math.round((13.5 - W) * SLIDE_CANVAS_PX_PER_IN))
  assert.equal(o.bottom, Math.round((7.6 - H) * SLIDE_CANVAS_PX_PER_IN))
})

test('PPTX 超页：左侧负坐标计入 left', () => {
  const o = measureSceneOverflow([text(-0.1, 1, 4, 1)], W, H)
  assert.equal(o.left, Math.round(0.1 * SLIDE_CANVAS_PX_PER_IN))
})

test('PPTX 超页：空场景不报溢出', () => {
  assert.deepEqual(measureSceneOverflow([], W, H), { top: 0, right: 0, bottom: 0, left: 0 })
})

test('PPTX 超页：装饰元素（金句引号/章节序号）不参与超页判定', () => {
  const decorative = text(14, 7.6, 1, 1)
  decorative.decorative = true
  assert.deepEqual(measureSceneOverflow([decorative], W, H), {
    top: 0, right: 0, bottom: 0, left: 0,
  })
})

test('PPTX 设计画布：1280 显示只缩放，不改变字号磅值', () => {
  assert.equal(slideCanvasScale(1280, 720), 1)
  assert.equal(slideCanvasScale(1920, 1080), 1.5)
  assert.equal(slideCanvasScale(2560, 1440), 2)
  assert.equal(slideCanvasPxToPt(20), 15)
  assert.equal(slideCanvasPxToPt(48), 36)
  assert.equal(slideCanvasPxToPt(40), 30)
})

test('PPTX 余量：按正文/代码/表格行高预留一行，限制在 0.3~0.5in', () => {
  const t = text(1, 1, 4, 1)
  t.lineSpacingPt = 18
  assert.equal(sceneReserveIn([t]), 0.3)

  t.lineSpacingPt = 24
  assert.equal(sceneReserveIn([t]), 0.333)

  t.lineSpacingPt = 48
  assert.equal(sceneReserveIn([t]), 0.5)

  const decorative = text(1, 1, 4, 1)
  decorative.decorative = true
  assert.equal(sceneReserveIn([decorative]), 0)

  const table: SceneNode = {
    t: 'table', x: 1, y: 1, w: 8, h: 2, colW: [4, 4], rowH: [1, 1],
    rows: [[
      { text: 'a', color: '000000', sizePt: 40, latin: 'Arial', ea: 'Arial', borders: [null, null, null, null], align: 'left', valign: 'top', margin: [0, 0, 0, 0] },
    ]],
  }
  assert.equal(sceneReserveIn([table]), 0.5)
  assert.equal(sceneReserveIn([]), 0)
})

test('PPTX 余量：内容进入页底预留区时返回像素量，纯图形页不预留', () => {
  const t = text(1, 6.4, 4, 1) // 底部 7.4in，距页底 0.1in
  assert.equal(measureSceneTightPx([t], H, 0.3), Math.round(0.2 * SLIDE_CANVAS_PX_PER_IN))

  t.y = 7.0 // 底部 8.0in，真实溢出 0.5in；预留后为 0.8in
  assert.equal(measureSceneTightPx([t], H, 0.3), Math.round(0.8 * SLIDE_CANVAS_PX_PER_IN))

  const decorative = text(1, 7.2, 4, 1)
  decorative.decorative = true
  assert.equal(measureSceneTightPx([decorative], H, 0.3), 0)

  const footer = text(1, 7.2, 4, 1)
  footer.footer = true
  assert.equal(sceneReserveIn([footer]), 0)
  assert.equal(measureSceneTightPx([footer], H, 0.3), 0)

  const image: SceneNode = { t: 'image', x: 1, y: 7.0, w: 4, h: 1, src: 'data:image/png;base64,' }
  assert.equal(measureSceneTightPx([image], H, 0.3), 0)

  assert.equal(measureSceneTightPx([t], H, 0), 0)
  assert.equal(measureSceneTightPx([], H, 0.3), 0)
})
