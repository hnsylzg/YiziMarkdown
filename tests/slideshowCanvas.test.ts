import test from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'

const css = fs.readFileSync(new URL('../src/styles/slideshow.css', import.meta.url), 'utf8')
const slideshow = fs.readFileSync(new URL('../src/components/Slideshow.tsx', import.meta.url), 'utf8')

test('演示模式：1920×1080 固定画布 + 整体缩放', () => {
  assert.match(css, /\.ys-stage-viewport\s*\{[^}]*display:\s*flex;[^}]*overflow:\s*hidden;/s)
  assert.match(css, /\.ys-stage\s*\{[^}]*transform-origin:\s*center center;[^}]*background:\s*var\(--editor-bg, #ffffff\);/s)

  assert.match(slideshow, /className="ys-stage-viewport"/)
  assert.match(slideshow, /className="ys-stage"/)
  assert.match(slideshow, /width:\s*SLIDE_CANVAS_WIDTH,\s*height:\s*SLIDE_CANVAS_HEIGHT/)
  assert.match(slideshow, /slideCanvasScale\(/)
})

test('演示模式：1920 基线下正文字号与标题规格固定', () => {
  assert.match(css, /--ys-u:\s*calc\(30px \* var\(--ys-font-scale, 1\)\)/)
  assert.match(css, /font-size:\s*calc\(30px \* var\(--ys-font-scale, 1\)\)/)
  assert.match(css, /--ys-page-title-size:\s*60px;/)
  assert.match(css, /--ys-title-h1-size:\s*72px;/)
  assert.match(css, /--ys-title-cover-size:\s*115\.2px;/)
})

test('演示模式：幻灯片内容不再依赖真实视口单位（HUD / 备注除外）', () => {
  const slide = css.slice(css.indexOf('.ys-slide {'), css.indexOf('.ys-slide.ys-active'))
    .replace(/\/\*[\s\S]*?\*\//g, '')
  assert.doesNotMatch(slide, /\dvw/)
  assert.doesNotMatch(slide, /\dvh/)
  assert.match(slide, /padding:\s*calc\(540px \* \(1 - var\(--ys-fill-h, 0\.84\)\)\) calc\(960px \* \(1 - var\(--ys-fill-w, 0\.82\)\)\)/)
})
