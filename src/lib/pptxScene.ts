/**
 * pptxScene.ts — 把「已经渲染好的演示 DOM」测量成 **场景图**（见 pptxTypes.ts）。
 *
 * 为什么先量再画：PowerPoint 没有 CSS 布局。要让导出与演示逐块对齐，唯一可行的做法
 * 是在固定 1280×720 逻辑画布里让浏览器排一次版，再把每个盒子的位置 / 字号 / 颜色读出来，
 * 映射到 13.333×7.5in。这里**不复制** slideshow.css 的数值，而是读 computed style：
 * 主题、明暗、字号缩放、填充比例、分栏样式改了，场景图自动跟着变。
 *
 * 装饰（标题下划线、章节大序号、金句对角引号、agenda 数字、列表圆点/序号）来自
 * 伪元素：几何（left/top/width/height）、颜色、字号/字重都从 `getComputedStyle(el, '::before'|'::after')`
 * 读，只有**文本内容**（计数器 / attr()）由这里按语义补上 —— 浏览器不会把 counter 解析成字符串。
 *
 * 单位：场景图里 x/y/w/h 是 inch，字号是 pt，颜色是 6 位十六进制（已 flatten 到幻灯片背景上）。
 */
import type { SlideKind } from './slides.ts'
import type { PptxFrame } from './pptxFrame.ts'
import type {
  CardNode, CodeLine, MathNode, PptxWarning, RectNode, SceneNode, SceneOverflow, SceneParagraph, SceneRun, SceneRunStyle,
  SceneTableCell, SlideScene, TextNode,
} from './pptxTypes.ts'
import {
  SLIDE_CANVAS_PX_PER_IN,
  slideCanvasPxToPt,
} from './slideCanvas.ts'
import {
  borderRadiusPx, cleanText, colorAlpha, contentBoxPx, flattenColor, fontSizePx, hasStrike,
  hasUnderline, isBold, letterSpacingPt, lineHeightPx, num, parseBoxShadow, parseColor,
  rectOf, resolveFontStack, round2, round3, styleOf, textAlignOf,
} from './pptxTheme.ts'

export interface SceneOptions {
  /** 1 起的页码 */
  index: number
  kind: SlideKind
  title?: string
  notes: string
  /** 页脚章节名（不含时传空） */
  chapter?: string
  /** 页脚 `3 / 12` */
  pos?: string
  includeFooter: boolean
}

interface Box { x: number; y: number; w: number; h: number }
export interface CodeFlowExpansion {
  x: number
  w: number
  bottom: number
  extra: number
}

const EPS = 2
const SLIDE_W_IN = 40 / 3
const SLIDE_H_IN = 7.5

/** 按代码框实际坐标调整同一排卡片，使代码框上下留白相等。 */
export function fitCodeRowCards(cards: CardNode[], framesByCard: RectNode[][]): void {
  if (!cards.length) return
  const baseH = Math.max(...cards.map((card) => card.h))
  const fittedH = cards.reduce((maxH, card, i) => {
    const frames = framesByCard[i] || []
    if (!frames.length) return maxH
    const topGap = Math.min(...frames.map((frame) => frame.y - card.y))
    const codeBottom = Math.max(...frames.map((frame) => frame.y + frame.h))
    return Math.max(maxH, codeBottom - card.y + topGap)
  }, baseH)
  cards.forEach((card) => {
    card.h = round3(fittedH)
  })
}

function horizontalOverlap(a: { x: number; w: number }, b: { x: number; w: number }): boolean {
  const overlap = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x)
  return overlap > Math.min(a.w, b.w) * 0.25
}

function shiftNodeY(node: SceneNode, delta: number): void {
  if (node.t === 'line') {
    node.y1 += delta
    node.y2 += delta
  } else {
    node.y += delta
  }
}

function nodeFlowBox(node: SceneNode): { x: number; w: number; y: number } {
  if (node.t === 'line') {
    return {
      x: Math.min(node.x1, node.x2),
      w: Math.abs(node.x2 - node.x1),
      y: Math.min(node.y1, node.y2),
    }
  }
  return { x: node.x, w: node.w, y: node.y }
}

/**
 * 导出时为避免横向溢出，代码行可能被额外折行并撑高色块。
 * 浏览器坐标不会知道这次增高，因此同一栏、代码块下方的后续节点必须同步下移。
 */
export function applyCodeFlowExpansion(nodes: SceneNode[], expansions: CodeFlowExpansion[]): void {
  if (!expansions.length) return
  const applied: Array<CodeFlowExpansion & { shiftedBottom: number }> = []
  for (const expansion of [...expansions].sort((a, b) => a.bottom - b.bottom)) {
    const priorShift = applied.reduce((sum, previous) => {
      return previous.bottom <= expansion.bottom + 0.01 &&
        horizontalOverlap(previous, expansion)
        ? sum + previous.extra
        : sum
    }, 0)
    const shiftedBottom = expansion.bottom + priorShift
    for (const node of nodes) {
      const box = nodeFlowBox(node)
      if (box.y >= shiftedBottom - 0.01 && horizontalOverlap(box, expansion)) {
        shiftNodeY(node, expansion.extra)
      }
    }
    applied.push({ ...expansion, shiftedBottom })
  }
}

/**
 * 用场景节点的真实包围盒计算四边溢出（px）。
 *
 * 为什么不继续量 `slide.scrollHeight`：
 *   · 封面/章节/金句等居中版式的内容可以向上溢出到负坐标，scrollHeight 只统计
 *     向下溢出，顶部会被低估（实际溢出 80px 可能只报 40px）；
 *   · scrollHeight 不含子元素 margin，也不含横向溢出；
 *   · 这里按所有最终要落进 PPT 的节点取并集，顶部、底部、左右都能精确报告。
 *
 * 节点坐标是 inch，返回值统一换算成 1280×720 基准的 CSS px。
 */
export function measureSceneOverflow(
  nodes: SceneNode[],
  slideWIn: number,
  slideHIn: number,
): SceneOverflow {
  if (!nodes.length) return { top: 0, right: 0, bottom: 0, left: 0 }
  let minX = Infinity
  let minY = Infinity
  let maxX = -Infinity
  let maxY = -Infinity

  for (const n of nodes) {
    // 装饰元素（金句大引号、章节大序号、agenda 数字）允许贴边或略微出血，
    // 不应因为它们把整页判定成内容超页。
    if (n.t === 'text' && n.decorative) continue
    const x1 = n.t === 'line' ? Math.min(n.x1, n.x2) : n.x
    const y1 = n.t === 'line' ? Math.min(n.y1, n.y2) : n.y
    const x2 = n.t === 'line' ? Math.max(n.x1, n.x2) : n.x + n.w
    const y2 = n.t === 'line' ? Math.max(n.y1, n.y2) : n.y + n.h
    if (![x1, y1, x2, y2].every(Number.isFinite)) continue
    minX = Math.min(minX, x1)
    minY = Math.min(minY, y1)
    maxX = Math.max(maxX, x2)
    maxY = Math.max(maxY, y2)
  }
  if (!Number.isFinite(minX)) return { top: 0, right: 0, bottom: 0, left: 0 }

  const px = (inch: number) => Math.round(Math.max(0, inch) * SLIDE_CANVAS_PX_PER_IN)
  return {
    top: px(-minY),
    right: px(maxX - slideWIn),
    bottom: px(maxY - slideHIn),
    left: px(-minX),
  }
}

/**
 * PowerPoint 折行余量：正文/代码/表格一行的高度，限制在 0.3~0.5in。
 *
 * 导出测量发生在浏览器里，PowerPoint 的字体度量不同，可能把某一行拆成两行。
 * 页底预留一行高度，让“多折一行”不至于直接把内容顶出幻灯片。
 * 纯图形页（没有正文/代码/表格）不预留。
 */
const MIN_RESERVE_IN = 0.3
const MAX_RESERVE_IN = 0.5

export function sceneReserveIn(nodes: SceneNode[]): number {
  let maxLinePt = 0
  for (const n of nodes) {
    if (n.t === 'text' && (n.decorative || n.footer)) continue
    if (n.t === 'text' && !n.decorative) {
      if (Number.isFinite(n.lineSpacingPt)) maxLinePt = Math.max(maxLinePt, n.lineSpacingPt)
    } else if (n.t === 'code') {
      if (Number.isFinite(n.lineSpacingPt)) maxLinePt = Math.max(maxLinePt, n.lineSpacingPt)
    } else if (n.t === 'table') {
      for (const row of n.rows) {
        for (const cell of row) {
          if (Number.isFinite(cell.sizePt)) maxLinePt = Math.max(maxLinePt, cell.sizePt * 1.25)
        }
      }
    }
  }
  if (maxLinePt <= 0) return 0
  return round3(Math.min(MAX_RESERVE_IN, Math.max(MIN_RESERVE_IN, maxLinePt / 72)))
}

/**
 * 内容底部进入预留区的像素数（0 = 距页底仍有足够余量）。
 * 与 `measureSceneOverflow` 的 bottom 不同：这里把“预留的一行”也算进去。
 */
export function measureSceneTightPx(
  nodes: SceneNode[],
  slideHIn: number,
  reserveIn: number,
): number {
  if (!(reserveIn > 0)) return 0
  let maxY = -Infinity
  for (const n of nodes) {
    if (n.t !== 'text' && n.t !== 'code' && n.t !== 'table') continue
    if (n.t === 'text' && (n.decorative || n.footer)) continue
    const y2 = n.y + n.h
    if (!Number.isFinite(y2)) continue
    maxY = Math.max(maxY, y2)
  }
  if (!Number.isFinite(maxY)) return 0
  return Math.round(Math.max(0, maxY + reserveIn - slideHIn) * SLIDE_CANVAS_PX_PER_IN)
}

/** 判断两个普通文本 run 样式是否完全一致；公式 run 永远不能参与文本合并。 */
export function canMergeRunStyles(a: SceneRunStyle, b: SceneRunStyle): boolean {
  return a.color === b.color && a.sizePt === b.sizePt && a.bold === b.bold &&
    a.italic === b.italic && a.underline === b.underline && a.strike === b.strike &&
    a.latin === b.latin && a.ea === b.ea && a.charSpacingPt === b.charSpacingPt &&
    a.highlight === b.highlight && a.link === b.link && !a.mathTex && !b.mathTex
}

/**
 * 代码行按可用宽度折行（按字符，不依赖空格）。
 *
 * PowerPoint 没有“元素内部横向滚动条”，长代码行只能折行。这里在场景层
 * 直接切成多行，避免依赖 PowerPoint 的断词规则（长串下划线/路径通常不会断）。
 */
export function wrapCodeRuns(
  line: SceneRun[],
  availPx: number,
  measure: (ch: string) => number,
): SceneRun[][] {
  const out: SceneRun[][] = []
  let current: SceneRun[] = []
  let width = 0
  const flush = () => {
    if (current.length) out.push(current)
    current = []
    width = 0
  }
  for (const run of line) {
    for (const ch of run.text) {
      const w = measure(ch)
      if (current.length && width + w > availPx) flush()
      const prev = current[current.length - 1]
      if (prev && canMergeRunStyles(prev.style, run.style)) prev.text += ch
      else current.push({ text: ch, style: run.style })
      width += w
    }
  }
  flush()
  return out.length ? out : [line]
}

/** `annotation[encoding="application/x-tex"]` → TeX；KaTeX 的可见 HTML 不可信，必须取源。 */
function katexTex(el: Element): string {
  const ann = el.querySelector('annotation[encoding="application/x-tex"]')
  const tex = (ann?.textContent || '').trim()
  return tex
}

/** 元素的内容盒（文字真正开始的位置，去掉 padding / border） */
function contentBox(el: Element, win: Window): Box {
  const b = contentBoxPx(el, win)
  return { x: b.x, y: b.y, w: b.w, h: b.h }
}

function textBoxOf(el: Element, win: Window): Box {
  const b = contentBox(el, win)
  return { x: b.x, y: b.y, w: b.w, h: b.h + EPS }
}

interface VisualLineSlot {
  top: number
  bottom: number
  center: number
}

/** 读取一组 DOM Range 矩形，按纵向位置合并成浏览器的实际视觉行。 */
function visualLineSlots(root: Element, win: Window): VisualLineSlot[] {
  const range = win.document.createRange()
  try {
    range.selectNodeContents(root)
  } catch {
    return []
  }
  const rects = Array.from(range.getClientRects())
    .filter((r) => r.width > 0.5 && r.height > 0.5)
    .sort((a, b) => a.top - b.top || a.left - b.left)
  const lineHeight = Math.max(1, lineHeightPx(styleOf(root, win)))
  const tolerance = Math.max(4, lineHeight * 0.42)
  const slots: VisualLineSlot[] = []

  for (const rect of rects) {
    const center = rect.top + rect.height / 2
    let slot = slots.find((s) => {
      const overlap = Math.min(s.bottom, rect.bottom) - Math.max(s.top, rect.top)
      return overlap > Math.min(s.bottom - s.top, rect.height) * 0.25 ||
        Math.abs(s.center - center) <= tolerance
    })
    if (!slot) {
      slot = { top: rect.top, bottom: rect.bottom, center }
      slots.push(slot)
    } else {
      slot.top = Math.min(slot.top, rect.top)
      slot.bottom = Math.max(slot.bottom, rect.bottom)
      slot.center = (slot.top + slot.bottom) / 2
    }
  }
  return slots.sort((a, b) => a.top - b.top)
}

/**
 * Range 矩形落在第几个视觉行。
 *
 * 不同字号的 span 会改变矩形高度；先按纵向重叠判断，重叠不足时再看中心距离。
 * `fallback` 用于纯空白字符——它们通常没有自己的 ClientRect。
 */
function lineIndexAtRect(
  rect: { top: number; bottom: number; height: number } | null | undefined,
  slots: VisualLineSlot[],
  fallback: number,
): number {
  if (!rect || !slots.length) return fallback
  const center = rect.top + rect.height / 2
  let best = fallback
  let bestScore = -Infinity
  slots.forEach((slot, idx) => {
    const overlap = Math.min(slot.bottom, rect.bottom) - Math.max(slot.top, rect.top)
    const distance = Math.abs(slot.center - center)
    const score = overlap * 10 - distance
    if (score > bestScore) {
      bestScore = score
      best = idx
    }
  })
  return best
}

function lineIndexAtElement(el: Element, slots: VisualLineSlot[], fallback: number): number {
  const r = rectOf(el)
  return lineIndexAtRect({ top: r.y, bottom: r.y + r.h, height: r.h }, slots, fallback)
}

/** 绝对定位伪元素的盒子：优先用它自己的 left/right/top/bottom + 尺寸。 */
function pseudoBox(host: Element, pcs: CSSStyleDeclaration, w: number, h: number): Box {
  const r = rectOf(host)
  const left = pcs.left !== 'auto' ? num(pcs.left, NaN) : NaN
  const right = pcs.right !== 'auto' ? num(pcs.right, NaN) : NaN
  const top = pcs.top !== 'auto' ? num(pcs.top, NaN) : NaN
  const bottom = pcs.bottom !== 'auto' ? num(pcs.bottom, NaN) : NaN
  const x = Number.isFinite(left) ? r.x + left
    : Number.isFinite(right) ? r.x + r.w - right - w
    : r.x
  const y = Number.isFinite(top) ? r.y + top
    : Number.isFinite(bottom) ? r.y + r.h - bottom - h
    : r.y
  return { x, y, w, h }
}

function pseudoSize(pcs: CSSStyleDeclaration): { w: number; h: number } {
  const w = num(pcs.width, 0)
  const h = num(pcs.height, 0)
  return { w, h }
}

function pseudoColor(pcs: CSSStyleDeclaration, slideBg: string, extraOpacity = 1): string {
  const opacity = extraOpacity * (num(pcs.opacity, 1) || 1)
  return flattenColor(pcs.color || pcs.backgroundColor, slideBg, opacity)
}

function hasVisiblePseudo(pcs: CSSStyleDeclaration): boolean {
  const c = (pcs.content || '').trim()
  return !!c && c !== 'none' && c !== 'normal' && pcs.display !== 'none'
}

/** `linear-gradient(...)` 里的第一个颜色（封面/内容页下划线就是 accent → transparent） */
function gradientFirstColor(backgroundImage: string): string | null {
  if (!backgroundImage || backgroundImage === 'none') return null
  if (!backgroundImage.includes('gradient')) return null
  const m = /(rgba?\([^)]+\)|color\([^)]+\)|#[0-9a-f]{3,8})/i.exec(backgroundImage)
  return m ? m[1] : null
}

/** 伪元素背景：渐变取首个色标，否则取 background-color */
function pseudoFill(pcs: CSSStyleDeclaration, slideBg: string): { color: string; gradient: boolean } | null {
  const grad = gradientFirstColor(pcs.backgroundImage)
  if (grad) return { color: flattenColor(grad, slideBg), gradient: true }
  const bg = parseColor(pcs.backgroundColor)
  if (!bg || bg.a <= 0) return null
  return { color: flattenColor(pcs.backgroundColor, slideBg), gradient: false }
}

const BLOCK_TAGS = new Set([
  'P', 'DIV', 'UL', 'OL', 'LI', 'PRE', 'TABLE', 'BLOCKQUOTE', 'NAV', 'SECTION', 'ASIDE',
  'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'HR', 'FIGURE', 'FIGCAPTION',
])

export function buildSlideScene(frame: PptxFrame, opts: SceneOptions): SlideScene {
  const win = frame.window
  const doc = frame.document
  const slide = frame.slide
  const warnings: PptxWarning[] = []
  const nodes: SceneNode[] = []
  const codeFramesByCol = new Map<Element, RectNode[]>()
  const codeFlowExpansions: CodeFlowExpansion[] = []

  const slideRect = rectOf(slide)
  // 演示与导出共用固定 1280×720 画布；按 96 px/in
  // 映射到 PowerPoint 的 13.333×7.5in，即 1px = 0.75pt。
  const toIn = (px: number) => round3(px / SLIDE_CANVAS_PX_PER_IN)
  const toPt = (px: number) => round2(slideCanvasPxToPt(px))
  const slideCS = styleOf(slide, win)
  const slideBg = flattenColor(slideCS.backgroundColor, '#FFFFFF')
  const padL = num(slideCS.paddingLeft)
  const padR = num(slideCS.paddingRight)

  // 文本宽度测量（金句大引号是文字，Georgia 的“宽度”必须量，不能猜）
  let measureCtx: CanvasRenderingContext2D | null = null
  const measure = (text: string, cs: CSSStyleDeclaration): number => {
    try {
      if (!measureCtx) measureCtx = document.createElement('canvas').getContext('2d')
      if (!measureCtx) return fontSizePx(cs) * 0.6
      measureCtx.font = `${cs.fontStyle} ${cs.fontWeight} ${fontSizePx(cs)}px ${cs.fontFamily}`
      return measureCtx.measureText(text).width
    } catch {
      return fontSizePx(cs) * 0.6
    }
  }

  // ---------- 样式 ----------

  const runStyleOf = (el: Element, bg = slideBg): SceneRunStyle => {
    const cs = styleOf(el, win)
    const font = resolveFontStack(cs.fontFamily)
    const isCode = el.tagName === 'CODE' && !el.closest('pre')
    const bgColor = parseColor(cs.backgroundColor)
    const highlight = isCode && bgColor && bgColor.a > 0
      ? flattenColor(cs.backgroundColor, bg)
      : undefined
    const anchor = el.tagName === 'A' ? el as HTMLAnchorElement : null
    const href = anchor?.getAttribute('href') || ''
    return {
      color: flattenColor(cs.color, bg),
      sizePt: toPt(fontSizePx(cs)),
      bold: isBold(cs) || undefined,
      italic: /italic|oblique/i.test(cs.fontStyle) || undefined,
      underline: hasUnderline(cs) || undefined,
      strike: hasStrike(cs) || undefined,
      latin: font.latin,
      ea: font.ea,
      charSpacingPt: letterSpacingPt(cs),
      highlight,
      link: /^https?:\/\//i.test(href) ? href : undefined,
    }
  }

  const trimRun = (r: SceneRunStyle): SceneRunStyle => {
    const out: SceneRunStyle = { ...r }
    if (!out.bold) delete out.bold
    if (!out.italic) delete out.italic
    if (!out.underline) delete out.underline
    if (!out.strike) delete out.strike
    if (out.charSpacingPt === undefined) delete out.charSpacingPt
    if (!out.highlight) delete out.highlight
    if (!out.link) delete out.link
    return out
  }

  // ---------- 行内文本收集 ----------

  const collectInline = (
    root: Element,
    skipNestedBlocks: boolean,
    bg = slideBg,
  ): { paragraphs: SceneParagraph[]; visualLines?: SceneRun[][] } => {
    const paragraphs: SceneParagraph[] = []
    const lineMap = new Map<number, SceneRun[]>()
    const lineSlots = visualLineSlots(root, win)
    let current: SceneRun[] = []
    let lastVisualLine = 0

    const appendRun = (target: SceneRun[], text: string, style: SceneRunStyle) => {
      if (!text) return
      const prev = target[target.length - 1]
      if (prev && canMergeRunStyles(prev.style, style)) {
        prev.text += text
        return
      }
      target.push({ text, style })
    }

    const flush = () => {
      const runs = current
      current = []
      if (!runs.length) return
      // 去掉块首/块尾空白（HTML 的软换行不进 PPT）
      runs[0].text = runs[0].text.replace(/^ +/, '')
      const last = runs[runs.length - 1]
      last.text = last.text.replace(/ +$/, '')
      const kept = runs.filter((r) => r.text !== '' || r.style.mathTex)
      if (kept.length) paragraphs.push({ runs: kept })
    }

    const push = (
      text: string,
      el: Element,
      extra?: Partial<SceneRunStyle>,
      lineIndex = lastVisualLine,
    ) => {
      if (!text) return
      const style = trimRun({ ...runStyleOf(el, bg), ...extra })
      appendRun(current, text, style)
      const line = lineMap.get(lineIndex) || []
      appendRun(line, text, style)
      lineMap.set(lineIndex, line)
      lastVisualLine = lineIndex
    }

    const pushTextNode = (node: Text, parentEl: Element) => {
      const raw = node.nodeValue || ''
      if (!raw) return
      const range = win.document.createRange()
      range.selectNodeContents(node)
      const rects = Array.from(range.getClientRects())
        .filter((r) => r.width > 0.5 && r.height > 0.5)

      if (rects.length <= 1) {
        const line = lineIndexAtRect(rects[0], lineSlots, lastVisualLine)
        push(cleanText(raw), parentEl, undefined, line)
        return
      }

      // 一个文本节点可能跨多行。逐字符归入浏览器视觉行，供表格等需要固定
      // 换行的场景使用；普通正文不改写视觉行，由 PowerPoint 自动换行。
      let buffer = ''
      let line = lineIndexAtRect(rects[0], lineSlots, lastVisualLine)
      for (let i = 0; i < raw.length;) {
        const cp = raw.codePointAt(i)
        const size = cp !== undefined && cp > 0xffff ? 2 : 1
        range.setStart(node, i)
        range.setEnd(node, i + size)
        const rect = range.getClientRects()[0]
        const nextLine = lineIndexAtRect(rect, lineSlots, line)
        if (nextLine !== line && buffer) {
          push(cleanText(buffer), parentEl, undefined, line)
          buffer = ''
        }
        line = nextLine
        buffer += raw.slice(i, i + size)
        i += size
      }
      if (buffer) push(cleanText(buffer), parentEl, undefined, line)
    }

    const visit = (node: Node, parentEl: Element) => {
      if (node.nodeType === 3) {
        pushTextNode(node as Text, parentEl)
        return
      }
      if (node.nodeType !== 1) return
      const el = node as Element
      const tag = el.tagName
      if (tag === 'SCRIPT' || tag === 'STYLE') return
      const cs = styleOf(el, win)
      if (cs.display === 'none' || cs.visibility === 'hidden') return
      if (el.classList.contains('katex')) {
        const tex = katexTex(el)
        if (tex) {
          // KaTeX 的 `.katex` 自带 `font-size: 1.21em`（它的字体按这个比例设计），
          // 直接取它会比正文大 21%。字号取**外层行内上下文**，才是 PPT 里该用的值。
          const host = el.parentElement || el
          const style = trimRun({ ...runStyleOf(host, bg), mathTex: tex })
          const line = lineIndexAtElement(el, lineSlots, lastVisualLine)
          appendRun(current, `$${tex}$`, style)
          const lineRuns = lineMap.get(line) || []
          appendRun(lineRuns, `$${tex}$`, style)
          lineMap.set(line, lineRuns)
          lastVisualLine = line
        } else {
          push(cleanText(el.textContent || ''), el)
        }
        return
      }
      if (tag === 'BR') { flush(); return }
      if (tag === 'IMG') return
      if (tag === 'INPUT') {
        const input = el as HTMLInputElement
        if (input.type === 'checkbox') {
          const accent = flattenColor(win.getComputedStyle(el).accentColor || '#4f46e5', bg)
          const line = lineIndexAtElement(el, lineSlots, lastVisualLine)
          push(input.checked ? '☑' : '☐', el, {
            color: input.checked ? accent : flattenColor(styleOf(el, win).color, bg),
            checkbox: true,
          }, line)
        }
        return
      }
      if (skipNestedBlocks && BLOCK_TAGS.has(tag) && el !== root) return
      const children = el.childNodes
      for (let i = 0; i < children.length; i++) visit(children[i], el)
    }

    const children = root.childNodes
    for (let i = 0; i < children.length; i++) visit(children[i], root)
    flush()

    const visualLines: SceneRun[][] = []
    for (const [, runs] of Array.from(lineMap.entries()).sort((a, b) => a[0] - b[0])) {
      if (!runs.length) continue
      runs[0].text = runs[0].text.replace(/^ +/, '')
      runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/ +$/, '')
      const kept = runs.filter((r) => r.text !== '' || r.style.mathTex)
      if (kept.length) visualLines.push(kept)
    }
    return { paragraphs, visualLines: visualLines.length ? visualLines : undefined }
  }

  // ---------- 节点 ----------

  const pushText = (
    box: Box,
    paragraphs: SceneParagraph[],
    cs: CSSStyleDeclaration,
    opt?: {
      align?: 'left' | 'center' | 'right'
      decorative?: boolean
      noWrap?: boolean
      bg?: string
      visualLines?: SceneRun[][]
      footer?: boolean
    },
  ) => {
    if (!paragraphs.length) return
    const font = resolveFontStack(cs.fontFamily)
    const bg = opt?.bg || slideBg
    const node: TextNode = {
      t: 'text',
      x: toIn(box.x),
      y: toIn(box.y),
      w: toIn(Math.max(1, box.w)),
      h: toIn(Math.max(1, box.h)),
      align: opt?.align || textAlignOf(cs),
      valign: 'top',
      lineSpacingPt: toPt(lineHeightPx(cs)),
      paragraphs,
      baseColor: flattenColor(cs.color, bg),
      baseSizePt: toPt(fontSizePx(cs)),
      latin: font.latin,
      ea: font.ea,
      visualLines: opt?.visualLines,
      decorative: opt?.decorative || undefined,
      footer: opt?.footer || undefined,
      noWrap: opt?.noWrap || undefined,
    }
    nodes.push(node)
  }

  const emitTextBlock = (el: Element) => {
    const cs = styleOf(el, win)
    const { paragraphs, visualLines } = collectInline(el, true)
    if (!paragraphs.length) return
    pushText(textBoxOf(el, win), paragraphs, cs, { align: textAlignOf(cs), visualLines })
  }

  const emitImage = (img: HTMLImageElement) => {
    const r = rectOf(img)
    if (r.w < 1 || r.h < 1) return
    const src = img.currentSrc || img.src || img.getAttribute('src') || ''
    if (!src.startsWith('data:')) {
      warnings.push({
        kind: 'image',
        messageKey: 'pptxExport.warnImageNotEmbedded',
        messageArgs: { src: (img.getAttribute('src') || '').slice(0, 80) },
      })
      return
    }
    const cs = styleOf(img, win)
    // Mermaid 栅格化后的 img 是导出侧产物：演示里它是 svg，本来就没有主题给图片加的阴影/圆角。
    const fromMermaid = img.hasAttribute('data-ys-mermaid')
    const shadow = fromMermaid ? undefined : parseBoxShadow(cs)
    nodes.push({
      t: 'image',
      x: toIn(r.x), y: toIn(r.y), w: toIn(r.w), h: toIn(r.h),
      src,
      radiusPt: toPt(borderRadiusPx(cs)),
      shadow,
      alt: img.getAttribute('alt') || undefined,
    })
  }

  const emitHr = (el: Element) => {
    const cs = styleOf(el, win)
    const r = rectOf(el)
    const width = Math.max(1, num(cs.borderTopWidth, 1))
    nodes.push({
      t: 'rect',
      x: toIn(r.x), y: toIn(r.y), w: toIn(r.w), h: toIn(width),
      fill: flattenColor(cs.borderTopColor, slideBg),
    })
  }

  const emitCard = (el: Element, box?: Box) => {
    const cs = styleOf(el, win)
    const bg = parseColor(cs.backgroundColor)
    const borderW = Math.max(
      num(cs.borderTopWidth), num(cs.borderRightWidth), num(cs.borderBottomWidth), num(cs.borderLeftWidth),
    )
    const shadow = parseBoxShadow(cs)
    const radius = borderRadiusPx(cs)
    if ((!bg || bg.a <= 0) && borderW <= 0 && !shadow) return
    const r = box || rectOf(el)
    nodes.push({
      t: 'card',
      x: toIn(r.x), y: toIn(r.y), w: toIn(r.w), h: toIn(r.h),
      fill: bg && bg.a > 0 ? flattenColor(cs.backgroundColor, slideBg) : undefined,
      lineColor: borderW > 0 ? flattenColor(cs.borderTopColor, slideBg) : undefined,
      linePt: borderW > 0 ? toPt(borderW) : undefined,
      radiusPt: toPt(radius),
      shadow,
    })
  }

  const emitBoxDecoration = (el: Element) => {
    const cs = styleOf(el, win)
    const r = rectOf(el)
    const sides: Array<'Top' | 'Right' | 'Bottom' | 'Left'> = ['Top', 'Right', 'Bottom', 'Left']
    for (const side of sides) {
      const w = num((cs as any)[`border${side}Width`])
      if (w < 0.5) continue
      const color = flattenColor((cs as any)[`border${side}Color`] || cs.borderTopColor, slideBg)
      if (side === 'Left' || side === 'Right') {
        const x = side === 'Left' ? r.x : r.x + r.w - w
        nodes.push({ t: 'rect', x: toIn(x), y: toIn(r.y), w: toIn(w), h: toIn(r.h), fill: color })
      } else {
        const y = side === 'Top' ? r.y : r.y + r.h - w
        nodes.push({ t: 'rect', x: toIn(r.x), y: toIn(y), w: toIn(r.w), h: toIn(w), fill: color })
      }
    }
  }

  const emitCode = (pre: Element) => {
    const cs = styleOf(pre, win)
    const r = rectOf(pre)
    const box = contentBox(pre, win)
    const code = pre.querySelector('code') || pre
    const codeCS = styleOf(code, win)

    const lines: CodeLine[] = [[]]
    const pushRun = (text: string, el: Element) => {
      if (!text) return
      const t = text.replace(/\r/g, '')
      if (!t) return
      const parts = t.split('\n')
      parts.forEach((part, i) => {
        if (i > 0) lines.push([])
        if (!part) return
        const style = trimRun(runStyleOf(el, slideBg))
        const line = lines[lines.length - 1]
        const prev = line[line.length - 1]
        if (prev && canMergeRunStyles(prev.style, style)) prev.text += part
        else line.push({ text: part, style })
      })
    }
    const visit = (node: Node, parentEl: Element) => {
      if (node.nodeType === 3) { pushRun(node.nodeValue || '', parentEl); return }
      if (node.nodeType !== 1) return
      const el = node as Element
      if (el.tagName === 'BR') { lines.push([]); return }
      const children = el.childNodes
      for (let i = 0; i < children.length; i++) visit(children[i], el)
    }
    visit(code, code)
    if (lines.length > 1) {
      const last = lines[lines.length - 1]
      if (!last.length) lines.pop()
    }
    lines.forEach((line) => {
      if (!line.length) line.push({ text: ' ', style: trimRun(runStyleOf(code, slideBg)) })
    })

    const fontPx = fontSizePx(codeCS)
    const linePx = lineHeightPx(codeCS)
    let measureChar: (ch: string) => number
    try {
      const ctx = win.document.createElement('canvas').getContext('2d')
      if (ctx) {
        ctx.font = `${fontPx}px ${codeCS.fontFamily}`
        const probe = ctx.measureText('M').width
        const fontOk = probe > fontPx * 0.2 && probe < fontPx * 1.6
        const fallback = (ch: string) => {
          const cp = ch.codePointAt(0) ?? 0
          return cp >= 0x2e80 ? fontPx : fontPx * 0.6
        }
        const cache = new Map<string, number>()
        measureChar = (ch) => {
          if (!fontOk) return fallback(ch)
          let w = cache.get(ch)
          if (w === undefined) {
            w = ctx.measureText(ch).width
            if (!(w > 0) || w > fontPx * 2) w = fallback(ch)
            cache.set(ch, w)
          }
          return w
        }
      } else {
        measureChar = (_ch) => fontPx * 0.6
      }
    } catch {
      measureChar = (_ch) => fontPx * 0.6
    }

    // PowerPoint 没有元素内部横向滚动条：长代码行按内容盒宽度折行，
    // 并同步放大代码块高度（后续内容可能因此进入页底预留区 / 红区）。
    const availPx = Math.max(1, box.w - 8)
    const wrappedLines = lines.flatMap((line) => wrapCodeRuns(line, availPx, measureChar))
    const extraPx = Math.max(0, wrappedLines.length - lines.length) * linePx
    if (extraPx > 0) {
      codeFlowExpansions.push({
        x: toIn(r.x),
        w: toIn(r.w),
        bottom: toIn(r.y + r.h),
        extra: toIn(extraPx),
      })
    }
    // 不再横向缩小代码：缩到 60% 仍放不下的长行改为折行，避免导出后横向溢出。

    const font = resolveFontStack(codeCS.fontFamily)
    // 背景/边框按**边框盒**（含 pre 的内边距）单独画，文本仍按内容盒排 ——
    // PowerPoint 的精确行距会把首行整体下移约半个行距，若背景也只覆盖内容盒，
    // 末行（尤其收尾的 `}`）会露到色块外。演示里的色块本来就是含 padding 的。
    const fill = (parseColor(cs.backgroundColor)?.a || 0) > 0 ? flattenColor(cs.backgroundColor, slideBg) : undefined
    const linePt = num(cs.borderTopWidth) > 0 ? toPt(num(cs.borderTopWidth)) : undefined
    const frame: RectNode = {
      t: 'rect',
      x: toIn(r.x), y: toIn(r.y), w: toIn(r.w), h: toIn(r.h + extraPx),
      fill,
      lineColor: linePt ? flattenColor(cs.borderTopColor, slideBg) : undefined,
      linePt,
      radiusPt: toPt(borderRadiusPx(cs)),
    }
    if (fill || linePt) nodes.push(frame)
    const col = pre.closest('.ys-col')
    if (col) {
      const frames = codeFramesByCol.get(col) || []
      frames.push(frame)
      codeFramesByCol.set(col, frames)
    }
    nodes.push({
      t: 'code',
      x: toIn(box.x), y: toIn(box.y), w: toIn(box.w), h: toIn(box.h + extraPx),
      lines: wrappedLines,
      sizePt: toPt(fontSizePx(codeCS)),
      lineSpacingPt: toPt(lineHeightPx(codeCS)),
      latin: font.latin,
      ea: font.ea,
      scale: 1,
    })
  }

  const cellFill = (cell: Element, row: Element): string | undefined => {
    for (const el of [cell, row]) {
      const cs = styleOf(el, win)
      const c = parseColor(cs.backgroundColor)
      if (c && c.a > 0.01) return flattenColor(cs.backgroundColor, slideBg)
    }
    return undefined
  }

  const emitTable = (table: Element) => {
    const rows = Array.from(table.querySelectorAll('tr'))
    if (!rows.length) return
    const tableRect = rectOf(table)
    const colWidths: number[] = []
    const rowHeights: number[] = []
    rows.forEach((row) => {
      rowHeights.push(row.getBoundingClientRect().height)
      Array.from(row.children).forEach((cell, idx) => {
        const w = cell.getBoundingClientRect().width
        colWidths[idx] = Math.max(colWidths[idx] || 0, w)
      })
    })
    const sceneRows: SceneTableCell[][] = rows.map((row) => {
      const cells = Array.from(row.children)
      return cells.map((cell) => {
        const cs = styleOf(cell, win)
        const font = resolveFontStack(cs.fontFamily)
        const borders: SceneTableCell['borders'] = []
        for (const side of ['Top', 'Right', 'Bottom', 'Left'] as const) {
          const w = num((cs as any)[`border${side}Width`])
          borders.push(w > 0.3
            ? { color: flattenColor((cs as any)[`border${side}Color`] || cs.borderTopColor, slideBg), pt: toPt(w) }
            : null)
        }
        const { paragraphs, visualLines } = collectInline(cell, true)
        // 表格单元格同样保留浏览器视觉行，避免 PowerPoint 用不同字体度量把
        // 单行重新折成两行，撑高行高后把后续行挤出页面。
        const text = (visualLines || paragraphs.map((p) => p.runs))
          .map((line) => line.map((r) => r.text).join(''))
          .join('\n')
        return {
          text,
          bold: isBold(cs) || undefined,
          color: flattenColor(cs.color, slideBg),
          sizePt: toPt(fontSizePx(cs)),
          latin: font.latin,
          ea: font.ea,
          fill: cellFill(cell, row),
          borders,
          align: textAlignOf(cs),
          valign: cs.verticalAlign === 'middle' ? 'middle' : cs.verticalAlign === 'bottom' ? 'bottom' : 'top',
          margin: [
            toIn(num(cs.paddingTop)),
            toIn(num(cs.paddingRight)),
            toIn(num(cs.paddingBottom)),
            toIn(num(cs.paddingLeft)),
          ],
          header: cell.tagName === 'TH' || undefined,
        }
      })
    })
    nodes.push({
      t: 'table',
      x: toIn(tableRect.x), y: toIn(tableRect.y),
      w: toIn(colWidths.reduce((a, b) => a + b, 0) || tableRect.w),
      h: toIn(rowHeights.reduce((a, b) => a + b, 0) || tableRect.h),
      colW: colWidths.map(toIn),
      rowH: rowHeights.map(toIn),
      rows: sceneRows,
    })
  }

  const emitMathDisplay = (display: Element) => {
    const r = rectOf(display)
    // 同样避开 `.katex` 的 1.21em：display 容器继承正文的 font-size
    const cs = styleOf(display, win)
    const tex = katexTex(display)
    if (!tex) return
    const font = resolveFontStack(cs.fontFamily)
    const node: MathNode = {
      t: 'math',
      x: toIn(r.x), y: toIn(r.y), w: toIn(r.w), h: toIn(r.h),
      tex,
      sizePt: toPt(fontSizePx(cs)),
      color: flattenColor(cs.color, slideBg),
      latin: font.latin,
      ea: font.ea,
    }
    nodes.push(node)
  }

  const isOnlyDisplayMath = (el: Element, display: Element): boolean => {
    const kids = Array.from(el.childNodes)
    return kids.every((n) =>
      (n.nodeType === 3 && !(n.nodeValue || '').trim()) ||
      (n.nodeType === 1 && (n === display || (n as Element).contains(display)))
    )
  }

  // ---------- 列表 ----------

  const emitList = (list: Element, ordered: boolean, agenda: boolean) => {
    const lis = Array.from(list.children).filter((c) => c.tagName === 'LI')
    lis.forEach((li, i) => {
      if (styleOf(li, win).display === 'none') return
      const isTask = li.classList.contains('task-list-item')
      const liCS = styleOf(li, win)
      const liBox = contentBox(li, win)

      if (!agenda && !isTask) {
        const pcs = styleOf(li, win, '::before')
        if (hasVisiblePseudo(pcs) && (pcs.content || '').trim() !== 'none') {
          if (ordered) {
            const text = `${i + 1}.`
            const w = Math.max(num(pcs.minWidth, 0), liBox.x - rectOf(li).x + 1)
            const h = Math.max(fontSizePx(pcs), lineHeightPx(pcs))
            const box = pseudoBox(li, pcs, w, h)
            const font = resolveFontStack(pcs.fontFamily || liCS.fontFamily)
            const style: SceneRunStyle = {
              color: pseudoColor(pcs, slideBg),
              sizePt: toPt(fontSizePx(pcs)),
              bold: isBold(pcs) || undefined,
              latin: font.latin,
              ea: font.ea,
            }
            // 伪元素的 left/right 决定对齐方向
            const align = pcs.right !== 'auto' && pcs.left === 'auto' ? 'right' : 'left'
            pushText(box, [{ runs: [{ text, style }] }], pcs, { align })
          } else {
            const size = pseudoSize(pcs)
            if (size.w > 0.5 && size.h > 0.5) {
              const box = pseudoBox(li, pcs, size.w, size.h)
              nodes.push({
                t: 'ellipse',
                x: toIn(box.x), y: toIn(box.y), w: toIn(box.w), h: toIn(box.h),
                fill: flattenColor(pcs.backgroundColor, slideBg),
              })
            }
          }
        }
      }

      if (agenda) {
        const pcs = styleOf(li, win, '::before')
        const w = Math.max(liBox.x - rectOf(li).x, fontSizePx(pcs) * 1.6)
        // 盒高必须 >= 实际行高：伪元素是 1.6em 大字号，若行距大于字号，
        // 只按 font-size 给高度会让大数字在 PPT 里被裁掉半截。
        const h = Math.max(fontSizePx(pcs), lineHeightPx(pcs), 1)
        const box = pseudoBox(li, pcs, w * 0.9, h)
        const font = resolveFontStack(pcs.fontFamily || liCS.fontFamily)
        const style: SceneRunStyle = {
          color: pseudoColor(pcs, slideBg),
          sizePt: toPt(fontSizePx(pcs)),
          bold: isBold(pcs) || undefined,
          latin: font.latin,
          ea: font.ea,
        }
        pushText(box, [{ runs: [{ text: String(i + 1), style }] }], pcs, { align: textAlignOf(pcs), decorative: true })
      }

      // 列表项自身内容（跳过嵌套列表，它们单独排版）。
      // 松散列表会生成 `<li><p>…</p></li>`：此时按块渲染，否则 collectInline 会把 p 整块跳过。
      const ownBlocks = Array.from(li.children).filter((c) =>
        c.tagName === 'P' || c.tagName === 'PRE' || c.tagName === 'TABLE' || c.tagName === 'BLOCKQUOTE')
      if (ownBlocks.length) {
        ownBlocks.forEach((child) => emitBlock(child))
      } else {
        const { paragraphs, visualLines } = collectInline(li, true)
        if (paragraphs.length) {
          pushText(textBoxOf(li, win), paragraphs, liCS, { align: textAlignOf(liCS), visualLines })
        }
      }

      // 嵌套列表
      Array.from(li.children).forEach((child) => {
        if (child.tagName === 'UL') emitList(child, false, false)
        else if (child.tagName === 'OL') emitList(child, true, false)
      })
    })
  }

  // ---------- 走块 ----------

  const emitChildrenAsBlocks = (el: Element) => {
    Array.from(el.children).forEach((child) => emitBlock(child))
  }

  const emitBlock = (el: Element) => {
    const tag = el.tagName
    const cls = el.classList
    const cs = styleOf(el, win)
    if (cs.display === 'none' || cs.visibility === 'hidden') return

    if (tag === 'HR') { emitHr(el); return }

    if (tag === 'PRE') { emitCode(el); return }

    if (tag === 'TABLE') { emitTable(el); return }

    if (tag === 'IMG') { emitImage(el as HTMLImageElement); return }

    if (cls.contains('katex-display')) { emitMathDisplay(el); return }

    if (tag === 'P') {
      const display = el.querySelector('.katex-display')
      if (display && isOnlyDisplayMath(el, display)) { emitMathDisplay(display); return }
      const imgs = Array.from(el.querySelectorAll('img'))
      if (imgs.length === 1 && isOnlyDisplayMath(el, imgs[0])) { emitImage(imgs[0] as HTMLImageElement); return }
      const onlyImg = el.childNodes.length === 1 && el.firstElementChild?.tagName === 'IMG'
      if (onlyImg) { emitImage(el.firstElementChild as HTMLImageElement); return }
      emitTextBlock(el)
      return
    }

    if (tag === 'UL' || tag === 'OL') {
      emitList(el, tag === 'OL', !!el.closest('.ys-layout-agenda'))
      return
    }

    if (cls.contains('markdown-toc')) {
      emitCard(el)
      // TOC 的条目没有项目符号（CSS content:none）；作为普通文本逐条渲染
      const lis = Array.from(el.querySelectorAll('li'))
      lis.forEach((li) => {
        const { paragraphs, visualLines } = collectInline(li, true)
        const liStyle = styleOf(li, win)
        if (paragraphs.length) {
          pushText(textBoxOf(li, win), paragraphs, liStyle, {
            align: textAlignOf(liStyle),
            visualLines,
          })
        }
      })
      return
    }

    if (cls.contains('ys-cols')) {
      const cols = Array.from(el.children).filter((c) => c.classList.contains('ys-col'))
      const rowCards: CardNode[] = []
      const cardCols: Element[] = []
      // 卡片：先铺底再放内容
      const style = styleOf(el, win)
      const lineMode = !!slide.closest('.yizi-slideshow')?.classList.contains('ys-col-style-line') ||
        !!el.closest('.ys-col-style-line')
      if (lineMode) {
        cols.forEach((col, idx) => {
          if (idx === 0) return
          const pcs = styleOf(col, win, '::before')
          const w = Math.max(1, num(pcs.width, 1))
          const r = rectOf(col)
          const left = pcs.left !== 'auto' ? num(pcs.left) : -18
          nodes.push({
            t: 'rect',
            x: toIn(r.x + left), y: toIn(r.y), w: toIn(w), h: toIn(r.h),
            fill: flattenColor(pcs.backgroundColor || style.borderTopColor, slideBg),
          })
        })
      }
      cols.forEach((col) => {
        const cardIndex = nodes.length
        emitCard(col)
        const card = nodes[cardIndex]
        if (card?.t === 'card') {
          rowCards.push(card)
          cardCols.push(col)
        }
        Array.from(col.children).forEach((child) => emitBlock(child))
      })
      fitCodeRowCards(rowCards, cardCols.map((col) => codeFramesByCol.get(col) || []))
      return
    }

    if (tag === 'BLOCKQUOTE') {
      emitBoxDecoration(el)
      const hasBlockChild = Array.from(el.children).some((c) => BLOCK_TAGS.has(c.tagName))
      if (hasBlockChild) {
        // 子块（p / 列表 / 代码…）各自成文本块。容器自身若还有裸文本，
        // 才额外补一个文本块；否则 collectInline 会把子块里的文字再输出一遍。
        Array.from(el.children).forEach((child) => {
          if (BLOCK_TAGS.has(child.tagName)) emitBlock(child)
        })
        const own = Array.from(el.childNodes).some((n) => n.nodeType === 3 && (n.nodeValue || '').trim())
        if (own) {
          const { paragraphs, visualLines } = collectInline(el, true)
          if (paragraphs.length) {
            pushText(textBoxOf(el, win), paragraphs, cs, { align: textAlignOf(cs), visualLines })
          }
        }
      } else {
        // 只有行内元素（常见：`> 引用` 单段）⇒ 整块作为一个文本块输出
        const { paragraphs, visualLines } = collectInline(el, true)
        if (paragraphs.length) {
          pushText(textBoxOf(el, win), paragraphs, cs, { align: textAlignOf(cs), visualLines })
        }
      }
      return
    }

    if (tag === 'DIV' || tag === 'SECTION' || tag === 'NAV' || tag === 'ASIDE' || tag === 'FIGURE') {
      // 容器里只有一个 <img> —— 典型是 Mermaid 栅格化后的 `.mermaid > img`：
      // 它既没有块级子元素也没有文本，`collectInline` 会跳过 IMG，不特判就会整块丢失。
      if (el.childNodes.length === 1 && el.firstElementChild?.tagName === 'IMG') {
        emitImage(el.firstElementChild as HTMLImageElement)
        return
      }
      // 有可见背景 / 边框的容器先画成卡片（分栏卡片、目录块、脚注块……）
      emitCard(el)
      const hasBlockChild = Array.from(el.children).some((c) => BLOCK_TAGS.has(c.tagName))
      if (hasBlockChild) {
        emitChildrenAsBlocks(el)
        // 容器自身的纯文本（rare）：只取不在子块里的文本
        const own = Array.from(el.childNodes).some((n) => n.nodeType === 3 && (n.nodeValue || '').trim())
        if (own) {
          const { paragraphs, visualLines } = collectInline(el, true)
          if (paragraphs.length) {
            pushText(textBoxOf(el, win), paragraphs, cs, { align: textAlignOf(cs), visualLines })
          }
        }
      } else {
        const { paragraphs, visualLines } = collectInline(el, true)
        if (paragraphs.length) {
          pushText(textBoxOf(el, win), paragraphs, cs, { align: textAlignOf(cs), visualLines })
        }
      }
      return
    }

    // 标题 / 其它行内块
    emitTextBlock(el)
  }

  // ---------- 装饰（先画：文字压在上面） ----------

  // 1) 章节大序号（伪元素在流内，不能直接量；用首块位置反推）
  if (opts.kind === 'section' && slide.hasAttribute('data-num')) {
    const pcs = styleOf(slide, win, '::before')
    if (hasVisiblePseudo(pcs)) {
      const first = slide.firstElementChild
      const fsz = fontSizePx(pcs)
      const lh = lineHeightPx(pcs)
      const mb = num(pcs.marginBottom)
      const firstTop = first ? rectOf(first).y : slideRect.y + num(slideCS.paddingTop)
      const box: Box = {
        x: slideRect.x + padL,
        y: firstTop - mb - lh,
        w: Math.max(1, slideRect.w - padL - padR),
        h: lh,
      }
      const font = resolveFontStack(pcs.fontFamily)
      const style: SceneRunStyle = {
        color: pseudoColor(pcs, slideBg),
        sizePt: toPt(fsz),
        bold: isBold(pcs) || undefined,
        latin: font.latin,
        ea: font.ea,
      }
      pushText(box, [{ runs: [{ text: slide.getAttribute('data-num') || '', style }] }], pcs, {
        align: textAlignOf(pcs), decorative: true,
      })
    }
  }

  // 2) 标题下划线（内容页 h2 / 封面·结尾页首块；渐变用 objectName 标给 JSZip 后处理）
  Array.from(slide.querySelectorAll('h1, h2, h3, h4')).forEach((h) => {
    const pcs = styleOf(h, win, '::after')
    if (!hasVisiblePseudo(pcs)) return
    const fill = pseudoFill(pcs, slideBg)
    if (!fill) return
    const size = pseudoSize(pcs)
    if (size.w < 1 || size.h < 1) return
    const r = rectOf(h)
    let box: Box
    if (pcs.position === 'absolute') {
      box = pseudoBox(h, pcs, size.w, size.h)
    } else {
      // 封面 / 结尾页的下划线是块级伪元素：margin auto 居中，贴在元素底部
      box = { x: r.x + (r.w - size.w) / 2, y: r.y + r.h - size.h, w: size.w, h: size.h }
    }
    nodes.push({
      t: 'rect',
      x: toIn(box.x), y: toIn(box.y), w: toIn(box.w), h: toIn(box.h),
      fill: fill.color,
      gradient: fill.gradient ? { color: fill.color } : undefined,
      radiusPt: toPt(borderRadiusPx(pcs)),
    })
  })

  // 3) 金句页对角大引号：逐个子块检查 ::before/::after。
  //    不能默认“第一个子块＝上引号、最后一个子块＝下引号”：金句页允许在引用块之后
  //    再跟正常文案，此时两个引号都挂在 leading blockquote 上，不能再把下引号
  //    放到最后一页的普通段落上。
  Array.from(slide.children).forEach((host) => {
    (['::before', '::after'] as const).forEach((which) => {
      const pcs = styleOf(host, win, which)
      if (!hasVisiblePseudo(pcs)) return
      const raw = (pcs.content || '').trim().replace(/^["']|["']$/g, '')
      if (!raw || raw.length > 4) return
      const fsz = fontSizePx(pcs)
      if (fsz < 24) return
      const w = Math.max(1, measure(raw, pcs) + fsz * 0.08)
      const h = Math.max(fsz, lineHeightPx(pcs))
      const r = rectOf(host)
      let box: Box
      if (which === '::before') {
        // 上引号：right:100%; bottom:calc(100% - .35em) ⇒ 右下角贴着文案左上角
        const right = pcs.right !== 'auto' ? num(pcs.right) : r.w
        const bottom = pcs.bottom !== 'auto' ? num(pcs.bottom) : r.h - fsz * 0.35
        box = { x: r.x + r.w - right - w, y: r.y + r.h - bottom - h, w, h }
      } else {
        const left = pcs.left !== 'auto' ? num(pcs.left) : r.w
        const top = pcs.top !== 'auto' ? num(pcs.top) : r.h
        box = { x: r.x + left, y: r.y + top, w, h }
      }
      const font = resolveFontStack(pcs.fontFamily)
      const style: SceneRunStyle = {
        color: pseudoColor(pcs, slideBg),
        sizePt: toPt(fsz),
        latin: font.latin,
        ea: font.ea,
      }
      pushText(box, [{ runs: [{ text: raw, style }] }], pcs, {
        align: which === '::before' ? 'right' : 'left', decorative: true,
      })
    })
  })

  // ---------- 正文 ----------

  // 分栏页：先铺卡片/分隔线（在 .ys-cols 分支里完成）
  Array.from(slide.children).forEach((child) => emitBlock(child))
  applyCodeFlowExpansion(nodes, codeFlowExpansions)

  // ---------- 页脚 ----------

  if (opts.includeFooter && (opts.chapter || opts.pos)) {
    const footer = doc.querySelector('.ys-footer')
    const chapterEl = footer?.querySelector('.ys-chapter') || null
    const posEl = footer?.querySelector('.ys-pos') || null
    if (footer && (chapterEl || posEl)) {
      const fcs = styleOf(footer, win)
      const runs: SceneRun[] = []
      if (chapterEl && (chapterEl.textContent || '').trim()) {
        runs.push({ text: (chapterEl.textContent || '').trim(), style: trimRun(runStyleOf(chapterEl, slideBg)) })
      }
      if (chapterEl && posEl) {
        runs.push({ text: '   ', style: trimRun(runStyleOf(footer, slideBg)) })
      }
      if (posEl && (posEl.textContent || '').trim()) {
        runs.push({ text: (posEl.textContent || '').trim(), style: trimRun(runStyleOf(posEl, slideBg)) })
      }
      if (runs.length) {
        // 页脚是 `position:fixed` + flex 的收缩宽度：DOM 里刚好一行，但 PowerPoint
        // 的字体度量（尤其 CJK 的 em 宽度）会把 `3 / 9` 折到下一行 —— 直接禁掉换行。
        const box: Box = { x: rectOf(footer).x, y: rectOf(footer).y, w: rectOf(footer).w, h: rectOf(footer).h + EPS }
        pushText(box, [{ runs }], fcs, { align: 'left', noWrap: true, footer: true })
      }
    }
  }

  // ---------- 快速体检（按最终场景节点包围盒，四边精确测量） ----------

  const overflow = measureSceneOverflow(
    nodes,
    SLIDE_W_IN,
    SLIDE_H_IN,
  )
  const overflowPx = Math.max(overflow.top, overflow.right, overflow.bottom, overflow.left)
  // 为 PowerPoint 字体度量差异多折的一行预留底部空间；进入预留区只算“接近页底”，
  // 仍由导出层决定告警 / 压缩，避免在场景层重复改字号。
  const reserveIn = sceneReserveIn(nodes)
  const tightPx = measureSceneTightPx(nodes, SLIDE_H_IN, reserveIn)
  if (overflowPx > EPS) {
    warnings.push({
      kind: 'overflow',
      slide: opts.index,
      messageKey: 'pptxExport.warnOverflowPx',
      messageArgs: { px: overflowPx },
    })
  }

  return {
    index: opts.index,
    kind: opts.kind,
    title: opts.title,
    background: slideBg,
    notes: opts.notes,
    nodes,
    overflowPx,
    reserveIn,
    tightPx,
    overflow,
    warnings,
    fontScale: 1,
  }
}

/** 供调试/日志：场景节点计数 */
export function sceneSummary(scene: SlideScene): string {
  const counts: Record<string, number> = {}
  scene.nodes.forEach((n) => { counts[n.t] = (counts[n.t] || 0) + 1 })
  return Object.entries(counts).map(([k, v]) => `${k}:${v}`).join(' ')
}

/** 装饰色辅助（导出层调试用） */
export function accentOf(frame: PptxFrame): string {
  return flattenColor(frame.window.getComputedStyle(frame.slide).color, '#ffffff')
}

export { colorAlpha }
