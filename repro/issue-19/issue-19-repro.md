---
title: Issue #19 PPTX 复现样例
author: hnsylzg
date: 2026-11-16
---

## 问题 1：代码卡片底部留白

左侧代码块用于检查 PowerPoint 额外折行。

***

```text
const issue19_left_line_01 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
const issue19_left_line_02 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
const issue19_left_line_03 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
const issue19_left_line_04 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
```

***

右侧内容较短，用于对照左栏代码卡片增高后的等高与底部留白。

```text
const issue19_right_01 = "short";
const issue19_right_02 = "short";
```

***

<!-- notes: 预期：PowerPoint 额外折行后，代码框上下留白仍一致；同一排两张卡片同步增高并保持等高。 -->

---

## 问题 2：分栏区域与栏外文案间距

**⬆ 上边距检查点：这段与卡片顶部的间距。**

***

左栏文案，用于建立分栏卡片高度。

***

右栏文案，用于建立分栏卡片高度。

***

**⬇ 下边距检查点：“落款”自身的上边距。**

<!-- notes: 预期：卡片上方的检查文案与卡片顶部、卡片底部与“落款”之间，都应保留正常的正文外部间距；“落款”应有明确的上边距。 -->

---

## 问题 3：代码折行后的后续内容

- 第一项：代码块之前的内容。

```text
const issue19_flow_01 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
const issue19_flow_02 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
const issue19_flow_03 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
const issue19_flow_04 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
const issue19_flow_05 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
const issue19_flow_06 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
```

- 第二项：位于代码块之后，预期随代码块额外折行同步下移。
- 第三项：用于观察第二项与第三项之间的原始列表间距是否保持不变。

<!-- notes: 预期：长代码在 PowerPoint 中额外折行后，第二项和第三项同步下移；后续节点不重叠、不侵入代码框。 -->
