# 元素动画

PPTForge 支持为幻灯片上的任意元素（文字、形状、图片等）添加动画效果。

## 快速开始

```javascript
// 单个动画
slide.addText('Hello', {
  x: 1, y: 2, w: 8, h: 1,
  animation: { type: 'fadeIn', dur: 500, trigger: 'onClick' }
})

// 多个动画（入场 + 退场）
slide.addText('Bye', {
  x: 1, y: 3, w: 8, h: 1,
  animation: [
    { type: 'flyIn', dir: 'bottom', dur: 800, trigger: 'onClick' },
    { type: 'fadeOut', dur: 500, trigger: 'afterPrev', delay: 2000 }
  ]
})
```

## AnimationProps 接口

```typescript
interface AnimationProps {
  type: AnimationType       // 动画效果类型（必填）
  class?: 'entrance' | 'exit'  // 入场/退场，默认取效果自身默认值
  dir?: 'top' | 'bottom' | 'left' | 'right'  // 方向（飞入/擦除等）
  inOut?: 'in' | 'out'     // 内/外（方框、菱形等形状动画）
  dur?: number              // 持续时间(毫秒)，默认 500
  delay?: number            // 延迟(毫秒)，默认 0
  trigger?: 'onClick' | 'withPrev' | 'afterPrev'  // 触发方式，默认 onClick
  order?: number            // 播放顺序（1起始），不指定则按添加顺序
  byParagraph?: boolean     // 逐段动画（仅文本），默认 false
}
```

## 支持的动画类型

### P0 核心动画（8 种）

| 类型 | 说明 | 默认分类 | 支持方向 |
|------|------|----------|----------|
| `appear` | 出现 | 入场 | - |
| `fadeIn` | 淡入 | 入场 | - |
| `flyIn` | 飞入 | 入场 | top / bottom / left / right |
| `wipe` | 擦除 | 入场 | down / left / right / up |
| `zoom` | 缩放 | 入场 | - |
| `disappear` | 消失 | 退场 | - |
| `fadeOut` | 淡出 | 退场 | - |
| `flyOut` | 飞出 | 退场 | top / bottom / left / right |

### P1 滤镜扩展动画（13 种）

| 类型 | 说明 | 参数 |
|------|------|------|
| `blinds` | 百叶窗 | dir: horizontal / vertical |
| `box` | 方框 | inOut: in / out |
| `checkerboard` | 棋盘 | dir: across |
| `circle` | 圆形 | inOut: in / out |
| `diamond` | 菱形 | inOut: in / out |
| `dissolve` | 溶解 | - |
| `plus` | 十字 | inOut: in / out |
| `randomBars` | 随机线条 | dir: horizontal / vertical |
| `split` | 拆分 | dir + inOut |
| `strips` | 条纹 | dir: downLeft 等 |
| `wedge` | 楔形 | - |
| `wheel` | 轮子 | dir: '1'-'8' (辐条数) |
| `random` | 随机 | - |

## 触发方式

### onClick（点击触发）

每个 `onClick` 动画独占一个"点击组"。放映时需手动点击才播放。

```javascript
slide.addText('点击出现', {
  x: 1, y: 1, w: 8, h: 1,
  animation: { type: 'fadeIn', trigger: 'onClick' }
})
```

### withPrev（与前一个同时）

与前一个动画同时播放，不需要额外点击。

```javascript
// 标题飞入后，副标题同时淡入
slide.addText('标题', {
  animation: { type: 'flyIn', dir: 'top', trigger: 'onClick' }
})
slide.addText('副标题', {
  animation: { type: 'fadeIn', trigger: 'withPrev' }
})
```

### afterPrev（前一个结束后）

前一个动画播放完毕后自动播放，可配合 `delay` 设置延迟。

```javascript
slide.addText('步骤 1', {
  animation: { type: 'flyIn', dir: 'left', trigger: 'onClick' }
})
slide.addText('步骤 2', {
  animation: { type: 'flyIn', dir: 'left', trigger: 'afterPrev', delay: 300 }
})
slide.addText('步骤 3', {
  animation: { type: 'flyIn', dir: 'left', trigger: 'afterPrev', delay: 300 }
})
```

## 逐段动画 (byParagraph)

当文本包含多个段落时，可以让每段依次出现：

```javascript
slide.addText([
  { text: '第一点：架构设计', options: { fontSize: 18 } },
  { text: '第二点：技术选型', options: { fontSize: 18 } },
  { text: '第三点：实现方案', options: { fontSize: 18 } },
], {
  x: 1, y: 2, w: 8, h: 3,
  animation: { type: 'fadeIn', dur: 400, trigger: 'onClick', byParagraph: true }
})
```

效果：每次点击显示一个段落。

## 多动画组合

一个元素可以同时拥有入场和退场动画：

```javascript
slide.addText('临时公告', {
  x: 1, y: 2, w: 8, h: 1,
  animation: [
    { type: 'flyIn', dir: 'bottom', dur: 500, trigger: 'onClick' },
    { type: 'fadeOut', dur: 800, trigger: 'afterPrev', delay: 3000 }
  ]
})
```

## OOXML 原理

### Timing Tree 结构

PowerPoint 动画基于 `<p:timing>` XML，采用树状结构：

```xml
<p:timing>
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
        <p:childTnLst>
          <!-- mainSeq: 主动画序列 -->
          <p:seq concurrent="1" nextAc="seek">
            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
              <p:childTnLst>
                <!-- Click Group 1 -->
                <p:par>
                  <p:cTn id="3" fill="hold">
                    <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
                    <p:childTnLst>
                      <!-- 具体动画效果 -->
                    </p:childTnLst>
                  </p:cTn>
                </p:par>
                <!-- Click Group 2 ... -->
              </p:childTnLst>
            </p:cTn>
          </p:seq>
        </p:childTnLst>
      </p:cTn>
    </p:par>
  </p:tnLst>
</p:timing>
```

### Click Group 分组规则

- `onClick` → 创建新的 Click Group
- `withPrev` → 加入当前 Click Group（`nodeType="withEffect"`）
- `afterPrev` → 加入当前 Click Group（`nodeType="afterEffect"`）

### 动画方法分类

PPTForge 中每种动画使用不同的 XML 方法：

| 方法 | 用途 | XML 元素 |
|------|------|----------|
| `set` | 瞬间出现/消失 | `<p:set>` (visibility) |
| `animEffect` | 滤镜动画 | `<p:animEffect filter="...">` |
| `anim` | 关键帧动画 | `<p:anim>` (位置移动) |

每个动画都包含 `<p:set>` 可见性控制 + 具体效果 XML。

### _animShapeId

每个幻灯片对象在创建时分配 `_animShapeId`，用于 timing XML 中的 `<p:spTgt spid="N"/>` 引用。该 ID 在 `gen-xml.ts` 中通过 `intTableNum` 或索引计算。
