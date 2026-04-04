# API 参考

## 创建演示文稿

```javascript
import pptxgen from 'pptforge'

const pptx = new pptxgen()
```

### 演示文稿属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `layout` | string | 幻灯片布局：`'LAYOUT_16x9'`, `'LAYOUT_16x10'`, `'LAYOUT_4x3'`, `'LAYOUT_WIDE'` |
| `author` | string | 作者名 |
| `company` | string | 公司名 |
| `subject` | string | 主题 |
| `title` | string | 标题 |

```javascript
pptx.layout = 'LAYOUT_16x9'   // 10" x 5.625"
pptx.author = 'PPTForge'
pptx.title = '演示文稿标题'
```

### 自定义布局尺寸

```javascript
pptx.defineLayout({ name: 'A4', width: 11.7, height: 8.3 })
pptx.layout = 'A4'
```

---

## 添加幻灯片

```javascript
const slide = pptx.addSlide()

// 使用指定布局
const slide = pptx.addSlide({ masterName: 'MASTER_SLIDE' })
```

### 幻灯片属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `background` | BackgroundProps | 背景颜色或图片 |
| `color` | string | 默认字体颜色 |
| `hidden` | boolean | 是否隐藏 |
| `transition` | TransitionProps | 幻灯片过渡效果 |
| `slideNumber` | SlideNumberProps | 幻灯片页码 |

```javascript
slide.background = { color: '1A1A2E' }
slide.transition = { type: 'fade', dur: 700 }
```

---

## addText()

添加文本框到幻灯片。

```typescript
slide.addText(text: string | TextProps[], options?: TextPropsOptions): Slide
```

### 参数

**简单文本：**

```javascript
slide.addText('Hello World', {
  x: 1, y: 1, w: 8, h: 1,
  fontSize: 24, color: '333333',
})
```

**富文本（多段落/多样式）：**

```javascript
slide.addText([
  { text: '粗体标题', options: { fontSize: 24, bold: true } },
  { text: '普通正文', options: { fontSize: 16 } },
  { text: '红色强调', options: { fontSize: 16, color: 'FF0000' } },
], {
  x: 1, y: 1, w: 8, h: 3,
})
```

### TextPropsOptions

| 属性 | 类型 | 默认 | 说明 |
|------|------|------|------|
| `x` | Coord | - | 水平位置（英寸或百分比） |
| `y` | Coord | - | 垂直位置 |
| `w` | Coord | - | 宽度 |
| `h` | Coord | - | 高度 |
| `fontSize` | number | 18 | 字号（pt） |
| `fontFace` | string | - | 字体名 |
| `color` | string | - | 颜色（hex） |
| `bold` | boolean | false | 粗体 |
| `italic` | boolean | false | 斜体 |
| `underline` | boolean | false | 下划线 |
| `align` | string | 'left' | 水平对齐：left / center / right / justify |
| `valign` | string | 'top' | 垂直对齐：top / middle / bottom |
| `wrap` | boolean | true | 自动换行 |
| `rotate` | number | 0 | 旋转角度 |
| `fill` | object | - | 背景填充 `{ color: 'FF0000' }` |
| `shadow` | ShadowProps | - | 阴影效果 |
| `animation` | AnimationProps \| AnimationProps[] | - | **动画**（详见 [animation.md](./animation.md)） |

---

## addMedia()

添加音频或视频到幻灯片。

```typescript
slide.addMedia(options: MediaProps): Slide
```

### MediaProps

| 属性 | 类型 | 默认 | 说明 |
|------|------|------|------|
| `type` | string | - | **必填**：`'audio'`, `'video'`, `'online'` |
| `path` | string | - | URL 或本地路径 |
| `data` | string | - | base64 编码 |
| `link` | string | - | 在线嵌入链接（YouTube） |
| `cover` | string | - | 封面图（base64 或路径） |
| `extn` | string | - | 文件扩展名（路径无扩展名时使用） |
| `x` | Coord | - | 水平位置 |
| `y` | Coord | - | 垂直位置 |
| `w` | Coord | - | 宽度 |
| `h` | Coord | - | 高度 |
| `autoPlay` | boolean | false | **自动播放**（详见 [video-background.md](./video-background.md)） |
| `loop` | boolean | false | **循环播放**（详见 [video-background.md](./video-background.md)） |

### 示例

```javascript
// 视频背景
slide.addMedia({
  type: 'video',
  path: 'https://example.com/bg.mp4',
  x: 0, y: 0, w: '100%', h: '100%',
  autoPlay: true, loop: true,
})

// 普通视频
slide.addMedia({
  type: 'video',
  path: '/videos/demo.mp4',
  x: 1, y: 1, w: 8, h: 4.5,
})

// 音频
slide.addMedia({
  type: 'audio',
  path: '/audio/bgm.mp3',
  x: 1, y: 6, w: 2, h: 0.4,
})
```

---

## addFormula()

添加 LaTeX 公式到幻灯片。

```typescript
slide.addFormula(options: FormulaProps): Slide
```

### FormulaProps

| 属性 | 类型 | 默认 | 说明 |
|------|------|------|------|
| `latex` | string | - | **必填**：LaTeX 公式字符串 |
| `x` | Coord | - | 水平位置 |
| `y` | Coord | - | 垂直位置 |
| `w` | Coord | - | 宽度 |
| `h` | Coord | - | 高度 |
| `fontSize` | number | 14 | 字号 |
| `color` | string | - | 颜色（hex） |
| `align` | string | - | 水平对齐 |
| `valign` | string | - | 垂直对齐 |
| `animation` | AnimationProps \| AnimationProps[] | - | 动画 |

### 示例

```javascript
slide.addFormula({
  latex: 'E = mc^2',
  x: 1, y: 1, w: 4, h: 0.8,
  fontSize: 24,
})

slide.addFormula({
  latex: '\\sum_{i=1}^{n} x_i = x_1 + x_2 + \\cdots + x_n',
  x: 1, y: 2.5, w: 8, h: 1,
  fontSize: 18,
  animation: { type: 'fadeIn', trigger: 'onClick' }
})
```

---

## addImage()

添加图片到幻灯片。

```typescript
slide.addImage(options: ImageProps): Slide
```

### ImageProps

| 属性 | 类型 | 说明 |
|------|------|------|
| `path` | string | URL 或本地路径 |
| `data` | string | base64 编码 |
| `x`, `y`, `w`, `h` | Coord | 位置和尺寸 |
| `altText` | string | 无障碍文字 |
| `flipH` | boolean | 水平翻转 |
| `flipV` | boolean | 垂直翻转 |
| `hyperlink` | HyperlinkProps | 超链接 |
| `rotate` | number | 旋转角度 |
| `rounding` | boolean | 圆角 |
| `sizing` | object | 裁剪/适应 `{ type: 'cover' | 'contain' | 'crop' }` |
| `transparency` | number | 透明度 (0-100) |
| `animation` | AnimationProps | 动画 |

```javascript
slide.addImage({
  path: 'https://example.com/photo.jpg',
  x: 1, y: 1, w: 4, h: 3,
  animation: { type: 'zoom', dur: 500, trigger: 'onClick' }
})
```

---

## addShape()

添加形状到幻灯片。

```typescript
slide.addShape(shapeName: SHAPE_NAME, options?: ShapeProps): Slide
```

```javascript
import pptxgen from 'pptforge'

slide.addShape(pptxgen.shapes.RECTANGLE, {
  x: 1, y: 1, w: 4, h: 2,
  fill: { color: 'FF6600' },
  shadow: { type: 'outer', blur: 3, offset: 2, color: '000000' }
})
```

常用形状：`RECTANGLE`, `ROUNDED_RECTANGLE`, `OVAL`, `LINE`, `ARROW_RIGHT`, `STAR_5`, ...

---

## addChart()

添加图表到幻灯片。

```typescript
slide.addChart(type: CHART_NAME, data: IOptsChartData[], options?: IChartOpts): Slide
```

```javascript
slide.addChart(pptxgen.charts.BAR, [
  { name: 'Q1', labels: ['A','B','C'], values: [10, 20, 30] },
  { name: 'Q2', labels: ['A','B','C'], values: [15, 25, 35] },
], {
  x: 1, y: 1, w: 8, h: 4,
  showTitle: true, title: '季度报告',
})
```

图表类型：`BAR`, `BAR3D`, `LINE`, `PIE`, `DOUGHNUT`, `AREA`, `SCATTER`, `BUBBLE`, `RADAR`

---

## addTable()

添加表格到幻灯片。

```typescript
slide.addTable(rows: TableRow[], options?: TableProps): Slide
```

```javascript
slide.addTable([
  [{ text: '姓名', options: { bold: true } }, { text: '分数', options: { bold: true } }],
  ['张三', '95'],
  ['李四', '87'],
], {
  x: 1, y: 1, w: 8,
  border: { type: 'solid', color: '333333', pt: 1 },
  colW: [4, 4],
})
```

---

## addNotes()

添加演讲者备注。

```typescript
slide.addNotes(notes: string): Slide
```

---

## 幻灯片过渡

```typescript
interface TransitionProps {
  type?: 'fade' | 'push' | 'wipe' | 'split' | 'cut' | 'cover' | 'uncover' | 'none'
  dir?: 'left' | 'right' | 'top' | 'bottom'
  dur?: number          // 毫秒
  speed?: 'slow' | 'med' | 'fast'
  advOnClick?: boolean  // 点击切换（默认 true）
  advAfter?: number     // 自动切换延迟（毫秒）
}
```

```javascript
slide.transition = { type: 'wipe', dir: 'left', speed: 'slow' }
slide.transition = { type: 'fade', advAfter: 5000, advOnClick: false }
```

---

## 输出文件

```javascript
// 写入文件（Node.js）
await pptx.writeFile({ fileName: 'output.pptx' })

// 获取 base64
const base64 = await pptx.write({ outputType: 'base64' })

// 获取 ArrayBuffer
const buffer = await pptx.write({ outputType: 'arraybuffer' })

// 获取 Blob（浏览器）
const blob = await pptx.write({ outputType: 'blob' })

// 获取 Buffer（Node.js）
const nodeBuffer = await pptx.write({ outputType: 'nodebuffer' })
```
