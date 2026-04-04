# PPTForge 文档

> PPTForge — 基于 PptxGenJS 的增强型 PowerPoint 生成引擎

## 简介

PPTForge 是一个 TypeScript/JavaScript 库，用于程序化生成 PowerPoint (.pptx) 文件。它在开源 [PptxGenJS](https://github.com/gitbrent/PptxGenJS) v4.0 基础上，新增了三大核心能力：

| 能力 | 说明 | 文档 |
|------|------|------|
| 元素动画 | 21 种动画效果 + 3 种触发方式 + 逐段动画 | [animation.md](./animation.md) |
| 视频背景 | 视频全屏自动播放 + 无限循环 + 点击暂停 | [video-background.md](./video-background.md) |
| LaTeX 公式 | LaTeX → MathML → OMML 自动转换 | [formula.md](./formula.md) |

## 架构概览

```
┌─────────────────────────────────────────────────┐
│                 用户 API 层                       │
│  pptxgen.ts  +  slide.ts                         │
│  addText() / addMedia() / addFormula() / ...     │
├─────────────────────────────────────────────────┤
│               对象定义层                          │
│  gen-objects.ts                                   │
│  将用户参数转为 ISlideObject 内部数据结构          │
├──────────┬──────────┬──────────┬────────────────┤
│ gen-xml  │gen-anim  │gen-media │ gen-charts     │
│ 幻灯片XML│动画Timing│媒体处理  │ 图表XML        │
│ gen-tables│         │          │                │
├──────────┴──────────┴──────────┴────────────────┤
│              公式转换层（自定义）                   │
│  formula-parser.ts  +  formula-converter.ts      │
│  LaTeX 解析 → KaTeX(MathML) → OMML              │
├─────────────────────────────────────────────────┤
│              工具层                               │
│  gen-utils.ts  +  core-enums.ts                  │
│  core-interfaces.ts                              │
├─────────────────────────────────────────────────┤
│              输出层                               │
│  JSZip → .pptx (ZIP 包含 XML 文件)               │
└─────────────────────────────────────────────────┘
```

## 快速开始

```bash
npm install
npm run build
```

```javascript
import pptxgen from './dist/pptxgen.es.js'

const pptx = new pptxgen()
pptx.layout = 'LAYOUT_16x9'

const slide = pptx.addSlide()

// 添加带动画的文字
slide.addText('Hello PPTForge!', {
  x: 1, y: 2, w: 8, h: 1.5,
  fontSize: 36, bold: true, color: 'FF6600',
  animation: { type: 'fadeIn', dur: 800, trigger: 'onClick' }
})

// 添加视频背景
slide.addMedia({
  type: 'video',
  path: 'https://example.com/bg.mp4',
  x: 0, y: 0, w: '100%', h: '100%',
  autoPlay: true, loop: true,
})

// 添加 LaTeX 公式
slide.addFormula({
  latex: 'E = mc^2',
  x: 1, y: 4, w: 4, h: 0.8,
  fontSize: 24,
})

await pptx.writeFile({ fileName: 'output.pptx' })
```

## 文档导航

| 文档 | 内容 |
|------|------|
| [architecture.md](./architecture.md) | 架构设计：OOXML 原理、模块职责、数据流 |
| [animation.md](./animation.md) | 元素动画：21 种效果、触发机制、OOXML 原理 |
| [video-background.md](./video-background.md) | 视频背景：自动播放、循环、rId 分配原理 |
| [formula.md](./formula.md) | LaTeX 公式：转换流水线、自动检测 |
| [api-reference.md](./api-reference.md) | API 参考：所有方法和参数 |
| [development.md](./development.md) | 开发指南：构建、调试、工具 |

## 源码结构

```
src/
├── pptxgen.ts            # 主入口类 PptxGenJS
├── slide.ts              # Slide 类（addText/addMedia/addFormula 等）
├── core-interfaces.ts    # 所有 TypeScript 类型定义
├── core-enums.ts         # 枚举常量（形状、颜色、布局等）
├── gen-objects.ts        # 对象定义（用户参数 → 内部数据结构）
├── gen-xml.ts            # 幻灯片 XML 生成（核心）
├── gen-animation.ts      # 动画 + 视频媒体 timing XML 生成（自定义）
├── gen-charts.ts         # 图表 XML 生成
├── gen-tables.ts         # 表格 XML 生成
├── gen-media.ts          # 媒体处理（文件读取、base64）
├── gen-utils.ts          # 工具函数（EMU 转换、颜色、rId 分配）
├── formula-converter.ts  # LaTeX → MathML → OMML 转换（自定义）
└── formula-parser.ts     # LaTeX 公式自动检测解析（自定义）
```

## 致谢

PPTForge 基于 [PptxGenJS](https://github.com/gitbrent/PptxGenJS) (MIT License, Brent Ely) 开发。
