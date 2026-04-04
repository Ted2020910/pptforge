# 架构设计

## OOXML / PPTX 文件格式

PowerPoint `.pptx` 文件本质上是一个 **ZIP 压缩包**，内部包含一组 XML 文件。解压后的典型结构：

```
my-presentation.pptx (ZIP)
├── [Content_Types].xml          # 内容类型注册表
├── _rels/
│   └── .rels                    # 顶层关系文件
├── docProps/
│   ├── app.xml                  # 应用属性
│   └── core.xml                 # 核心属性（标题、作者）
└── ppt/
    ├── presentation.xml         # 演示文稿主文件
    ├── presProps.xml             # 演示属性
    ├── tableStyles.xml           # 表格样式
    ├── viewProps.xml             # 视图属性
    ├── _rels/
    │   └── presentation.xml.rels # 演示文稿关系
    ├── slideLayouts/             # 幻灯片布局模板
    │   ├── slideLayout1.xml
    │   └── _rels/
    ├── slideMasters/             # 幻灯片母版
    │   ├── slideMaster1.xml
    │   └── _rels/
    ├── slides/                   # 幻灯片内容（核心！）
    │   ├── slide1.xml
    │   ├── slide2.xml
    │   └── _rels/
    │       ├── slide1.xml.rels   # 幻灯片1的关系（媒体引用等）
    │       └── slide2.xml.rels
    ├── media/                    # 媒体文件（图片、视频）
    │   ├── image1.png
    │   └── video1.mp4
    └── theme/
        └── theme1.xml            # 主题
```

### 关键概念

**关系 (Relationships)**：OOXML 中，文件之间通过 `rId`（关系 ID）相互引用。例如 `slide1.xml` 中引用一张图片：

```xml
<!-- slide1.xml 中 -->
<a:blip r:embed="rId3"/>

<!-- slide1.xml.rels 中 -->
<Relationship Id="rId3" Type=".../image" Target="../media/image1.png"/>
```

**EMU 单位**：PowerPoint 内部使用 EMU (English Metric Units)：
- 1 英寸 = 914,400 EMU
- 1 点 = 12,700 EMU

PPTForge 接受英寸或百分比作为输入，内部自动转换为 EMU。

## 模块职责

### 核心模块

| 文件 | 行数 | 职责 |
|------|------|------|
| `pptxgen.ts` | ~800 | 主入口。创建演示文稿、管理幻灯片、输出为 ZIP |
| `slide.ts` | ~280 | Slide 类。提供 `addText()`, `addMedia()`, `addFormula()` 等 API |
| `core-interfaces.ts` | ~2,300 | 所有 TypeScript 接口/类型定义 |
| `core-enums.ts` | ~800 | 枚举常量（形状名、占位符类型、默认值等） |

### 生成模块

| 文件 | 行数 | 职责 |
|------|------|------|
| `gen-objects.ts` | ~1,400 | 将用户 API 参数转为内部 `ISlideObject` 数据结构 |
| `gen-xml.ts` | ~1,900 | **核心**：将 `ISlideObject` 渲染为 slide XML 字符串 |
| `gen-animation.ts` | ~750 | **自定义**：生成 `<p:timing>` 动画 XML + 视频媒体 timing |
| `gen-charts.ts` | ~2,000 | 图表 XML 生成 |
| `gen-tables.ts` | ~750 | 表格 XML 生成（含自动分页） |
| `gen-media.ts` | ~300 | 媒体文件处理（读取、base64 编码） |
| `gen-utils.ts` | ~500 | 工具函数（EMU 转换、颜色处理、rId 分配） |

### 公式模块（自定义）

| 文件 | 行数 | 职责 |
|------|------|------|
| `formula-converter.ts` | ~270 | LaTeX → MathML（KaTeX） → OMML 转换 |
| `formula-parser.ts` | ~80 | 文本中 `$...$` / `$$...$$` 公式自动检测 |

## XML 生成流水线

```
用户调用 API
     │
     ▼
slide.addText({ animation: {...} })
slide.addMedia({ autoPlay: true })
slide.addFormula({ latex: '...' })
     │
     ▼
gen-objects.ts
  addTextDefinition()     → 创建 ISlideObject { _type: 'text', options, text }
  addMediaDefinition()    → 创建 ISlideObject { _type: 'media', mtype: 'video' }
  addFormulaDefinition()  → 创建 ISlideObject { _type: 'formula', _omml }
     │ 每个对象分配 _animShapeId（用于动画引用）
     ▼
gen-xml.ts :: makeXmlSlide()
  遍历 slide._slideObjects
     │
     ├── text/formula → 生成 <p:sp> (文本框 + OMML)
     ├── media        → 生成 <p:pic> (视频/音频)
     ├── image        → 生成 <p:pic> (图片)
     ├── chart        → 生成 <c:chartSpace>
     └── table        → 生成 <a:tbl>
     │
     ▼
gen-animation.ts :: makeXmlTiming()
  ├── collectAnimationEntries()  → 收集带 animation 的元素
  ├── collectMediaEntries()      → 收集 autoPlay/loop 的视频
  ├── buildClickGroups()         → 按触发方式分组
  └── 生成 <p:timing> XML
     │
     ▼
gen-xml.ts :: slideObjectRelationsToXml()
  生成 slide.xml.rels (rId 关系映射)
     │
     ▼
pptxgen.ts :: writeFile()
  JSZip 打包所有 XML + 媒体文件 → .pptx
```

## rId 分配机制

每个幻灯片的 `.rels` 文件中，rId 按以下顺序分配：

```
rId1 = slideLayout (默认)
rId2 = notesSlide  (默认)
rId3 = 第一个 _rels 引用（图片、超链接等）
rId4 = ...
...
rIdN = _relsMedia 引用（视频文件、MS media、封面图）
```

**视频特殊处理**：每个视频消耗 **3 个 rId**：
1. `videoFile r:link` — 视频文件关系
2. `p14:media r:embed` — Microsoft 媒体关系
3. `a:blip r:embed` — 封面图关系

`getNewRelId()` 函数计算新的 rId 时，必须加上 `+2` 偏移量来跳过 `rId1`(slideLayout) 和 `rId2`(notesSlide)：

```typescript
// gen-utils.ts
export function getNewRelId(target: PresSlide): number {
  return target._rels.length + target._relsChart.length
       + target._relsMedia.length + 2 + 1
  //                                ^^^ defaultRels 偏移
}
```

## 数据流图

```
          AnimationProps                MediaProps
          { type, dur,                 { autoPlay,
            trigger }                    loop }
               │                           │
               ▼                           ▼
        ISlideObject                 ISlideObject
        ._animShapeId=3             ._animShapeId=3
        .options.animation          .options.autoPlay
               │                           │
               └──────────┬────────────────┘
                          ▼
                   makeXmlTiming()
                          │
            ┌─────────────┼─────────────────┐
            ▼             ▼                  ▼
        mainSeq       <p:video>        interactiveSeq
     (click groups   (循环控制)        (点击暂停)
      + playFrom)
            │             │                  │
            └─────────────┼──────────────────┘
                          ▼
                    <p:timing> XML
                    (合并输出)
```
