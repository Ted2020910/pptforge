# LaTeX 公式

PPTForge 支持在幻灯片中插入 LaTeX 数学公式，自动将 LaTeX 转换为 PowerPoint 原生的 OMML (Office Math Markup Language) 格式。

## 快速开始

```javascript
const slide = pptx.addSlide()

// 简单公式
slide.addFormula({
  latex: 'E = mc^2',
  x: 1, y: 1, w: 4, h: 0.8,
  fontSize: 24,
})

// 复杂公式
slide.addFormula({
  latex: '\\int_0^\\infty e^{-x^2} dx = \\frac{\\sqrt{\\pi}}{2}',
  x: 1, y: 2.5, w: 6, h: 1,
  fontSize: 18,
})

// 矩阵
slide.addFormula({
  latex: '\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}',
  x: 1, y: 4, w: 3, h: 1.5,
})
```

## FormulaProps 接口

```typescript
interface FormulaProps extends PositionProps {
  latex: string            // LaTeX 公式字符串（必填）
  _omml?: string           // 预转换的 OMML（可选，跳过自动转换）
  fontSize?: number        // 字号，默认 18
  fontFace?: string        // 字体
  color?: string           // 颜色（hex）
  bold?: boolean           // 粗体
  italic?: boolean         // 斜体
  animation?: AnimationProps | AnimationProps[]  // 动画
}
```

## 转换流水线

PPTForge 使用三步流水线将 LaTeX 转换为 PowerPoint 可识别的格式：

```
LaTeX 字符串
    │
    ▼  [KaTeX 库]
MathML (XML)
    │
    ▼  [formula-converter.ts]
OMML (Office Math Markup Language)
    │
    ▼  [gen-xml.ts]
PowerPoint slide XML 中的 <a14:m> 元素
```

### Step 1: LaTeX → MathML

使用 [KaTeX](https://katex.org/) 库将 LaTeX 渲染为 MathML：

```typescript
// formula-converter.ts
import katex from 'katex'

function latexToMathML(latex: string): string {
  return katex.renderToString(latex, {
    output: 'mathml',
    throwOnError: false,
    displayMode: true,
    strict: false,
  })
}
```

### Step 2: MathML → OMML

自定义转换器将 MathML 元素逐个映射为 OMML：

| MathML 元素 | OMML 元素 | 说明 |
|-------------|-----------|------|
| `<mfrac>` | `<m:f>` | 分数 |
| `<msqrt>` | `<m:rad>` | 平方根 |
| `<mroot>` | `<m:rad>` | n次方根 |
| `<msub>` | `<m:sSub>` | 下标 |
| `<msup>` | `<m:sSup>` | 上标 |
| `<msubsup>` | `<m:sSubSup>` | 上下标 |
| `<mover>` | `<m:limUpp>` / `<m:acc>` | 上方标记 |
| `<munder>` | `<m:limLow>` | 下方标记 |
| `<mtable>` | `<m:m>` | 矩阵/表格 |
| `<mrow>` | (递归处理) | 行 |
| `<mi>` | `<m:r>` | 标识符（斜体） |
| `<mn>` | `<m:r>` | 数字 |
| `<mo>` | `<m:r>` | 运算符 |

特殊处理：
- **积分/求和**：`∫`, `∑`, `∏` 等大运算符转为 `<m:nary>` 元素
- **Unicode 数学斜体**：变量自动转为 Unicode 数学斜体字符（如 `x` → `𝑥`）
- **括号匹配**：`(`, `)`, `[`, `]`, `{`, `}` 转为 `<m:d>` 定界符

### Step 3: OMML → PowerPoint XML

在幻灯片 XML 中，OMML 嵌入在文本框的 `<a14:m>` 元素中：

```xml
<p:sp>
  <p:txBody>
    <a:p>
      <a14:m>
        <m:oMathPara>
          <m:oMath>
            <!-- OMML 公式内容 -->
            <m:r>
              <m:rPr><m:sty m:val="p"/></m:rPr>
              <m:t>E</m:t>
            </m:r>
            <m:r><m:t>=</m:t></m:r>
            <m:r><m:t>m</m:t></m:r>
            <m:sSup>
              <m:e><m:r><m:t>c</m:t></m:r></m:e>
              <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
            </m:sSup>
          </m:oMath>
        </m:oMathPara>
      </a14:m>
    </a:p>
  </p:txBody>
</p:sp>
```

## 自动公式检测

`formula-parser.ts` 支持在普通文本中自动检测 LaTeX 公式：

```typescript
import { parseTextWithFormulas, hasFormula } from './formula-parser'

// 检测文本是否包含公式
hasFormula('普通文本 $E=mc^2$ 更多文本')  // true
hasFormula('没有公式的文本')                  // false

// 解析文本，分离普通文本和公式
parseTextWithFormulas('前缀 $x^2$ 中缀 $$\\frac{a}{b}$$ 后缀')
// 返回:
// [
//   { type: 'text', content: '前缀 ' },
//   { type: 'formula', content: 'x^2', isBlock: false },
//   { type: 'text', content: ' 中缀 ' },
//   { type: 'formula', content: '\\frac{a}{b}', isBlock: true },
//   { type: 'text', content: ' 后缀' },
// ]
```

语法规则：
- `$...$` — 行内公式
- `$$...$$` — 块级公式

`gen-objects.ts` 中的 `addTextDefinition()` 在处理文本时，会自动调用公式检测，将混合文本+公式的内容正确渲染。

## 支持的 LaTeX 语法

以下是经过测试的常用语法：

| 类别 | LaTeX | 效果 |
|------|-------|------|
| 分数 | `\frac{a}{b}` | a/b |
| 上标 | `x^2` | x² |
| 下标 | `x_i` | xᵢ |
| 根号 | `\sqrt{x}` | √x |
| 积分 | `\int_a^b f(x)dx` | ∫ₐᵇf(x)dx |
| 求和 | `\sum_{i=1}^n x_i` | Σᵢ₌₁ⁿ xᵢ |
| 矩阵 | `\begin{pmatrix}a&b\\c&d\end{pmatrix}` | (a b; c d) |
| 希腊字母 | `\alpha, \beta, \gamma` | α, β, γ |
| 关系符 | `\leq, \geq, \neq` | ≤, ≥, ≠ |
| 箭头 | `\rightarrow, \Leftarrow` | →, ⇐ |
| 重音 | `\hat{x}, \bar{x}, \vec{x}` | x̂, x̄, x⃗ |

## 公式 + 动画

公式也支持动画效果：

```javascript
slide.addFormula({
  latex: '\\int_0^\\infty e^{-x^2} dx = \\frac{\\sqrt{\\pi}}{2}',
  x: 1, y: 2, w: 6, h: 1,
  fontSize: 20,
  animation: { type: 'fadeIn', dur: 800, trigger: 'onClick' }
})
```
