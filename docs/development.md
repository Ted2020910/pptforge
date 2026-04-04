# 开发指南

## 环境要求

- Node.js >= 16
- npm >= 8
- TypeScript >= 5.0

## 安装依赖

```bash
cd backend/Tools/PptxGenJS
npm install
```

## 构建

```bash
# 完整构建（生成 src/bld/ 目录下的 JS 文件）
npm run build

# 等价于
npx rollup -c --bundleConfigAsCjs

# 监听模式（文件变更自动重新构建）
npm run watch
```

构建输出到 `src/bld/` 目录：

| 文件 | 格式 | 用途 |
|------|------|------|
| `pptxgen.es.js` | ES Module | Node.js / 现代浏览器 |
| `pptxgen.cjs.js` | CommonJS | 旧版 Node.js |
| `pptxgen.js` | IIFE | `<script>` 标签直接引用 |

## 运行测试

### 动画测试

```bash
node examples/test-animation.js
```

生成 `examples/test-animation.pptx`，包含多种动画效果的测试幻灯片。

### 视频背景测试

```bash
node examples/test-video-bg.js
```

生成 `examples/test-video-bg.pptx`，包含 3 张测试幻灯片：
1. 全屏视频背景 + 自动播放 + 循环
2. 仅自动播放（不循环）
3. 视频背景 + 元素动画混合

> 注意：视频背景测试使用远程 URL，需要网络连接。

## 调试方法

### 1. 解压 PPTX 分析 XML

PPTX 是 ZIP 文件，可以直接解压查看内部 XML：

**手动解压：**

```bash
# 复制为 .zip 后解压
cp output.pptx output.zip
unzip output.zip -d output_extracted/

# 查看幻灯片 XML
cat output_extracted/ppt/slides/slide1.xml

# 查看关系文件
cat output_extracted/ppt/slides/_rels/slide1.xml.rels
```

**使用 inspect_pptx_xml.py 工具：**

```bash
python tools/inspect_pptx_xml.py output.pptx
```

输出每张幻灯片的元素统计：
- `<p:sp>` 形状数量
- `<p:pic>` 图片/视频数量
- `<a:tbl>` 表格数量
- 图表引用数量

### 2. XML 格式化

解压后的 XML 通常是压缩格式（单行），可以用 `xmllint` 格式化：

```bash
xmllint --format output_extracted/ppt/slides/slide1.xml
```

或使用在线工具如 [XML Formatter](https://www.freeformatter.com/xml-formatter.html)。

### 3. 对比参考 PPTX

`examples/` 目录下保留了两个参考文件：

| 文件 | 用途 |
|------|------|
| `动画示例.pptx` | 从 PowerPoint 导出的标准动画参考 |
| `视频背景.pptx` | 从 PowerPoint 导出的视频背景参考 |

分析步骤：
1. 将参考 PPTX 解压
2. 查看 `slide1.xml` 中的 `<p:timing>` 结构
3. 与 PPTForge 生成的 XML 对比
4. 确认结构一致

### 4. 常见调试场景

**动画不生效？**
- 检查 `<p:timing>` XML 是否存在
- 确认 `presetID`, `presetClass`, `presetSubtype` 值正确
- 确认 `<p:spTgt spid="N"/>` 的 spid 与元素 ID 匹配

**视频不显示？**
- 检查 `.rels` 文件中 rId 映射是否正确
- 确认 `<a:videoFile r:link="rIdN"/>` 指向视频（不是 slideLayout）
- 检查媒体文件名是否包含 URL 查询参数（如 `?download`）

**公式显示为空白？**
- 检查 KaTeX 是否正确安装：`npm ls katex`
- 在代码中添加 `console.log(omml)` 查看转换结果
- 确认 `<a14:m>` 元素在 slide XML 中存在

## 项目结构

```
PptxGenJS/
├── src/                      # TypeScript 源码
│   ├── pptxgen.ts            # 主入口
│   ├── slide.ts              # Slide 类
│   ├── core-interfaces.ts    # 类型定义
│   ├── core-enums.ts         # 枚举常量
│   ├── gen-objects.ts        # 对象生成
│   ├── gen-xml.ts            # XML 生成
│   ├── gen-animation.ts      # 动画生成（自定义）
│   ├── gen-charts.ts         # 图表
│   ├── gen-tables.ts         # 表格
│   ├── gen-media.ts          # 媒体处理
│   ├── gen-utils.ts          # 工具函数
│   ├── formula-converter.ts  # 公式转换（自定义）
│   └── formula-parser.ts     # 公式解析（自定义）
├── types/
│   └── index.d.ts            # 公共类型定义
├── libs/
│   ├── jszip.min.js          # JSZip 库
│   └── polyfill.min.js       # 浏览器兼容
├── tools/
│   └── inspect_pptx_xml.py   # PPTX XML 分析工具
├── examples/
│   ├── test-animation.js     # 动画测试脚本
│   ├── test-video-bg.js      # 视频背景测试脚本
│   ├── 动画示例.pptx          # 参考：标准动画
│   └── 视频背景.pptx          # 参考：视频背景
├── docs/                     # 文档
├── package.json              # NPM 配置
├── tsconfig.json             # TypeScript 配置
├── rollup.config.mjs         # Rollup 打包配置
├── eslint.config.mjs         # ESLint 配置
├── index.js                  # ESM 桥接
└── .gitignore
```

## Rollup 打包配置

`rollup.config.mjs` 配置了三种输出格式：

```javascript
export default {
  input: 'src/pptxgen.ts',
  output: [
    { file: './src/bld/pptxgen.js',     format: 'iife', name: 'PptxGenJS' },
    { file: './src/bld/pptxgen.cjs.js', format: 'cjs' },
    { file: './src/bld/pptxgen.es.js',  format: 'es' },
  ],
  external: [/^node:.*/, ...Object.keys(pkg.dependencies)],
  plugins: [resolve(), commonjs(), typescript()],
}
```

外部依赖不打包进输出文件：`jszip`, `katex`, `@xmldom/xmldom`, `image-size` 等。

## 修改代码后的工作流

1. 修改 `src/*.ts` 文件
2. `npm run build` 重新编译
3. 运行测试脚本验证
4. 解压生成的 PPTX 检查 XML
5. 用 PowerPoint 打开验证效果
