# 视频背景

PPTForge 支持将视频作为幻灯片背景，实现自动播放、无限循环，并支持点击暂停/恢复。

## 快速开始

```javascript
const slide = pptx.addSlide()

// 全屏视频背景（作为第一个元素，位于底层）
slide.addMedia({
  type: 'video',
  path: 'https://example.com/background.mp4',
  x: 0, y: 0, w: '100%', h: '100%',
  autoPlay: true,
  loop: true,
})

// 在视频上叠加文字
slide.addText('标题文字', {
  x: 0.5, y: 0.5, w: 9, h: 1,
  fontSize: 36, bold: true, color: 'FFFFFF',
})
```

## MediaProps 扩展属性

在原有 `addMedia()` 基础上，新增两个属性：

```typescript
interface MediaProps {
  type: 'audio' | 'video'   // 媒体类型
  path?: string              // URL 或本地路径
  data?: string              // base64 编码
  x: Coord                   // 水平位置
  y: Coord                   // 垂直位置
  w: Coord                   // 宽度
  h: Coord                   // 高度

  // ====== 新增 ======
  autoPlay?: boolean         // 自动播放，默认 false
  loop?: boolean             // 循环播放，默认 false
}
```

## 使用场景

### 场景 1：视频背景 + 文字叠加

```javascript
const slide = pptx.addSlide()

// 视频铺满全屏
slide.addMedia({
  type: 'video',
  path: 'https://cdn.example.com/ocean.mp4',
  x: 0, y: 0, w: '100%', h: '100%',
  autoPlay: true,
  loop: true,
})

// 白色大标题
slide.addText('Ocean Theme', {
  x: 0.5, y: 2, w: 9, h: 1.5,
  fontSize: 48, bold: true, color: 'FFFFFF',
  animation: { type: 'fadeIn', dur: 1000, trigger: 'afterPrev' }
})
```

### 场景 2：仅自动播放（不循环）

```javascript
slide.addMedia({
  type: 'video',
  path: videoUrl,
  x: 1, y: 1, w: 8, h: 4.5,
  autoPlay: true,
  loop: false,   // 播放一次后停止
})
```

### 场景 3：视频 + 元素动画混合

```javascript
const slide = pptx.addSlide()

// 视频背景
slide.addMedia({
  type: 'video',
  path: videoUrl,
  x: 0, y: 0, w: '100%', h: '100%',
  autoPlay: true, loop: true,
})

// 带动画的文字
slide.addText('动画标题', {
  x: 0.5, y: 1, w: 9, h: 1,
  fontSize: 36, color: 'FFFFFF',
  animation: { type: 'fadeIn', dur: 800, trigger: 'onClick' }
})

slide.addText('飞入副标题', {
  x: 0.5, y: 2.5, w: 9, h: 0.6,
  fontSize: 18, color: 'FFFF00',
  animation: { type: 'flyIn', dir: 'bottom', dur: 500, trigger: 'afterPrev', delay: 300 }
})
```

## OOXML 实现原理

### 视频元素结构

视频在 slide XML 中以 `<p:pic>` 元素表示：

```xml
<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="3" name="Media 0">
      <a:hlinkClick r:id="" action="ppaction://media"/>
    </p:cNvPr>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr>
      <a:videoFile r:link="rId3"/>           <!-- 视频文件引用 -->
      <p:extLst>
        <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">
          <p14:media r:embed="rId4"/>         <!-- MS 媒体引用 -->
        </p:ext>
      </p:extLst>
    </p:nvPr>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId5"/>                  <!-- 封面图引用 -->
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="0" y="0"/>
      <a:ext cx="9144000" cy="5143500"/>      <!-- 全屏尺寸 -->
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>
```

### 三重 rId 分配

每个视频需要 **3 个关系 ID**：

| rId | 关系类型 | 用途 |
|-----|----------|------|
| rId3 | `relationships/video` | 视频文件（r:link） |
| rId4 | `relationships/media` | Microsoft 媒体扩展（r:embed） |
| rId5 | `relationships/image` | 封面截图（a:blip r:embed） |

在 `.rels` 文件中：

```xml
<Relationships>
  <Relationship Id="rId1" Type=".../slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type=".../notesSlide" Target="../notesSlides/notesSlide1.xml"/>
  <Relationship Id="rId3" Type=".../video" Target="../media/media-1-1.mp4"/>
  <Relationship Id="rId4" Type=".../media" Target="../media/media-1-1.mp4"/>
  <Relationship Id="rId5" Type=".../image" Target="../media/image-1-3.png"/>
</Relationships>
```

> **注意**：`rId1` 和 `rId2` 固定分配给 slideLayout 和 notesSlide，这就是 `getNewRelId()` 需要 `+2` 偏移的原因。

### 自动播放 Timing XML

自动播放通过 `<p:timing>` 中的 `playFrom(0.0)` 命令实现：

```xml
<p:timing>
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
        <p:childTnLst>
          <!-- mainSeq -->
          <p:seq concurrent="1" nextAc="seek">
            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
              <p:childTnLst>
                <!-- playFrom 命令 -->
                <p:par>
                  <p:cTn id="3" fill="hold">
                    <p:stCondLst>
                      <p:cond delay="indefinite"/>
                      <!-- 关键：onBegin 事件引用 mainSeq 的 id -->
                      <p:cond evt="onBegin" delay="0">
                        <p:tn val="2"/>
                      </p:cond>
                    </p:stCondLst>
                    <p:childTnLst>
                      <p:par>
                        <p:cTn id="4" fill="hold">
                          <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                          <p:childTnLst>
                            <p:par>
                              <p:cTn id="5" presetID="1" presetClass="mediacall"
                                     presetSubtype="0" fill="hold" nodeType="afterEffect">
                                <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                <p:childTnLst>
                                  <p:cmd type="call" cmd="playFrom(0.0)">
                                    <p:cBhvr additive="base">
                                      <p:cTn id="6" dur="1" fill="hold"/>
                                      <p:tgtEl><p:spTgt spid="3"/></p:tgtEl>
                                    </p:cBhvr>
                                  </p:cmd>
                                </p:childTnLst>
                              </p:cTn>
                            </p:par>
                          </p:childTnLst>
                        </p:cTn>
                      </p:par>
                    </p:childTnLst>
                  </p:cTn>
                </p:par>
              </p:childTnLst>
            </p:cTn>
          </p:seq>

          <!-- 循环控制 -->
          <p:video fullScrn="0">
            <p:cMediaNode>
              <p:cTn id="7" repeatCount="indefinite" fill="hold" display="1">
                <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
              </p:cTn>
              <p:tgtEl><p:spTgt spid="3"/></p:tgtEl>
            </p:cMediaNode>
          </p:video>

          <!-- 点击暂停/恢复 -->
          <p:seq concurrent="1" nextAc="seek">
            <p:cTn id="8" restart="whenNotActive" fill="hold"
                   evtFilter="cancelBubble" nodeType="interactiveSeq">
              <!-- onClick 触发 togglePause -->
            </p:cTn>
          </p:seq>

        </p:childTnLst>
      </p:cTn>
    </p:par>
  </p:tnLst>
</p:timing>
```

### 关键机制解析

**自动触发（onBegin）**：
- `<p:cond delay="indefinite"/>` 表示需要点击触发
- `<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>` 表示当 mainSeq（id=2）开始时自动触发
- 双重条件的效果：幻灯片开始放映时，mainSeq 自动开始，onBegin 事件触发 playFrom

**循环（repeatCount）**：
- `<p:video>` 节点作为 tmRoot 的直接子节点（与 mainSeq 并列）
- `repeatCount="indefinite"` 表示无限重复

**点击暂停（togglePause）**：
- `interactiveSeq` 监听视频元素的 onClick 事件
- 执行 `togglePause` 命令切换播放/暂停状态

### 动画 + 视频混合

当幻灯片同时包含元素动画和视频时，PPTForge 将两者合并到同一个 `<p:timing>` 根节点中：

```
<p:timing>
  └── tmRoot
      ├── mainSeq
      │   ├── [动画 Click Groups]     ← 元素动画
      │   └── [playFrom 命令]         ← 视频自动播放
      ├── <p:video>                   ← 视频循环控制
      └── interactiveSeq             ← 视频点击暂停
```

每个 OOXML 幻灯片只允许一个 `<p:timing>` 节点，因此 `makeXmlTiming()` 函数必须将所有动画和媒体 timing 合并输出。

## URL 远程视频注意事项

使用远程 URL 作为视频源时：
- PPTForge 会自动处理 URL 中的查询参数（如 `?download`）
- 文件扩展名从 URL 路径部分提取，不包含查询参数
- 示例：`https://cdn.example.com/video.mp4?download` → 扩展名 `mp4`
