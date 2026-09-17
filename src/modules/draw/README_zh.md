# Draw 模块

[English](README.md)

共享绘图引擎:一份结构化显示列表,一个遍历器,多个后端。零依赖。

```typescript
import { toSvg, rasterizeToRgba, type DrawList } from "documonster/draw";
```

## 特性

- **一份列表,所有后端** —— 构建一次 `DrawList`,即可得到标记、像素或 PDF 页面
- **没有往返** —— 渲染器之间不再传递 SVG 字符串,任何后端都不必重新解析另一个后端的输出
- **绝对坐标** —— 遍历器负责合成变换,交给后端的已是变换后的数值,因此新增一个后端只需几十行,而不是再造一个平行渲染器
- **自带文本测量** —— `measureText` / `wrapText`,生产者在构建列表之前就能确定盒子尺寸
- **零依赖** —— 纯 TypeScript,无外部包
- **跨平台** —— Node.js 与浏览器 API 一致

---

## 快速开始

```typescript
import { rasterizeToRgba, toSvg, type DrawList } from "documonster/draw";

const list: DrawList = {
  width: 120,
  height: 60,
  children: [
    {
      kind: "rect",
      x: 10,
      y: 10,
      width: 100,
      height: 40,
      rx: 6,
      paint: {
        fill: { r: 0.2, g: 0.4, b: 0.8, a: 1 },
        stroke: { r: 0, g: 0, b: 0, a: 1 },
        strokeWidth: 1
      }
    },
    {
      kind: "text",
      x: 60,
      y: 30,
      lines: [{ text: "Hello", dy: 4 }],
      style: { size: 14, anchor: "middle", fill: { r: 1, g: 1, b: 1, a: 1 } }
    }
  ]
};

const svg = toSvg(list); // 标记
const image = rasterizeToRgba(list, { scale: 2 }); // { width, height, data }
```

同一份列表画到 PDF 页面上 —— 第三个后端,它住在拥有"页面"这一概念的模块里:

```typescript
import { renderDrawList } from "documonster/draw";
import { Pdf, createPdfDrawSurface } from "documonster/pdf";

const doc = new Pdf.Builder();
const page = doc.addPage({ width: 595, height: 842 });
renderDrawList(list, createPdfDrawSurface(page, { x: 72, y: 700, width: 120, height: 60 }));
const bytes = await doc.build();
```

---

## 显示列表

`DrawList` 就是一个尺寸加一棵节点树,其中没有任何后端专属的东西。

```typescript
interface DrawList {
  readonly width: number;
  readonly height: number;
  readonly children: readonly DrawNode[];
  /** `<defs>` 原始标记,由 `group.svgFilterId` 引用。名字里就写明只服务 SVG。 */
  readonly svgDefs?: readonly string[];
}
```

### 图元

| `kind`     | 形状                                         |
| ---------- | -------------------------------------------- |
| `group`    | 一个变换、一个可选的轴对齐 `clip`,以及子节点 |
| `rect`     | 矩形,可选 `rx` 圆角半径                      |
| `ellipse`  | 由中心与半径确定的椭圆                       |
| `line`     | 单条线段                                     |
| `polyline` | 开放或闭合的点串                             |
| `path`     | move / line / cubic 命令                     |
| `sector`   | 扇形或环形楔块,由半径与角度确定              |
| `text`     | 一行或多行,按基线定位                        |

扇形是一等图元而非降级为 path,因为三个后端处理它的方式确实各不相同 —— SVG 原生支持弧,PDF 需要三次曲线,而光栅器可以逐像素判断半径与角度以得到精确边缘 —— 在上游降级会让其中两个变差。

### 涂色

```typescript
interface DrawPaint {
  readonly fill?: Rgba01;
  readonly stroke?: Rgba01;
  readonly strokeWidth?: number; // 用户单位,有描边时默认 1
  readonly dash?: readonly number[]; // 交替的实/虚,SVG 语义
  readonly fillRule?: "nonzero" | "evenodd";
}
```

颜色是 `Rgba01` —— 各通道取值 `0..1`,含 alpha。`cssColour("#3366cc")` 把 CSS 记号解析成它;`translucent("#3366cc", 0.5)` 一步完成解析并设置 alpha。

### 文本

文本按**基线**定位,因为这是所有后端都能如实表达的方式。垂直居中是生产者自己做的算术,不是它能请求的可移植属性。

```typescript
import { measureText, wrapText, POINTS_PER_PIXEL } from "documonster/draw";

const style = { size: 14, family: "Arial" };
const width = measureText("Revenue by region", style);
const lines = wrapText("A longer label that has to fit", style, 120);
```

注意单位:`measureText` 针对磅值返回用户单位,内部已应用 `POINTS_PER_PIXEL`。跳过这次换算会让每个标签宽出 4/3 —— 图例过宽、居中标题偏左就是这么来的。

### 裁剪

`group.clip` 是轴对齐矩形,在该 group 自己的坐标空间中书写,因此会随 `transform` 一起移动和缩放。嵌套裁剪取交集。三个后端都能精确表达:SVG 的 `clipPath`、PDF 的 `q … W n … Q`、光栅器里的裁剪测试。

### 变换

`DrawMatrix` 辅助函数:`IDENTITY`、`translate`、`scale`、`uniformScale`、`rotate`、`multiply`、`apply`、`rotationOf`。路径辅助函数:`arcToCubics`、`flattenPath`、`roundedRectToPath`、`sectorToPath`、`rectNode`。

---

## 后端

| 表面     | 位置                                      | 输出                                  |
| -------- | ----------------------------------------- | ------------------------------------- |
| SVG      | `documonster/draw`(`toSvg`)               | 标记                                  |
| 光栅     | `documonster/draw`(`rasterizeToRgba`)     | 经 `BasicRasterCanvas` 得到 RGBA 像素 |
| PDF 矢量 | `documonster/pdf`(`createPdfDrawSurface`) | 页面算子;Y 轴翻转由它独家负责         |

当一个表面只需要显示列表时,它与引擎住在一起;当它要画到引擎无从知晓的东西上时,它就住在目标旁边 —— 这就是 PDF 那个表面在 PDF 模块里的原因。

再写一个后端就是实现 `DrawSurface`。遍历器负责变换合成、描边与虚线缩放、文本旋转,所以表面只管落笔。`pushClip` / `popClip` 是可选的;省略它们会导致不裁剪地绘制,而不是丢弃内容。

### 是像素,不是 PNG

`rasterizeToRgba` 返回 `RgbaImage` —— `{ width, height, data }`,`width * height * 4` 字节的直通 alpha RGBA —— 而不是编码后的文件。编码 PNG 需要 DEFLATE,它位于本模块之上一层;把压缩库拉下来会让绘图引擎的每个使用者都为此付费。

所以接缝定在图像上,而配对只有两行。`documonster/archive` 拥有 DEFLATE 和 CRC-32,而 PNG 就是这两样东西:

```typescript
import { rasterizeToRgba } from "documonster/draw";
import { encodePng } from "documonster/archive";

const image = rasterizeToRgba(list, { scale: 2 });
const png = encodePng(image.data, image.width, image.height, { dpi: 192 });
```

在浏览器里你可能更愿意用平台自带的编码器 —— 用 `putImageData` 把像素放进 canvas,再调 `toBlob`。无论哪种方式,绘图引擎都不参与。

### 字体,以及中日韩文字是怎么画出来的

SVG 只写下字体名,由查看器去解析。光栅化必须拿到真实的字形轮廓,所以 `rasterizeToRgba` 会去找一张能覆盖该文本的已安装字体 —— 和 PDF 内嵌用的是同一套发现逻辑 —— 并且**逐字符沿一条字体链解析**,因为没有哪一张字体能同时覆盖 `Mixed 混合 ABC`。

由此有两件事,而且都是可上报的,不是静默的:

```typescript
import { rasterizeToRgba } from "documonster/draw";

const image = rasterizeToRgba(list);
if (image.uncoveredCodePoints.length > 0) {
  // 本机没有任何字体能画这些码点。它们是图上的空洞,而不是一个错误。
  console.warn(image.uncoveredCodePoints.map(cp => cp.toString(16)));
}
```

当你无法依赖宿主环境时就直接给字节 —— **在浏览器里这是拿到 ASCII 以外字形的唯一途径**,因为没有字体目录可搜,而内置的笔画字体只覆盖 ASCII 32–126,其余一律画成 `?`:

```typescript
const image = rasterizeToRgba(list, { fonts: [notoSansScBytes] });
```

`fonts` 同样接受已解析的字体和字体集合中的某一 face。反复渲染时请先解析一次:传字节的话每次调用都会重新解析,对中日韩字体意味着每次重建一张 43000 条目的 `cmap`,而且新对象还会错过字形缓存。

```typescript
import { parseRasterFont } from "documonster/draw";

const regular = parseRasterFont(songtiTtcBytes, 3); // .ttc 中的某一 face
const image = rasterizeToRgba(list, { fonts: [regular] });

// 不想自己解析时,等价写法:
rasterizeToRgba(list, { fonts: [{ data: songtiTtcBytes, collectionIndex: 3 }] });
```

再加上 `useSystemFonts: false`,输出就**只**取决于你传入的字体。快照测试和可复现构建需要这一点 —— 字体发现无法保证两台机器给出相同的像素:

```typescript
const image = rasterizeToRgba(list, { fonts: [notoSansScBytes], useSystemFonts: false });
```

字体属于一次渲染,或属于一张画布(`BasicRasterCanvas.setFonts`)—— 刻意没有进程级注册表,这样一个调用方就无法改变另一个调用方画布所用的字体。

格式控制字符永远不会被上报:变体选择符、连接符、bidi 隔离符在任何字体里都没有字形,所以它们不会被查找、不占宽度,也不会出现在 `uncoveredCodePoints` 里。制表符是那个恰好相反的例外 —— 它不落墨,但确实占位,所以笔位仍会推进。

### 需要 shaping 的文字

覆盖率不是文本出错的唯一途径。阿拉伯字母会按位置改变形状并与邻字连接;印度语的元音符号存储在辅音之后、却要画在它之前;泰文的符号需要叠加;从右往左的文本必须先重排再绘制。

其中**不依赖 OpenType 表**的那部分,光栅化器通过 `@utils/text-shaping` 已经做了:阿拉伯字母取其词首、词中、词尾或独立形态,从右往左的行被放成视觉顺序,印度语元音符号被移到它实际绘制的一侧。所以阿拉伯文现在是连写的,希伯来文的顺序也是对的而不是反的。

仍然缺的那部分需要字体自己的 GSUB 与 GPOS 表:合体连字仍按独立字形绘制,叠加符号没有定位。这一点选择上报,而不是假装:

```typescript
const image = rasterizeToRgba(list);
for (const warning of image.textWarnings) {
  console.warn(warning); // "Text contains Arabic. This rasteriser applies contextual forms…"
}
```

与 `uncoveredCodePoints` 不同,提供字体在这里没有帮助 —— 每个字形其实都画出来了。正因为它**看不出**坏掉,才必须上报。

**`toSvg` 和 DOCX 写入器对这些文字仍然是完全正确的**,因为两者都原样携带文本,把 shaping 交给查看器或 Word。**PDF 写入器现在也做 shaping。** 两者都只在某张字体确实能画出该形态时才替换:很多字体提供基础字母而完全不提供 presentation form,不加判断地替换在 PDF 里代价是 `.notdef` 豆腐块,而在光栅化器里 —— 因为它没有 notdef 可退 —— 代价是**整行空白**。所以两个后端都可能顺序正确却仍未连写,并各自上报自己实际做到了哪一半。

每段文本要做两个决定,而它们由不同的谓词把关。shaping 在文本**可能**需要时运行;而字形 advance 是否归一化到测量宽度,取决于内置 advance 表是否描述得了这段文本 —— 它只覆盖拉丁、CJK、西里尔和希腊字母,对泰米尔文报出的宽度约为真实值的一半,所以复杂文字改用字体自身的 advance 来排布。用同一个谓词决定这两件事,会在一个方向上把泰米尔文挤成重叠,在另一个方向上让夹有双向控制符的拉丁文整体位移。

`family`、`bold`、`italic` 决定的是**用哪张字体**,而不只是宽度。要求 `Courier New` 的标签在该字体已安装时就用它,要求加粗则使用该家族的粗体脸;回退目标是常规拉丁家族,而不是随便一张能覆盖 ASCII 的脸。文本测量仍来自内置的 advance 表,因此无论宿主装了什么字体,排版在每台机器上都一致。

---

## 刻意不做的事

- **元素级 opacity。** alpha 改为按 paint 表达。两者只在半透明描边压住自己填充的那条带上不同:元素 opacity 先合成完整图形再整体淡化,而按 paint 的 alpha 是先淡化填充、再把描边混合上去。要弥合这个差距,每个后端都得做离屏合成 —— PDF 要透明组和 Form XObject,光栅器要第二个缓冲区 —— 为一条不到一像素的接缝上这么多机械,不值得。
- **渐变与图案。** 没有任何生产者产出它们;图表引擎在任何渲染器看到之前就已把渐变填充降级为一个代表色。加上它们只会让 IR 和每个表面都变大,却不改变任何一个输出。
- **滤镜**,除了作为 SVG 专属的逃生口。DrawingML 的阴影在内容流或扫描线光栅器里没有对应物,所以 `DrawList.svgDefs` 和 `group.svgFilterId` 在名字里就写明 SVG,而不假装可移植。

---

## 生产者

任何产出 `DrawList` 的东西都免费获得三个后端。本仓库里的生产者是 Excel 图表引擎(`documonster/chart`)和 Mermaid 图表引擎([`documonster/mermaid`](../mermaid/README_zh.md))—— 后者完全没有实现任何后端,这正是检验上述主张是否成立的试纸。

跨后端一致性由 `src/modules/pdf/__tests__/draw-backend-parity.test.ts` 强制:同一份显示列表分别渲染为标记、像素和 PDF 算子,几何必须在各后端坐标约定的差异之内保持一致。
