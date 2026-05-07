# Excel 模块

[English](README.md)

现代 TypeScript Excel 工作簿管理器 — 读取、操作和写入 XLSX 与 JSON 格式的电子表格，零运行时依赖。

## 功能特性

- **创建、读取和修改 XLSX 文件** — 完整 Open XML 支持
- **多工作表支持** — 添加、删除、重排序、复制
- **单元格样式** — 字体、颜色、边框、填充、对齐、数字格式
- **单元格合并和格式化** — 合并区域、富文本、超链接
- **行列属性** — 宽度、高度、隐藏、大纲级别、自动适应
- **冻结窗格和拆分视图** — 冻结行/列、在指定位置拆分
- **富文本支持** — 单个单元格内多种字体/样式
- **公式和计算值** — 共享公式、定义名称
- **数据验证** — 列表、整数、小数、日期、文本长度、自定义
- **条件格式** — 单元格值、色阶、数据条、图标集
- **图片** — JPEG、PNG、GIF，支持单单元格和双单元格锚点
- **超链接** — 内部链接、外部链接、邮件链接
- **数据透视表** — 读取和保留数据透视表定义
- **图表** — 创建/读取/编辑 classic chart、ChartEx 现代图表、组合图、数据透视图、图表工作表，并提供 SVG/PNG/PDF 预览
- **表格** — 自动筛选、汇总行、结构化引用
- **批注和备注** — 线程批注、旧版备注
- **复选框** — 表单控件和单元格级复选框
- **页面设置** — 打印区域、打印标题、页眉/页脚、分页符
- **数据保护** — 带密码的工作表保护（SHA-512）
- **流式处理** — `WorkbookReader` 和 `WorkbookWriter` 处理大文件
- **CSV 导入/导出** — `readCsv`、`writeCsv`、`readCsvFile`、`writeCsvFile`
- **Markdown 导入/导出** — `readMarkdown`、`writeMarkdown`、`readMarkdownFile`、`writeMarkdownFile`
- **PDF 导出** — `excelToPdf()`，完整支持样式、分页、字体、加密
- **浏览器支持** — `xlsx.load()`、`xlsx.writeBuffer()`，无需 polyfill

## 快速开始

### 创建工作簿

```typescript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("My Sheet");

// 添加数据
sheet.addRow(["姓名", "年龄", "邮箱"]);
sheet.addRow(["张三", 30, "zhang@example.com"]);
sheet.addRow(["李四", 25, "li@example.com"]);

// Node.js：写入文件
await workbook.xlsx.writeFile("output.xlsx");

// 浏览器：写入缓冲区
const buffer = await workbook.xlsx.writeBuffer();
```

### 读取工作簿

```typescript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();

// Node.js：从文件读取
await workbook.xlsx.readFile("input.xlsx");

// 浏览器：从 ArrayBuffer 读取
await workbook.xlsx.load(arrayBuffer);

const worksheet = workbook.getWorksheet(1);
worksheet.eachRow((row, rowNumber) => {
  console.log("行 " + rowNumber + " = " + JSON.stringify(row.values));
});
```

### 设置单元格样式

```typescript
const cell = worksheet.getCell("A1");
cell.value = "你好";
cell.font = {
  name: "Arial",
  size: 16,
  bold: true,
  color: { argb: "FFFF0000" }
};
cell.fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFFFFF00" }
};
cell.border = {
  top: { style: "thin" },
  left: { style: "thin" },
  bottom: { style: "thin" },
  right: { style: "thin" }
};
cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
cell.numFmt = "$#,##0.00";
```

### 数字格式

```typescript
// 货币
cell.numFmt = "$#,##0.00";

// 百分比
cell.numFmt = "0.00%";

// 日期
cell.numFmt = "yyyy-mm-dd";

// 自定义
cell.numFmt = '#,##0.00 "单位"';
```

### 富文本

```typescript
cell.value = {
  richText: [
    { text: "粗体 ", font: { bold: true } },
    { text: "和 ", font: {} },
    { text: "红色", font: { color: { argb: "FFFF0000" } } }
  ]
};
```

### 公式

```typescript
cell.value = { formula: "SUM(A1:A10)" };
cell.value = { formula: "A1+B1", result: 42 }; // 带缓存结果

// 共享公式
sheet.getCell("A1").value = { formula: "B1*2", shareType: "shared", ref: "A1:A10" };

// 定义名称
workbook.definedNames.add("MyRange", "Sheet1!$A$1:$B$10");
```

### 数据验证

```typescript
worksheet.getCell("A1").dataValidation = {
  type: "list",
  allowBlank: true,
  formulae: ['"选项1,选项2,选项3"']
};

worksheet.getCell("B1").dataValidation = {
  type: "whole",
  operator: "between",
  formulae: [1, 100],
  showErrorMessage: true,
  errorTitle: "无效",
  error: "请输入 1 到 100 之间的数字"
};
```

### 条件格式

```typescript
worksheet.addConditionalFormatting({
  ref: "A1:A100",
  rules: [
    {
      type: "cellIs",
      operator: "greaterThan",
      formulae: [90],
      style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FF00FF00" } } },
      priority: 1
    }
  ]
});
```

### 图片

```typescript
import { readFileSync } from "fs";

const imageId = workbook.addImage({
  buffer: readFileSync("logo.png"),
  extension: "png"
});

worksheet.addImage(imageId, {
  tl: { col: 0, row: 0 },
  br: { col: 3, row: 5 }
});
```

### 表格

```typescript
worksheet.addTable({
  name: "SalesTable",
  ref: "A1",
  headerRow: true,
  totalsRow: true,
  columns: [
    { name: "产品", totalsRowLabel: "合计", filterButton: true },
    { name: "收入", totalsRowFunction: "sum", filterButton: true }
  ],
  rows: [
    ["小工具", 1000],
    ["大工具", 2500]
  ]
});
```

### 合并单元格

```typescript
worksheet.mergeCells("A1:D1");
worksheet.getCell("A1").value = "合并标题";
worksheet.getCell("A1").alignment = { horizontal: "center" };
```

### 冻结窗格

```typescript
// 冻结首行
worksheet.views = [{ state: "frozen", ySplit: 1 }];

// 冻结首列
worksheet.views = [{ state: "frozen", xSplit: 1 }];

// 同时冻结
worksheet.views = [{ state: "frozen", xSplit: 1, ySplit: 1 }];
```

### 页面设置

```typescript
worksheet.pageSetup = {
  paperSize: 9, // A4
  orientation: "landscape",
  fitToPage: true,
  fitToWidth: 1,
  fitToHeight: 0,
  margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75 }
};

// 打印区域
worksheet.pageSetup.printArea = "A1:G20";

// 打印标题（每页重复第 1-2 行）
worksheet.pageSetup.printTitlesRow = "1:2";
```

### 工作表保护

```typescript
await worksheet.protect("password123", {
  selectLockedCells: true,
  selectUnlockedCells: true,
  formatCells: false,
  insertRows: false,
  deleteRows: false,
  sort: true,
  autoFilter: true
});
```

### 批注

```typescript
worksheet.getCell("A1").note = "简单批注";

worksheet.getCell("B1").note = {
  texts: [{ text: "作者：", font: { bold: true } }, { text: "这是一个富文本批注" }]
};
```

### 自动适应列宽

```typescript
worksheet.columns.forEach(column => {
  column.width = column.values
    ? Math.max(...column.values.map(v => String(v ?? "").length)) + 2
    : 10;
});
```

## 图表

ExcelTS 提供结构化图表 API、模板 raw XML 保留，以及确定性的预览渲染器。目标不是只保留图表 XML，而是能直接创建、修改、导出图表预览。

> **启用方式：** 图表功能是 opt-in 的，不会增大不使用图表的 bundle。在使用任何图表 API（`addChart`、`addLineChart`、图表加载/写入等）前，需在启动时调用一次 `installChartSupport()`：
>
> ```typescript
> import { installChartSupport } from "@cj-tech-master/excelts/chart";
> installChartSupport(); // 启动时调用一次
> ```
>
> 不调用此函数时，`worksheet.addChart()` 和 `writeFile()` 中的图表序列化会抛错。

> 完整可运行示例位于 [`src/modules/excel/examples/charts.ts`](examples/charts.ts)，涵盖 70+ 张图表——包含所有 classic 与 ChartEx 类型、各种 preset、combo/pivot/chartsheet 布局，并导出 SVG/PNG/PDF 预览。运行：`pnpm exec tsx src/modules/excel/examples/charts.ts`。

### Classic 图表

```typescript
const ws = workbook.addWorksheet("Sales");
ws.addRows([
  ["Month", "Revenue", "Profit"],
  ["Jan", 120, 32],
  ["Feb", 180, 49],
  ["Mar", 160, 41]
]);

ws.addChart(
  {
    type: "bar",
    barDir: "col",
    grouping: "clustered",
    title: "Revenue",
    series: [
      {
        name: "Revenue",
        categories: "Sales!$A$2:$A$4",
        values: "Sales!$B$2:$B$4",
        dataLabels: { showVal: true },
        trendline: { type: "linear", lineDash: "dash" },
        errorBars: { type: "fixedVal", value: 5 }
      }
    ],
    categoryAxis: { title: "Month" },
    valueAxis: { title: "USD", min: 0 }
  },
  "E1:M16"
);
```

### 预设和快捷 API

```typescript
import {
  EXCEL_CHART_PRESETS,
  EXCEL_CHART_EX_PRESETS,
  applyChartPreset,
  applyChartExPreset
} from "@cj-tech-master/excelts/chart";

// 99 个 classic preset + 10 个 ChartEx preset（对齐 Excel UI 别名）
ws.addPresetChart("col3DConeStacked100", { series: [{ values: "Sales!$B$2:$B$4" }] }, "E1:M16");
ws.addPresetChartEx(
  "boxAndWhisker",
  { series: [{ values: "Samples!$A$2:$A$50" }] },
  "N1:V16"
);

// 按类型的快捷方法 —— `type` 字段自动带入
ws.addColumnChart({ series: [...] }, "E18:M32");
ws.addBarChart({ series: [...] }, "E34:M48");
ws.addLineChart({ series: [...] }, "E50:M64");
ws.addAreaChart({ series: [...] }, "E66:M80");
ws.addPieChart({ series: [...] }, "P1:X16");
ws.addDoughnutChart({ series: [...] }, "P18:X32");
ws.addScatterChart({ series: [...] }, "P34:X48");
ws.addBubbleChart({ series: [...] }, "P50:X64");
ws.addRadarChart({ series: [...] }, "P66:X80");
ws.addStockChart({ series: [...] }, "AA1:AI16");
ws.addSurfaceChart({ series: [...] }, "AA18:AI32");
// ChartEx 快捷方法
ws.addHistogramChart({ series: [...] }, "AA34:AI48");
ws.addParetoChart({ series: [...] }, "AA50:AI64");
ws.addWaterfallChart({ series: [...] }, "AA66:AI80");
ws.addFunnelChart({ series: [...] }, "AK1:AS16");
ws.addTreemapChart({ series: [...] }, "AK18:AS32");
ws.addSunburstChart({ series: [...] }, "AK34:AS48");
ws.addBoxWhiskerChart({ series: [...] }, "AK50:AS64");
ws.addRegionMapChart({ series: [...] }, "AK66:AS80");

console.log(EXCEL_CHART_PRESETS.length, EXCEL_CHART_EX_PRESETS.length); // 99, 10
```

支持从 JS 数组或 Excel Table 构造 series：

```typescript
// 对象数组 → 图表：自动把行写入 worksheet，再按绝对引用生成 series
ws.addChartFromRows(
  [
    { day: "Mon", visits: 312 },
    { day: "Tue", visits: 400 },
    { day: "Wed", visits: 280 }
  ],
  { type: "bar", barDir: "col", x: "day", y: "visits", startCell: "A1" },
  "C1:K16"
);

// 柱状图快捷 —— 等价于上面的 `type: "bar", barDir: "col"`
ws.addColumnChartFromRows(rows, { x: "quarter", y: "revenue", startCell: "A1" }, "C1:K16");

// Excel Table → 图表：默认用 structured reference (`Table1[Col]`)，
// 表格扩行时图表自动跟进
const table = ws.addTable({ name: "Kpi", ref: "A1", headerRow: true, columns: [...], rows: [...] });
ws.addChartFromTable(
  table,
  { type: "bar", barDir: "col", categoryColumn: "Month", valueColumns: ["Revenue", "Profit"] },
  "F1:N18"
);

// ChartEx 对应的 helper
ws.addChartExFromRows(rows, { type: "histogram", x: "bucket", y: "count" }, "AA1:AI18");
ws.addChartExFromTable(
  table,
  { type: "funnel", categoryColumn: "Stage", valueColumns: ["Users"] },
  "AA20:AI40"
);

// 更底层的 range helper —— 生成带绝对引用的 series 对象
const s = ws.seriesFromColumns({
  categories: "Sales!$A$2:$A$7",
  values: "Sales!$B$2:$B$7",
  name: "Revenue"
});
ws.addChart({ type: "line", series: [s] }, "A20:I35");
```

### 组合图、ChartEx、数据透视图、图表工作表

```typescript
ws.addComboChart(
  {
    groups: [
      {
        type: "bar",
        barDir: "col",
        series: [{ name: "Revenue", categories: "Sales!$A$2:$A$4", values: "Sales!$B$2:$B$4" }]
      },
      {
        type: "line",
        useSecondaryAxis: true,
        series: [{ name: "Profit", categories: "Sales!$A$2:$A$4", values: "Sales!$C$2:$C$4" }]
      }
    ],
    title: "Revenue vs Profit",
    dataTable: { showKeys: true, showHorzBorder: true, showVertBorder: true }
  },
  "N1:V16"
);

// ChartEx —— Office 2016+ 的现代图表类型
// (histogram/pareto/waterfall/funnel/treemap/sunburst/boxWhisker/regionMap)
// 每个类型都有快捷方法；需要完全控制时直接传 `AddChartExOptions` 给 `addChartEx`
ws.addHistogramChart(
  { series: [{ name: "Distribution", values: "Sales!$B$2:$B$4" }], binning: { binType: "auto" } },
  "N18:V32"
);
ws.addWaterfallChart(
  {
    title: "Revenue waterfall",
    categories: "Sales!$A$2:$A$7",
    series: [{ name: "Delta", values: "Sales!$C$2:$C$7", subtotals: [0, 5] }],
    layout: { connectorLines: true }
  },
  "N34:V48"
);
ws.addTreemapChart(
  {
    categories: "Hier!$C$2:$C$10",
    series: [
      {
        name: "Sales",
        values: "Hier!$D$2:$D$10",
        hierarchy: ["Hier!$A$2:$A$10", "Hier!$B$2:$B$10"]
      }
    ],
    layout: { parentLabelLayout: "banner" }
  },
  "N50:V64"
);

// 数据透视图 —— 选项和 classic chart 相同，再加上对 pivot table 的引用；
// `pivotChartOptions` 控制 drop-zone 显示、打开时刷新、以及 Office 2014
// 引入的展开/收起字段按钮
const pivot = ws.addPivotTable({ sourceTable: src, rows: ["Region"], values: ["Revenue"] });
ws.addPivotChart(
  pivot,
  {
    type: "bar",
    barDir: "col",
    series: [{ name: "Revenue", categories: "Src!$A$2:$A$9", values: "Src!$D$2:$D$9" }],
    pivotChartOptions: {
      dropZonesVisible: true,
      dropZoneFilter: true,
      dropZoneCategories: true,
      dropZoneData: true,
      refreshOnOpen: true,
      showExpandCollapseFieldButtons: true
    }
  },
  "F1:N20"
);
ws.addPivotComboChart(pivot, { groups: [...] }, "F22:N40");

// Chartsheet —— 独占一个 sheet tab 的整页图。支持
// `AddChartOptions` / `AddComboChartOptions` / `AddChartExOptions` 任一形式
workbook.addChartsheet("Revenue Chart", {
  tabSelected: true,
  zoomToFit: true,
  chart: { type: "bar", series: [...] }
});

workbook.addPivotChartsheet("Pivot Dashboard", pivot, {
  chart: { type: "line", showMarker: true, series: [...] }
});
```

### 锚点形式

```typescript
// 字符串形式的 A1 区间 (two-cell anchor, 最常用)
ws.addChart({ type: "bar", series: [...] }, "A1:H15");

// 两格锚点 —— 显式 row/col 坐标
ws.addChart(options, { tl: { col: 1, row: 2 }, br: { col: 8, row: 17 } });

// 单格锚点 —— 固定在一个 cell，尺寸按 EMU 给（5×3 英寸；914400 EMU = 1 英寸）
ws.addChart(options, {
  tl: { col: 1, row: 19 },
  ext: { cx: 5 * 914400, cy: 3 * 914400 },
  editAs: "oneCell"
});

// 绝对锚点 —— 位置和大小都按 EMU 给，不跟随行高列宽
ws.addChart(options, {
  pos: { x: 914400, y: 36 * 914400 },
  ext: { cx: 5 * 914400, cy: 3 * 914400 },
  editAs: "absolute"
});
```

### 高级 series 格式化

```typescript
ws.addChart(
  {
    type: "line",
    title: {
      paragraphs: [
        { runs: [{ text: "Q2 ", properties: { bold: true, size: 1600 } }, { text: "Performance" }] }
      ]
    },
    series: [
      {
        name: "Revenue",
        categories: "Sales!$A$2:$A$7",
        values: "Sales!$B$2:$B$7",
        line: "4472C4",
        lineWidth: 2.5,
        lineDash: "solid",
        marker: { symbol: "circle", size: 8, fill: "4472C4", border: "FFFFFF" },
        trendline: {
          type: "linear",
          displayEq: true,
          displayRSqr: true,
          forward: 1,
          line: "ED7D31",
          lineDash: "dash"
        },
        errorBars: {
          direction: "y",
          barDir: "both",
          type: "percentage",
          value: 10
        },
        dataLabels: { showVal: true, position: "t", numFmt: "$#,##0" },
        // 单点 override
        dataPoints: [
          { index: 0, fill: "C00000" },
          { index: 5, fill: "70AD47", marker: { symbol: "diamond", size: 10 } }
        ]
      }
    ],
    categoryAxis: { title: "Month", textRotation: -45 },
    valueAxis: {
      title: "Revenue",
      numFmt: "$#,##0",
      min: 0,
      logBase: 10,
      majorGridlines: true,
      displayUnits: "thousands",
      displayUnitsLabel: "× 1 000"
    },
    legendOptions: {
      entries: [{ index: 1, hidden: true }],
      txPr: { size: 900, color: { srgb: "595959" } }
    },
    plotAreaOptions: { spPr: { fill: "FAFAFA", border: "D9D9D9" } }
  },
  "A1:L20"
);

// 图片填充 (pictureFill) —— 柱状条用图像填充。输入可接受
// 原始 Uint8Array、`data:` URL、base64 字符串、
// `{ workbookImageId }` 句柄，或结构化 `ChartPictureFillImageData`
ws.addChart(
  {
    type: "bar",
    barDir: "col",
    series: [
      {
        name: "Revenue",
        categories: "Sales!$A$2:$A$7",
        values: "Sales!$B$2:$B$7",
        pictureFill: { image: pngBytes, fillMode: "stretch" }
      }
    ]
  },
  "N1:V16"
);
```

### 图表 style

```typescript
// 旧版 2007/2010 built-in style (1..48)，写 `<c:style val="N"/>`
chart.setStyle(42);
chart.setBuiltInStyle(42); // xlsxwriter 风格的别名

// 现代 Office 2013+ sidecar —— 完整 styleN.xml + colorsN.xml
// 通过 `addChart` options 带入，或后续复制
ws.addChart(
  {
    type: "bar",
    series: [...],
    chartStyle: {
      id: 201,
      elements: {
        chartArea: { fillRefIdx: 1, lnRefIdx: 1, effectRefIdx: 0, fontRefIdx: "minor" },
        title: { fontRefIdx: "major" }
      }
    },
    chartColors: {
      method: "cycle",
      id: 10,
      colors: [{ srgb: "4472C4" }, { srgb: "ED7D31" }, { srgb: "A5A5A5" }]
    }
  },
  "A1:H15"
);
```

### 预览导出

```typescript
import { chartToPdf } from "@cj-tech-master/excelts/pdf";

const chart = ws.getCharts()[0];

// SVG / PNG 预览 —— PNG 返回 Promise，因为 Node 侧的 rasterizer 是异步的
const svg = chart.toSVG({ width: 800, height: 450, backgroundColor: "transparent" });
const png = await chart.toPNG({ width: 800, height: 450, scale: 2, dpi: 192 });

// 单图单页 PDF —— classic chart 走矢量 (文本可选中、缩放不糊)；
// ChartEx 能走矢量时也走矢量，必要时可通过 `forceRaster: true` 强制栅格
const pdf = await chartToPdf(chart, {
  title: "Revenue",
  width: 640,
  height: 400,
  margin: 36
});

// 从外部查询矢量/栅格路径选择：
import { canRenderChartExAsVectorPdf } from "@cj-tech-master/excelts/chart";
if (chart.chartExModel) {
  console.log(canRenderChartExAsVectorPdf(chart.chartExModel));
}
```

预览渲染是确定性、零依赖的。浏览器 PNG 使用 canvas，Node.js PNG 使用内置基础 rasterizer。它会绘制核心图表几何、坐标轴、次坐标轴、坐标轴标题、图例、标签、marker、趋势线和误差线，适合缩略图、测试、服务端预览；不是 Excel/Aspose 级 pixel-perfect 渲染器，也不是 Excel 完全一致的布局引擎。ChartEx `regionMap` 预览对已知国家使用内置小型 centroid 表和投影计算，对未知标签使用确定性 tile fallback；这是地理近似预览，不是 GIS/行政边界地图渲染器。

### 模板保真

加载后的图表 XML 如果不修改，会按字节原样保留。安全的高层 mutation 会只 patch 已知 XML 块，并保留未知扩展：

- classic charts：标题、图例、series 引用、series 格式、marker、data point、数据标签、趋势线、误差线、坐标轴、plot layout
- ChartEx charts：chart data、标题、图例、auto-title deletion、chart/plot 形状、plot-region layout、series 可见性/名称/axis 绑定、series data refs、layoutPr（含 `extLst` passthrough）、数据标签、data point、坐标轴
- 不安全的结构性 mutation 会回退到结构化重渲染

编辑已加载模板图表时，如果希望优先局部 patch，可使用 `chart.mutate(model => { ... }, { preferRawPatch: true })`。

严格模板工作流可以使用 `requireRawPatch: true`：如果 mutation 不能安全局部 patch，就直接失败，而不是回退结构化重渲染。

```typescript
chart.mutate(
  model => {
    model.chart.plotArea.chartTypes[0].series[0].val = {
      numRef: { formula: "Sales!$B$2:$B$100", cache: { points: [] } }
    };
  },
  { preferRawPatch: true, requireRawPatch: true }
);
```

这提供的是“支持的 patch 类型必须保留 raw 模板 XML，否则抛错”的硬保证。它不声称任意未知 OOXML 都能安全 mutation；不支持的结构性编辑会在 `requireRawPatch` 开启时被拒绝。

也可以在一次写出中对所有加载自模板的 chart/chartEx part 启用这个规则：

```typescript
await workbook.xlsx.writeBuffer({ templateMode: "strict" });
// 或
await workbook.xlsx.writeBuffer({ strictTemplateMode: true });
```

严格模板模式只影响从已有 workbook 加载并被编辑过的图表 part；新建图表仍按结构化 XML 正常写出。

### Oracle 和语料库测试

仓库提供可选的真实应用验证 harness。默认关闭，因为它们需要外部程序或私有 fixture 语料。

这些 harness 中生成的每个工作簿也会先跑 OOXML 包结构审计。审计会检查必要 part 的 content type、relationship target、重复 relationship ID、chart/ChartEx/drawing/chartsheet 基础结构、ChartEx data/axis 引用，以及 ChartEx external-data relationship ID，让常见 Excel “修复记录”类问题在 CI 中提前失败。启用 Office/LibreOffice open-validation 后，如果命令日志包含 repair/corruption/error 文本，测试会按硬失败处理。

```bash
# LibreOffice 视觉/PDF 导出 oracle
EXCELTS_LIBREOFFICE_VISUAL_ORACLE=1 LIBREOFFICE_BIN=/path/to/soffice \
  pnpm exec vitest run src/modules/excel/__tests__/chart-oracle.integration.test.ts

# LibreOffice 打开/转换验证生成的工作簿
EXCELTS_LIBREOFFICE_OPEN_VALIDATION=1 LIBREOFFICE_BIN=/path/to/soffice \
  pnpm exec vitest run src/modules/excel/__tests__/chart-oracle.integration.test.ts

# 专有 Office/Aspose 风格 CLI 验证 hook。命令参数通过
# EXCELTS_OFFICE_OPEN_ARGS 提供 {input} 和 {outDir} 占位符。
EXCELTS_OFFICE_OPEN_VALIDATION=1 EXCEL_OFFICE_BIN=/path/to/validator \
EXCELTS_OFFICE_OPEN_ARGS="--open {input} --outdir {outDir}" \
  pnpm exec vitest run src/modules/excel/__tests__/chart-oracle.integration.test.ts

# 企业语料 round-trip harness
EXCELTS_ENTERPRISE_CORPUS_DIR=/path/to/private/xlsx-corpus \
  pnpm exec vitest run src/modules/excel/__tests__/chart-oracle.integration.test.ts

# 企业语料 + LibreOffice 打开验证
EXCELTS_ENTERPRISE_CORPUS_DIR=/path/to/private/xlsx-corpus \
EXCELTS_CORPUS_LIBREOFFICE_OPEN_VALIDATION=1 LIBREOFFICE_BIN=/path/to/soffice \
  pnpm exec vitest run src/modules/excel/__tests__/chart-oracle.integration.test.ts
```

语料目录可放一个可选 `manifest.json` 标记预期结构：

```json
{
  "entries": [
    {
      "path": "charts/sales-dashboard.xlsx",
      "source": "Excel 365",
      "expectCharts": true,
      "expectChartEx": true,
      "openValidation": true
    },
    {
      "path": "pivot/pivot-chart.xlsx",
      "source": "Excel 365",
      "expectCharts": true,
      "expectPivotTables": true
    }
  ]
}
```

Excel、WPS、Aspose 可以用相同模式接入 CI：把生成工作簿导出成 PDF/图片，再和批准件对比。ExcelTS 自身保持零依赖，不内置专有渲染器。内置审计是结构 gate，不能替代真实 Office 的视觉/open-repair 验证。

### 兼容矩阵

| 能力           | 状态                                                                                                                                                                                                                                                                                                                                                                                                  |
| -------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Classic charts | bar、bar3D、line、line3D、pie、pie3D、doughnut、area、area3D、scatter、bubble、radar、stock、surface、surface3D、ofPie（见 3D 说明）                                                                                                                                                                                                                                                                  |
| ChartEx        | sunburst、treemap、waterfall、funnel、histogram、pareto、boxWhisker、regionMap（见 regionMap 说明）                                                                                                                                                                                                                                                                                                   |
| 高级图表能力   | 组合图、secondary axes、markers、data labels（`DataLabelPosition`、饼图 leader line、柱/线图碰撞避让）、trendlines、error bars、manual plot layout（edge 模式）、chartsheets                                                                                                                                                                                                                          |
| 数据透视图     | classic pivot chart source metadata、field buttons/filter metadata、pivot chartsheets                                                                                                                                                                                                                                                                                                                 |
| 预设           | 99 个 classic preset + 10 个 ChartEx preset —— cone/cylinder/pyramid、scatter variants、stock、surface/contour、exploded pie/doughnut、histogram/pareto/waterfall/funnel/treemap/sunburst/boxWhisker/regionMap (通过 `EXCEL_CHART_PRESETS` / `EXCEL_CHART_EX_PRESETS` 访问)                                                                                                                           |
| 渲染           | 确定性 SVG、浏览器 PNG、Node PNG fallback（支持文本 `rotate`）、PDF drawing bridge 已与 SVG 完全对齐：labels/markers/errorBars/trendlines/leader lines、文本 anchor/rotation/color/fontFamily（`bold`/`italic` 来自 `txPr/a:latin`）、radar/area/bubble/bar3D 真 alpha 经由 `PdfColor.a` → `/ExtGState`；文本宽度来自 `@excel/utils/text-metrics`（Calibri/Arial/Times 等 9 字体 + ~230 factor 兜底） |
| 商业级差距     | Excel 级精确渲染、任意未知 XML mutation、完整真实文件兼容矩阵仍需要外部 oracle 验证                                                                                                                                                                                                                                                                                                                   |

**3D 说明：** `bar3D`、`line3D`、`pie3D`、`area3D`、`surface3D` 以及 `view3D` 的旋转/透视/厚度元数据会**完整保留在 XML** 中，`Scene3D` / `View3D` / `ShapeProperties3D` 的解析器全部可用，round-trip 与 Excel 再打开不受影响。内置 SVG/PNG/PDF 预览有意把所有 3D 变体渲染为对应的 2D 形态（bar3D 仅加一个固定 6 像素的 depth 提示）——没有投影矩阵、没有光照、没有深度排序。这是 OOXML 保真的预览而非 3D 渲染引擎；需要商业级 3D 输出请用 Excel 或 LibreOffice。

**字体与 CJK：** `PdfDocumentBuilder` 在页面含非 WinAnsi 字符且未显式 `embedFont` 时会自动发现系统字体（与 `excelToPdf` 相同机制）。需要跨机器字节稳定的输出请调用 `disableFontAutoDiscovery()`，或通过 `embedFont(ttfBytes)` 显式指定字体。注册 `onWarning(handler)` 可按 build 收到一次诊断：每个未识别的 `fontFamily`（退回 Helvetica metrics）一条，每次 build 中出现但无字体覆盖的非 WinAnsi 字符（将渲染 Type3 NOTDEF 方块）一条。

**最小 PDF surface：** `ChartPdfDrawingSurface.drawPath?` 与 `drawCircle?` 为可选。surface 缺 `drawPath` 时，饼 / 甜甜圈 / ofPie 的扇形轮廓降级为 `drawLine` 多段线描边（形状保留，填色丢失）；area 与 radar-filled 的 fill 会丢失但周围的描边依然输出；marker 按 circle→rect→line 逐级降级。`PdfPageBuilder` / `PdfEditorPage` 两者都实现了完整接口，因此这只影响自定义 surface 的使用者。

**regionMap 说明：** ChartEx `regionMap` 预览内置 ~180 国的质心表与四种真投影公式（`mercator`、`miller`、`albers` 等积圆锥、`robinson`）。默认是质心打点的地理预览，未匹配标签回退到六边形 tile 网格。需要真国界多边形时，通过渲染选项 `regionMap: { topology, objectName, match, projection }` 传入 TopoJSON——渲染器会解码 features、按 `feature.id` 或 `feature.properties.<key>` 匹配标签、并绘制分级 choropleth。库本身保持零数据内置——调用方自行加载 `world-atlas`/`natural-earth` 文件。同一条三阶段管线（TopoJSON → 质心预览 → 六边形 tile 回退）**SVG 与矢量 PDF 两条后端都实现了**，`chartToPdf` 会把同一个 `regionMap` option 透传给 `drawChartExPdf`。详见 `src/modules/excel/chart/topojson.ts` 与导出的 `RegionMapDataOptions` / `TopologyLike` 类型。

**ChartEx PDF 说明：** classic charts 走 `drawChartPdf` 矢量路径。ChartEx 所有 layout 现在也全部走 `drawChartExPdf` 矢量路径：

- **矢量路径（默认）** — `sunburst`、`treemap`、`waterfall`、`funnel`、`histogram`、`pareto`、`boxWhisker`、`regionMap` 八种类型全部走 `drawChartExPdf`，与 SVG 渲染器共用几何 collectors，两条后端像素级等价（sunburst 圆弧用 cubic Bézier 近似，最大误差 ≤ 0.03%）。`regionMap` 复用同一套 TopoJSON 解码 + 投影 + 质心表。文本可选中、缩放不糊、PDF 文件小。唯一的故意视觉差异：SVG 的圆角框（`rx="14"`）在 PDF 中是直角框（`drawRect` 不暴露圆角半径）。
- **栅格按需** — 任何 ChartEx 类型都可通过 `chartToPdf(chart, { forceRaster: true })` 强制走栅格路径，适合需要与 SVG 预览像素完全一致而不在意文本可选中的场景。

统一入口仍是 `chartToPdf(chart, options)`（从 `@cj-tech-master/excelts/pdf`）—— 自动选择路径，可用 `{ forceRaster: true }` 强制栅格（例如需要与 SVG 预览像素一致而非可选中文本时）；或直接检查 `canRenderChartExAsVectorPdf(model)` 了解路由决策。

**内置 chart style：** `chart.setStyle(1..48)`（别名 `chart.setBuiltInStyle(1..48)`）在经典 chart 上写 `<c:style val="N"/>`，语义对齐 xlsxwriter 的 `chart.set_style(N)`。这是对应 2007/2010 style 目录的轻量开关。需要 Office 2013+ 现代 style（带 `styleN.xml` / `colorsN.xml` sidecar）时，用 `worksheet.addChart({ …, chartStyle: ChartStyleModel })`。

**3D 渲染边界（刻意不做）：** 除 `bar3D` 的提示性 depth 外，我们有意**不**实现:

- 任何 3D chart 类型的真 3D 投影（rotX/rotY/perspective → 矩阵 + 深度排序 + 光照）
- surface3D 的三角网格/线框/等高线渲染

这些特性对预览级渲染器价值比不上多周投入;需要 Excel 级 3D 输出请通过 Excel 或 LibreOffice round-trip。所需元数据（`Scene3D`、`View3D`、`ShapeProperties3D`）已在 XML 层完整 round-trip。

**严格模板模式：** 写入选项 `{ templateMode: "strict" }`（或 `{ strictTemplateMode: true }`）会拒绝任何会触发结构化重建的 chart / ChartEx 修改。当必须重建时，错误信息会列出 parser 观察到的未结构化 XML 元素路径（同时可通过 `ChartExModel.unknownElements` 读取），避免厂商扩展在加载的模板上被静默丢弃。

**测试范围边界（本库**不**测试的内容）：**

- **没有像素级视觉 diff。** 预览输出通过 SVG 结构断言 + PNG 头/签名 hash 验证 — 真正的 RMS/SSIM 像素 diff 需要打包 PNG 解码器和 diff 算法，而且预览本身明确**不是**像素对齐 Excel 的（见上面的渲染说明）。若工作流需要像素级对齐 Excel，请用 `chartToPdf(chart)` 经 LibreOffice headless PDF 导出后做比较。
- **仓库内不含 Excel/WPS/Aspose 真实生成的 fixture。** `src/modules/excel/__tests__/data/` 的每个 .xlsx 要么是 ExcelTS 自身生成要么是为回归测试手工写的最小样本。需要宿主应用兼容性覆盖时使用 opt-in 的 `EXCELTS_ENTERPRISE_CORPUS_DIR` — 指向一个用户自备的 fixture 目录，`chart-oracle.integration.test.ts` 会逐个审计。参考 `docs/enterprise-corpus-manifest.example.json` 的 manifest 格式，以及 `scripts/compatibility-report.ts`（`pnpm run compatibility:report`）的报告生成器。
- **CI 没有自动化的 Excel / WPS 运行时。** CI 的 open-validation 仅 gate 到 LibreOffice 上；Excel / WPS 二进制不出现在任何 CI runner 中，GUI 驱动的这两款 app 的 validation 超出范围。`EXCELTS_OFFICE_OPEN_VALIDATION` + `EXCELTS_OFFICE_OPEN_ARGS` 钩子允许自建的安装了 Office 的 runner 按同一 pattern 参与这一检查。

相比 ExcelJS，ExcelTS 有原生图表创建和编辑。相比 xlsx-populate，ExcelTS 在安全场景保留模板 XML 的同时提供结构化 chart API。相比 XlsxWriter/openpyxl/excelize，ExcelTS 提供 TypeScript/浏览器支持、ChartEx、数据透视图元数据、图表工作表和预览渲染入口。

### 迁移指南

详细 API 映射在独立文档中：

- **[`docs/FROM_EXCELJS.md`](../../../docs/FROM_EXCELJS.md)** — ExcelJS 没有原生 chart 创建 API；本文档说明如何把"模板原样导出"和"手工编辑 chart XML"的流程改为结构化的 `addChart` / `mutate` 调用，并介绍 ExcelJS 没有的预览渲染助手。
- **[`docs/FROM_XLSXWRITER.md`](../../../docs/FROM_XLSXWRITER.md)** — 逐条 cheat sheet，覆盖 XlsxWriter (Python)、openpyxl (Python)、excelize (Go)。内容包括 `set_style`、`add_series`、`set_x_axis`/`set_y_axis`、chart 尺寸、以及这些库都没有的现代 ChartEx 类型（sunburst/waterfall/funnel/boxWhisker/regionMap）。
- 企业语料验证 manifest 示例：[`docs/enterprise-corpus-manifest.example.json`](../../../docs/enterprise-corpus-manifest.example.json)。

## PDF 导出

零外部依赖将任意工作簿导出为 PDF：

```typescript
import { Workbook, excelToPdf } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("报告");
sheet.columns = [
  { header: "产品", key: "product", width: 20 },
  { header: "收入", key: "revenue", width: 15 }
];
sheet.addRow({ product: "小工具", revenue: 1000 });
sheet.getColumn("revenue").numFmt = "$#,##0.00";

const pdf = await excelToPdf(workbook, {
  showGridLines: true,
  showPageNumbers: true,
  title: "销售报告"
});

// Node.js
import { writeFileSync } from "fs";
writeFileSync("report.pdf", pdf);

// 浏览器
const blob = new Blob([pdf], { type: "application/pdf" });
window.open(URL.createObjectURL(blob));
```

### XLSX 转 PDF

```typescript
const workbook = new Workbook();
await workbook.xlsx.readFile("input.xlsx");
const pdf = await excelToPdf(workbook);
```

### PDF 加密

```typescript
const pdf = await excelToPdf(workbook, {
  encryption: {
    ownerPassword: "admin",
    userPassword: "reader",
    permissions: { print: true, copy: false }
  }
});
```

### Unicode / CJK 字体嵌入

```typescript
import { readFileSync } from "fs";
const pdf = await excelToPdf(workbook, {
  font: readFileSync("NotoSansSC-Regular.ttf")
});
```

## CSV 导入/导出

```typescript
import { Workbook } from "@cj-tech-master/excelts";
import fs from "fs";

const workbook = new Workbook();

// Node.js：读写 CSV 文件
await workbook.readCsvFile("data.csv");
await workbook.writeCsvFile("output.csv");

// 从流读取 CSV
await workbook.readCsv(fs.createReadStream("data.csv"), { sheetName: "导入数据" });

// 写入 CSV 到流
await workbook.writeCsv(fs.createWriteStream("output.csv"));

// 写入 CSV 到字符串 / 字节
const csvText = workbook.writeCsv();
const bytes = await workbook.writeCsvBuffer();

// 浏览器：从字符串/ArrayBuffer/File 读取
await workbook.readCsv(csvString);
await workbook.readCsv(arrayBuffer);
```

## Markdown 导入/导出

```typescript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();

// 读取 Markdown 表格
workbook.readMarkdown("| 姓名 | 年龄 |\n| --- | --- |\n| Alice | 30 |");
await workbook.readMarkdownFile("table.md");

// 写入 Markdown
const mdText = workbook.writeMarkdown();
await workbook.writeMarkdownFile("output.md");
const bytes = workbook.writeMarkdownBuffer();
```

## 流式 API

### 流式读取器

以最小内存使用量读取大型 XLSX 文件：

```typescript
import { WorkbookReader } from "@cj-tech-master/excelts";

const reader = new WorkbookReader("large-file.xlsx", {
  worksheets: "emit",
  sharedStrings: "cache",
  hyperlinks: "ignore",
  styles: "ignore"
});

for await (const worksheet of reader) {
  console.log(`正在读取：${worksheet.name}`);
  for await (const row of worksheet) {
    console.log(row.values);
  }
}
```

### 流式写入器

逐行写入大型 XLSX 文件：

```typescript
import { WorkbookWriter } from "@cj-tech-master/excelts";

const workbook = new WorkbookWriter({
  filename: "output.xlsx",
  useSharedStrings: true,
  useStyles: true
});

const sheet = workbook.addWorksheet("数据");
for (let i = 0; i < 1000000; i++) {
  sheet.addRow([`行 ${i}`, i, new Date()]).commit();
}

sheet.commit();
await workbook.commit();
```

### Web Streams（Node.js 22+ 和浏览器）

```typescript
import { WorkbookWriter, WorkbookReader } from "@cj-tech-master/excelts";

// 写入到 Web WritableStream
const chunks: Uint8Array[] = [];
const writable = new WritableStream({
  write(chunk) {
    chunks.push(chunk);
  }
});

const writer = new WorkbookWriter({ stream: writable });
const sheet = writer.addWorksheet("Sheet1");
sheet.addRow(["姓名", "分数"]).commit();
sheet.addRow(["Alice", 98]).commit();
await sheet.commit();
await writer.commit();

// 从 Web ReadableStream 读取
const bytes = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
let offset = 0;
for (const c of chunks) {
  bytes.set(c, offset);
  offset += c.length;
}

const readable = new ReadableStream({
  start(ctrl) {
    ctrl.enqueue(bytes);
    ctrl.close();
  }
});

const reader = new WorkbookReader(readable, { worksheets: "emit" });
for await (const ws of reader) {
  for await (const row of ws) {
    console.log(row.values);
  }
}
```

## 浏览器支持

### 使用打包工具（Vite、Webpack、Rollup、esbuild）

```typescript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("Sheet1");
sheet.getCell("A1").value = "你好，浏览器！";

const buffer = await workbook.xlsx.writeBuffer();
const blob = new Blob([buffer], {
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
});
const url = URL.createObjectURL(blob);
```

### 使用 Script 标签

```html
<script src="https://unpkg.com/@cj-tech-master/excelts/dist/iife/excelts.iife.min.js"></script>
<script>
  const { Workbook } = ExcelTS;
  const wb = new Workbook();
</script>
```

### 浏览器注意事项

- 使用 `xlsx.load(arrayBuffer)` 代替 `xlsx.readFile()`
- 使用 `xlsx.writeBuffer()` 代替 `xlsx.writeFile()`
- PDF 导出完全支持
- CSV 和 Markdown 操作完全支持
- 工作表密码保护使用纯 JS SHA-512

## 工具导出

```typescript
import {
  // 日期转换
  dateToExcel,
  excelToDate,
  DateParser,
  DateFormatter,

  // 二进制工具
  base64ToUint8Array,
  uint8ArrayToBase64,
  concatUint8Arrays,
  toUint8Array,
  stringToUint8Array,
  uint8ArrayToString,

  // XML 工具
  xmlEncode,
  xmlDecode,
  xmlEncodeAttr,
  validateXmlName,

  // PDF 导出
  pdf,
  excelToPdf,
  PageSizes,
  PdfError,
  isPdfError,

  // 错误处理
  BaseError,
  ExcelError,
  toError,
  errorToJSON,
  getErrorChain,
  getRootCause
} from "@cj-tech-master/excelts";
```

## 示例

查看 [examples 目录](examples/) 获取覆盖所有功能的可运行代码：

- 工作簿创建、读取和复制
- 单元格样式、字体、边框、填充
- 公式、数据验证、条件格式
- 图片（JPEG、PNG）、超链接、批注
- 带自动筛选和汇总的表格
- 合并单元格、冻结窗格、页面设置
- 流式读取器和写入器
- Web Streams 集成
- PDF 导出
- 更多...
