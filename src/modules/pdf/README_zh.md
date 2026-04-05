# PDF 模块

功能完整的零依赖 PDF 引擎，纯 TypeScript 从零构建。可**独立使用** `pdf()` 函数，或通过 `excelToPdf()` 桥接 API 作为 **Excel 转 PDF 转换器**。

## 功能特性

- **零依赖** — 纯 TypeScript，无外部包
- **独立引擎** — `pdf()` 函数直接从数据生成 PDF，无需 Workbook
- **Excel 桥接** — `excelToPdf()` 转换 Excel 工作簿为 PDF
- **跨平台** — Node.js 和浏览器
- **完整样式** — 字体、颜色、边框、填充、对齐
- **富文本** — 单元格内多字体/颜色
- **TrueType 字体嵌入** — Unicode/CJK 文本支持
- **图片** — JPEG 和 PNG，支持透明度
- **加密** — 所有者/用户密码、权限控制
- **自动分页** — 重复标题行
- **每工作表页面设置** — 纸张大小、方向、边距

## 快速开始

```typescript
// 独立使用 — 从数组生成 PDF
import { pdf } from "@cj-tech-master/excelts/pdf";

const bytes = pdf([
  ["产品", "收入"],
  ["小工具", 1000],
  ["大工具", 2500]
]);

// Excel 转 PDF
import { Workbook, excelToPdf } from "@cj-tech-master/excelts";

const workbook = new Workbook();
await workbook.xlsx.readFile("input.xlsx");
const pdfBytes = excelToPdf(workbook, { showGridLines: true });
```

## 详细文档

完整 API 参考请查看 [英文文档](README.md)。

## 示例

查看 [examples 目录](examples/) 获取可运行代码。
