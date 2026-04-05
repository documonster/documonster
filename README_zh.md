# ExcelTS

[![构建状态](https://github.com/cjnoname/excelts/actions/workflows/ci.yml/badge.svg?branch=main&event=push)](https://github.com/cjnoname/excelts/actions/workflows/ci.yml)

现代 TypeScript Excel 工作簿管理器 — 读取、操作和写入 XLSX 与 JSON 格式的电子表格数据和样式。

## AI 支持

ExcelTS 专为 **AI 友好** 设计 — 简洁一致的 API，AI 编程助手可以轻松生成和使用。一个库处理 Excel、PDF、CSV、Markdown、XML、ZIP 和流式操作，效率最大化。每个模块都配有完整的文档和可运行的示例供 AI 学习。

## 关于本项目

ExcelTS 是一个零依赖的 TypeScript 电子表格和文档工具包：

- **零运行时依赖** — 纯 TypeScript，无外部包
- **七大模块** — Excel、PDF、CSV、Markdown、XML、Archive、Stream
- **跨平台** — Node.js 22+、Bun、Chrome 89+、Firefox 102+、Safari 14.1+
- **ESM 优先** — 原生 ES Modules，兼容 CommonJS，完整的 tree-shaking 支持

## 翻译

- [English](README.md)

## 模块

ExcelTS 由七个独立模块组成，每个模块都有自己的文档和可运行示例。

### Excel — XLSX/JSON 工作簿管理器

创建、读取和修改 Excel 电子表格，完整支持样式、公式、图片和流式处理。

- [文档](src/modules/excel/README.md) | [中文文档](src/modules/excel/README_zh.md)
- [示例](src/modules/excel/examples/)（40+ 可运行脚本）

### PDF — 零依赖 PDF 引擎

功能完整的 PDF 生成，支持字体嵌入、加密、图片和 Excel 转 PDF。

- [文档](src/modules/pdf/README.md) | [中文文档](src/modules/pdf/README_zh.md)
- [示例](src/modules/pdf/examples/)

### CSV — RFC 4180 解析器/格式化器

高性能 CSV 解析和格式化，支持流式处理、动态类型、数据生成和工作线程池。

- [文档](src/modules/csv/README.md) | [中文文档](src/modules/csv/README_zh.md)
- [示例](src/modules/csv/examples/)

### Markdown — GFM 表格解析器/格式化器

解析和格式化 GitHub 风格 Markdown 表格，支持对齐方式保留和工作簿集成。

- [文档](src/modules/md/README.md) | [中文文档](src/modules/md/README_zh.md)
- [示例](src/modules/md/examples/)

### XML — SAX/DOM 解析器、查询引擎、写入器

流式和缓冲式 XML 处理，含查询引擎、命名空间支持和双模式写入。

- [文档](src/modules/xml/README.md) | [中文文档](src/modules/xml/README_zh.md)
- [示例](src/modules/xml/examples/)

### Archive — ZIP/TAR 创建/读取/编辑

ZIP 和 TAR 归档创建、读取、编辑、流式处理、加密和压缩工具。

- [文档](src/modules/archive/README.md) | [中文文档](src/modules/archive/README_zh.md)
- [示例](src/modules/archive/examples/)

### Stream — 跨平台流式处理

兼容 Node.js 的 Readable/Writable/Transform/Duplex，在 Node.js 和浏览器中行为一致。

- [文档](src/modules/stream/README.md) | [中文文档](src/modules/stream/README_zh.md)
- [示例](src/modules/stream/examples/)

## 安装

```bash
npm install @cj-tech-master/excelts
```

## 快速开始

```typescript
import { Workbook } from "@cj-tech-master/excelts";

// 创建
const workbook = new Workbook();
const sheet = workbook.addWorksheet("Sheet1");
sheet.addRow(["姓名", "年龄"]);
sheet.addRow(["Alice", 30]);
await workbook.xlsx.writeFile("output.xlsx");

// 读取
const wb = new Workbook();
await wb.xlsx.readFile("output.xlsx");
wb.getWorksheet(1).eachRow((row, n) => console.log(n, row.values));
```

## 子路径导出

每个模块都可以作为独立的子路径导出使用：

```typescript
import { Workbook, WorkbookWriter } from "@cj-tech-master/excelts";
import { SaxParser, parseXml, XmlWriter, query } from "@cj-tech-master/excelts/xml";
import { zip, unzip, ZipArchive, compress } from "@cj-tech-master/excelts/zip";
import { parseCsv, formatCsv, CsvParserStream } from "@cj-tech-master/excelts/csv";
import { parseMd, formatMd, parseMdAll } from "@cj-tech-master/excelts/md";
import { Readable, pipeline, createTransform } from "@cj-tech-master/excelts/stream";
```

每个子路径均支持 `browser`、`import`（ESM）和 `require`（CJS）条件导出。

## 浏览器支持

ExcelTS 原生支持浏览器，现代打包工具**零配置**即可使用。

```typescript
// 打包工具（Vite、Webpack、Rollup、esbuild）— 直接导入
import { Workbook } from "@cj-tech-master/excelts";
const buffer = await new Workbook().addWorksheet("S1").workbook.xlsx.writeBuffer();
```

```html
<!-- Script 标签（无需打包工具） -->
<script src="https://unpkg.com/@cj-tech-master/excelts/dist/iife/excelts.iife.min.js"></script>
```

| 浏览器  | 最低版本           |
| ------- | ------------------ |
| Chrome  | 89+（2021年3月）   |
| Edge    | 89+（2021年3月）   |
| Firefox | 102+（2022年6月）  |
| Safari  | 14.1+（2021年4月） |
| Opera   | 75+（2021年3月）   |

对于不支持原生 `CompressionStream` API 的旧版浏览器，ExcelTS 自动使用内置的纯 JavaScript DEFLATE 实现 — 无需 polyfill。

## 系统要求

- **Node.js >= 22.0.0**
- 浏览器兼容性见上表

## 许可证

MIT 许可证 — 见 LICENSE。

第三方声明：[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)

## 链接

- [GitHub 仓库](https://github.com/cjnoname/excelts)
- [问题追踪](https://github.com/cjnoname/excelts/issues)
- [更新日志](CHANGELOG.md)
- [迁移指南](MIGRATION.md)
- [路线图](ROADMAP.md)
