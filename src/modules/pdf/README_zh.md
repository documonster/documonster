# PDF 模块

[English](README.md)

功能完整的零依赖 PDF 引擎，纯 TypeScript 从零构建。**写入**：`pdf()` 函数独立生成或 `excelToPdf()` 桥接 API。**读取**：`readPdf()` 从任意 PDF 提取文本、图片、标注、表单和元数据。

## 功能特性

### 写入

- **零依赖** — 纯 TypeScript，无外部包
- **PDF 2.0** — 生成现代 PDF 2.0 格式
- **独立引擎** — `pdf()` 函数直接从数据生成 PDF，无需 Workbook
- **Excel 桥接** — `excelToPdf()` 一行代码转换 Excel 工作簿为 PDF
- **跨平台** — Node.js 和浏览器同一 API
- **完整样式** — 字体、颜色、边框、填充、对齐、合并单元格
- **富文本** — 单元格内多字体/颜色混排
- **自动分页** — 水平/垂直分页，支持重复标题行
- **图片** — JPEG 和 PNG 嵌入，支持透明度
- **AES-256 加密** — 密码保护 (V=5, R=5)，权限控制
- **字体嵌入** — TrueType 字体子集化，支持 Unicode/CJK
- **页面设置** — 每工作表独立纸张大小、方向、边距、打印区域
- **Tree-Shakeable** — 未导入则不打包

### 读取

- **通用读取器** — 读取所有主流 PDF 版本 (1.0 至 2.0)
- **文本提取** — 完整文本重建，多栏检测，表格识别
- **多语言** — WinAnsi、MacRoman、CJK (ToUnicode CMap)、Identity-H/V、Symbol、ZapfDingbats
- **图片提取** — JPEG、JPEG2000、CCITT、JBIG2、raw/Flate，支持 SMask/alpha
- **标注提取** — 链接、批注、高亮、图章、自由文本等
- **表单字段** — AcroForm 提取：文本输入、复选框、单选按钮、下拉列表、签名
- **元数据** — Info 字典 + XMP（标题、作者、日期、页数、页面尺寸）
- **全加密格式** — RC4-40、RC4-128、AES-128、AES-256（读取所有版本）
- **容错** — 交叉引用表/流恢复，增量更新支持

## 快速开始

### 读取 PDF

```typescript
import { readPdf } from "@cj-tech-master/excelts/pdf";
import { readFileSync } from "fs";

const bytes = readFileSync("document.pdf");
const result = readPdf(bytes);

// 全部文本
console.log(result.text);

// 逐页文本
for (const page of result.pages) {
  console.log(`第 ${page.pageNumber} 页: ${page.text.length} 字符`);
}

// 元数据
console.log(result.metadata.title);
console.log(result.metadata.pageCount);

// 图片
for (const page of result.pages) {
  for (const img of page.images) {
    console.log(img.format, img.width, img.height);
  }
}

// 标注（链接、批注、高亮等）
for (const page of result.pages) {
  for (const annot of page.annotations) {
    console.log(annot.subtype, annot.contents, annot.uri);
  }
}

// 表单字段
for (const field of result.formFields) {
  console.log(field.name, field.type, field.value);
}
```

### 读取加密 PDF

```typescript
const result = readPdf(bytes, { password: "secret" });
```

### 选择性提取

```typescript
// 仅第 1 和第 3 页，不提取图片
const result = readPdf(bytes, {
  pages: [1, 3],
  extractImages: false
});
```

### 独立使用 — 从数组生成 PDF

```typescript
import { pdf } from "@cj-tech-master/excelts/pdf";

const bytes = pdf([
  ["产品", "收入"],
  ["小工具", 1000],
  ["大工具", 2500]
]);
```

### Excel 转 PDF

```typescript
import { Workbook, excelToPdf } from "@cj-tech-master/excelts";

const workbook = new Workbook();
await workbook.xlsx.readFile("input.xlsx");
const pdfBytes = excelToPdf(workbook, { showGridLines: true });
```

## 加密

写入器使用 **AES-256** 加密（PDF 2.0, V=5, R=5）。读取器支持**所有主流加密格式**。

| 格式    | 版本       | 支持 |
| ------- | ---------- | ---- |
| RC4-40  | V=1, R=2   | 读取 |
| RC4-128 | V=2, R=3   | 读取 |
| AES-128 | V=4, R=4   | 读取 |
| AES-256 | V=5, R=5/6 | 读写 |

## 示例

查看 [examples 目录](examples/) 获取可运行代码。

| 文件                   | 演示内容                                             |
| ---------------------- | ---------------------------------------------------- |
| `pdf-basic.ts`         | 页面大小、边距、元数据、工作表选择、缩放             |
| `pdf-styled.ts`        | 字体、填充、边框、对齐、合并、旋转、富文本、数字格式 |
| `pdf-advanced.ts`      | 分页、分页符、加密、透明度、书签、隐藏行列           |
| `pdf-excel-to-pdf.ts`  | 读取 `.xlsx` 文件并转换为 PDF                        |
| `pdf-images.ts`        | 图片嵌入（JPEG、PNG 透明度）                         |
| `pdf-reader.ts`        | 文本提取、元数据、图片、加密 PDF、选择性提取         |
| `pdf-reader-stress.ts` | 大规模压力测试：数千单元格、加密往返、基准测试       |

```bash
npx tsx src/modules/pdf/examples/pdf-reader.ts
# 输出: tmp/pdf-reader-examples/
```
