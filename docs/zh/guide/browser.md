# 浏览器支持

ExcelTS 支持在浏览器中使用，并对现代打包器提供零配置支持。

## 打包器（Vite / Webpack / Rollup / esbuild）

```ts
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("Sheet1");

sheet.getCell("A1").value = "Hello, Browser!";

const buffer = await workbook.xlsx.writeBuffer();
const blob = new Blob([buffer], {
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
});

// 结合 URL.createObjectURL(blob) 触发下载。
```

## 注意

- 浏览器中使用 `xlsx.load(arrayBuffer)` 替代 `xlsx.readFile()`。
- 浏览器中使用 `xlsx.writeBuffer()` 替代 `xlsx.writeFile()`。
