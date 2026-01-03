# 快速开始

## 安装

```bash
npm install @cj-tech-master/excelts
```

## 示例

```ts
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("My Sheet");

sheet.addRow(["Name", "Age", "Email"]);
sheet.addRow(["John Doe", 30, "john@example.com"]);

await workbook.xlsx.writeFile("output.xlsx");
```

## 下一步

- 浏览器用法：见 [浏览器支持](/zh/guide/browser)
- 大文件处理：见 [流式读写](/zh/guide/streaming)
