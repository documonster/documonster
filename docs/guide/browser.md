# Browser Support

ExcelTS has native browser support with zero configuration required for modern bundlers.

## Bundlers (Vite/Webpack/Rollup/esbuild)

```ts
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("Sheet1");

sheet.getCell("A1").value = "Hello, Browser!";

const buffer = await workbook.xlsx.writeBuffer();
const blob = new Blob([buffer], {
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
});

// Use URL.createObjectURL(blob) and trigger a download in your app.
```

## Notes

- Use `xlsx.load(arrayBuffer)` instead of `xlsx.readFile()`.
- Use `xlsx.writeBuffer()` instead of `xlsx.writeFile()`.
