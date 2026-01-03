# Streaming

ExcelTS provides streaming reader and writer APIs for processing large Excel files.

## Streaming Reader

```ts
import { WorkbookReader } from "@cj-tech-master/excelts";

const reader = new WorkbookReader("large-file.xlsx", {
  worksheets: "emit",
  sharedStrings: "cache",
  hyperlinks: "ignore",
  styles: "ignore"
});

for await (const worksheet of reader) {
  for await (const row of worksheet) {
    console.log(row.values);
  }
}
```

## Streaming Writer

```ts
import { WorkbookWriter } from "@cj-tech-master/excelts";

const workbook = new WorkbookWriter({
  filename: "output.xlsx",
  useSharedStrings: true,
  useStyles: true
});

const sheet = workbook.addWorksheet("Data");

sheet.addRow(["Name", "Score"]).commit();

await sheet.commit();
await workbook.commit();
```
