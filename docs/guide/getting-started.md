# Getting Started

## Installation

```bash
npm install @cj-tech-master/excelts
```

## Quick Start

```ts
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("My Sheet");

sheet.addRow(["Name", "Age", "Email"]);
sheet.addRow(["John Doe", 30, "john@example.com"]);

await workbook.xlsx.writeFile("output.xlsx");
```

## Next steps

- Browser usage: see [Browser Support](/guide/browser)
- Large files: see [Streaming](/guide/streaming)
