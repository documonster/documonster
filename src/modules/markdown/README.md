# Markdown Module

[中文](README_zh.md)

GFM (GitHub Flavored Markdown) table parser and formatter with zero dependencies.

```typescript
import { parseMarkdown, parseMarkdownAll, formatMarkdown } from "@cjnoname/excelts/markdown";
```

## Features

- **GFM Compliant** — Parses standard GitHub Flavored Markdown table syntax
- **Zero Dependencies** — Pure TypeScript, no external packages
- **Cross-Platform** — Same API in Node.js and browsers
- **Round-Trip** — Parse and format in one package with alignment preservation
- **Column Alignment** — Detects and generates left, center, right, none alignment
- **Pipe Escaping** — Handles `\|` and `\\` in both parse and format directions
- **CJK/Emoji Width** — Built-in display width calculation for proper column alignment
- **Multiline Cells** — `<br>` tag support for newlines within cells
- **Multi-Table** — Extract all tables from a Markdown document with `parseMarkdownAll`
- **Workbook Integration** — `Workbook.readMarkdown()` / `writeMarkdown()` for Excel↔Markdown

---

## Quick Start

### Parsing

```typescript
import { parseMarkdown } from "@cjnoname/excelts/markdown";

const result = parseMarkdown("| Name | Age |\n| --- | --- |\n| Alice | 30 |");
// result.headers = ["Name", "Age"]
// result.rows = [["Alice", "30"]]
// result.alignments = ["none", "none"]

// With alignment detection
const aligned = parseMarkdown("| Left | Center | Right |\n|:---|:---:|---:|\n|a|b|c|");
// aligned.alignments = ["left", "center", "right"]

// From a larger document (finds the first table)
const doc = parseMarkdown("# Title\n\nSome text.\n\n| A |\n| --- |\n| 1 |");
// doc.headers = ["A"], doc.rows = [["1"]]
```

### Formatting

```typescript
import { formatMarkdown } from "@cjnoname/excelts/markdown";

formatMarkdown(
  ["Name", "Age"],
  [
    ["Alice", "30"],
    ["Bob", "25"]
  ]
);
// | Name  | Age |
// | ----- | --- |
// | Alice | 30  |
// | Bob   | 25  |

// With alignment
formatMarkdown(["Left", "Center", "Right"], [["a", "b", "c"]], {
  columns: [
    { header: "Left", alignment: "left" },
    { header: "Center", alignment: "center" },
    { header: "Right", alignment: "right" }
  ]
});

// Any value types — auto-stringified
formatMarkdown(["Name", "Age", "Active"], [["Alice", 30, true]]);
```

### Workbook Integration

```typescript
import { Workbook } from "@cjnoname/excelts";

const workbook = new Workbook();

// Read Markdown → Worksheet
const ws = workbook.readMarkdown("| Name | Age |\n| --- | --- |\n| Alice | 30 |");
console.log(ws.getRow(2).getCell(1).value); // "Alice"

// Worksheet → Markdown
const markdownText = workbook.writeMarkdown();

// Read all tables from a document
const sheets = workbook.readMarkdownAll(markdownDoc, { sheetName: "Table" });
// Creates "Table", "Table_2", "Table_3", ...

// File I/O (Node.js only)
await workbook.readMarkdownFile("data.md");
await workbook.writeMarkdownFile("output.md");
```

---

## Parsing API

### `parseMarkdown(input, options?)`

Parse the first Markdown table found in the input string.

```typescript
parseMarkdown(input: string, options?: MarkdownParseOptions): MarkdownParseResult
```

Throws `MarkdownParseError` if no valid table is found.

### `parseMarkdownAll(input, options?)`

Parse all Markdown tables from a document.

```typescript
parseMarkdownAll(input: string, options?: MarkdownParseOptions): MarkdownParseResult[]
```

Returns an empty array if no tables are found.

**Parse Options (`MarkdownParseOptions`):**

| Option          | Type      | Default | Description                                  |
| --------------- | --------- | ------- | -------------------------------------------- | ------------ |
| `trim`          | `boolean` | `true`  | Trim whitespace from cell values             |
| `unescape`      | `boolean` | `true`  | Unescape `\|` → `                            | `and`\\`→`\` |
| `skipEmptyRows` | `boolean` | `true`  | Skip rows where all cells are empty          |
| `maxRows`       | `number`  | —       | Maximum data rows to parse (excludes header) |
| `convertBr`     | `boolean` | `false` | Convert `<br>` tags to newline characters    |

**Result (`MarkdownParseResult`):**

```typescript
interface MarkdownParseResult {
  headers: string[]; // Column names from header row
  rows: string[][]; // Data rows (each row = array of cell values)
  alignments: MarkdownAlignment[]; // "left" | "center" | "right" | "none"
}
```

---

## Formatting API

### `formatMarkdown(headers, rows, options?)`

Format data as a Markdown table string.

```typescript
formatMarkdown(headers: string[], rows: unknown[][], options?: MarkdownFormatOptions): string
```

**Format Options (`MarkdownFormatOptions`):**

| Option            | Type                                 | Default  | Description                               |
| ----------------- | ------------------------------------ | -------- | ----------------------------------------- | ----------------------- |
| `columns`         | `(string \| MarkdownColumnConfig)[]` | —        | Per-column header and alignment config    |
| `alignment`       | `MarkdownAlignment`                  | `"left"` | Default alignment for all columns         |
| `padding`         | `boolean`                            | `true`   | Align columns to equal width with padding |
| `trailingNewline` | `boolean`                            | `true`   | Include trailing newline in output        |
| `escapeContent`   | `boolean`                            | `true`   | Escape `                                  | `and`\` in cell content |
| `stringify`       | `(value: unknown) => string`         | built-in | Custom value-to-string converter          |

**Column Config (`MarkdownColumnConfig`):**

```typescript
interface MarkdownColumnConfig {
  header: string;
  alignment?: MarkdownAlignment; // "left" | "center" | "right" | "none"
  minWidth?: number; // Minimum column width (default: 3)
}
```

---

## Multiline Cells

Newlines in cell content are converted to `<br>` tags during formatting, and can be converted back during parsing with `convertBr: true`.

```typescript
// Format: newlines become <br>
formatMarkdown(["Note"], [["Line 1\nLine 2"]]);
// | Note           |
// | -------------- |
// | Line 1<br>Line 2 |

// Parse: <br> back to newlines
parseMarkdown(table, { convertBr: true });
// rows[0] = ["Line 1\nLine 2"]
```

---

## CJK / Unicode Width

The formatter automatically accounts for CJK characters, fullwidth forms, and emoji when calculating column widths. No external dependencies needed.

```typescript
formatMarkdown(["Name", "名前"], [["Alice", "太郎"]]);
// | Name  | 名前 |
// | ----- | ---- |
// | Alice | 太郎 |
```

---

## Errors

```typescript
import { MarkdownParseError } from "@cjnoname/excelts/markdown";

try {
  parseMarkdown("no table here");
} catch (e) {
  if (e instanceof MarkdownParseError) {
    console.log(e.message); // "Line 1: No valid Markdown table found in input"
    console.log(e.line); // 1
  }
}
```

---

## Workbook Methods

| Method                                | Platform | Description                      |
| ------------------------------------- | -------- | -------------------------------- |
| `readMarkdown(input, options?)`       | All      | Parse Markdown table → Worksheet |
| `readMarkdownAll(input, options?)`    | All      | Parse all tables → Worksheet[]   |
| `writeMarkdown(options?)`             | All      | Worksheet → Markdown string      |
| `writeMarkdownBuffer(options?)`       | All      | Worksheet → Uint8Array (UTF-8)   |
| `readMarkdownFile(path, options?)`    | Node.js  | Read from file                   |
| `readMarkdownAllFile(path, options?)` | Node.js  | Read all tables from file        |
| `writeMarkdownFile(path, options?)`   | Node.js  | Write to file                    |

**Workbook Options (`MarkdownOptions`)** extends both `MarkdownParseOptions` and `MarkdownFormatOptions`, plus:

| Option             | Type                         | Description                                            |
| ------------------ | ---------------------------- | ------------------------------------------------------ |
| `sheetName`        | `string`                     | Worksheet name (for `readMarkdownAll`: used as prefix) |
| `sheetId`          | `number`                     | Worksheet ID to write                                  |
| `map`              | `(value, column) => unknown` | Custom value mapper for parsing                        |
| `dateFormat`       | `string`                     | Date format for writing                                |
| `dateUTC`          | `boolean`                    | Use UTC for dates                                      |
| `includeEmptyRows` | `boolean`                    | Include empty rows in output                           |
