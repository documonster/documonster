# Markdown 模块

[English](README.md)

GFM（GitHub 风格 Markdown）表格解析器和格式化器，零依赖。

```typescript
import { parseMarkdown, parseMarkdownAll, formatMarkdown } from "@cjnoname/excelts/markdown";
```

## 功能特性

- **GFM 兼容** — 解析标准 GitHub 风格 Markdown 表格语法
- **零依赖** — 纯 TypeScript，无外部包
- **跨平台** — Node.js 和浏览器使用相同 API
- **往返保持** — 解析和格式化一体，保留对齐方式
- **列对齐** — 检测和生成左对齐、居中、右对齐、无对齐
- **管道符转义** — 在解析和格式化方向上处理 `\|` 和 `\\`
- **CJK/Emoji 宽度** — 内置显示宽度计算，正确对齐列
- **多行单元格** — 通过 `<br>` 标签支持单元格内换行
- **多表格提取** — 使用 `parseMarkdownAll` 从 Markdown 文档中提取所有表格
- **工作簿集成** — `Workbook.readMarkdown()` / `writeMarkdown()` 实现 Excel↔Markdown 转换

---

## 快速开始

### 解析

```typescript
import { parseMarkdown } from "@cjnoname/excelts/markdown";

const result = parseMarkdown("| Name | Age |\n| --- | --- |\n| Alice | 30 |");
// result.headers = ["Name", "Age"]
// result.rows = [["Alice", "30"]]
// result.alignments = ["none", "none"]

// 对齐检测
const aligned = parseMarkdown("| Left | Center | Right |\n|:---|:---:|---:|\n|a|b|c|");
// aligned.alignments = ["left", "center", "right"]

// 从较大的文档中提取（查找第一个表格）
const doc = parseMarkdown("# Title\n\nSome text.\n\n| A |\n| --- |\n| 1 |");
// doc.headers = ["A"], doc.rows = [["1"]]
```

### 格式化

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

// 使用对齐
formatMarkdown(["Left", "Center", "Right"], [["a", "b", "c"]], {
  columns: [
    { header: "Left", alignment: "left" },
    { header: "Center", alignment: "center" },
    { header: "Right", alignment: "right" }
  ]
});

// 任意值类型 — 自动字符串化
formatMarkdown(["Name", "Age", "Active"], [["Alice", 30, true]]);
```

### 工作簿集成

```typescript
import { Workbook } from "@cjnoname/excelts";

const workbook = new Workbook();

// 读取 Markdown → 工作表
const ws = workbook.readMarkdown("| Name | Age |\n| --- | --- |\n| Alice | 30 |");
console.log(ws.getRow(2).getCell(1).value); // "Alice"

// 工作表 → Markdown
const markdownText = workbook.writeMarkdown();

// 从文档中读取所有表格
const sheets = workbook.readMarkdownAll(markdownDoc, { sheetName: "Table" });
// 创建 "Table"、"Table_2"、"Table_3"...

// 文件 I/O（仅 Node.js）
await workbook.readMarkdownFile("data.md");
await workbook.writeMarkdownFile("output.md");
```

---

## 解析 API

### `parseMarkdown(input, options?)`

解析输入字符串中找到的第一个 Markdown 表格。

```typescript
parseMarkdown(input: string, options?: MarkdownParseOptions): MarkdownParseResult
```

如果未找到有效表格，抛出 `MarkdownParseError`。

### `parseMarkdownAll(input, options?)`

解析文档中的所有 Markdown 表格。

```typescript
parseMarkdownAll(input: string, options?: MarkdownParseOptions): MarkdownParseResult[]
```

如果未找到表格，返回空数组。

**解析选项（`MarkdownParseOptions`）：**

| 选项            | 类型      | 默认值  | 描述                       |
| --------------- | --------- | ------- | -------------------------- | ----------- |
| `trim`          | `boolean` | `true`  | 去除单元格值的空白         |
| `unescape`      | `boolean` | `true`  | 反转义 `\|` → `            | `和`\\`→`\` |
| `skipEmptyRows` | `boolean` | `true`  | 跳过所有单元格为空的行     |
| `maxRows`       | `number`  | —       | 最大数据行数（不含表头）   |
| `convertBr`     | `boolean` | `false` | 将 `<br>` 标签转换为换行符 |

**结果（`MarkdownParseResult`）：**

```typescript
interface MarkdownParseResult {
  headers: string[]; // 表头行的列名
  rows: string[][]; // 数据行（每行 = 单元格值数组）
  alignments: MarkdownAlignment[]; // "left" | "center" | "right" | "none"
}
```

---

## 格式化 API

### `formatMarkdown(headers, rows, options?)`

将数据格式化为 Markdown 表格字符串。

```typescript
formatMarkdown(headers: string[], rows: unknown[][], options?: MarkdownFormatOptions): string
```

**格式化选项（`MarkdownFormatOptions`）：**

| 选项              | 类型                                 | 默认值   | 描述                     |
| ----------------- | ------------------------------------ | -------- | ------------------------ | ------ |
| `columns`         | `(string \| MarkdownColumnConfig)[]` | —        | 按列的表头和对齐配置     |
| `alignment`       | `MarkdownAlignment`                  | `"left"` | 所有列的默认对齐方式     |
| `padding`         | `boolean`                            | `true`   | 对齐列到等宽并填充       |
| `trailingNewline` | `boolean`                            | `true`   | 输出中包含尾部换行符     |
| `escapeContent`   | `boolean`                            | `true`   | 转义单元格内容中的 `     | `和`\` |
| `stringify`       | `(value: unknown) => string`         | 内置     | 自定义值到字符串的转换器 |

**列配置（`MarkdownColumnConfig`）：**

```typescript
interface MarkdownColumnConfig {
  header: string;
  alignment?: MarkdownAlignment; // "left" | "center" | "right" | "none"
  minWidth?: number; // 最小列宽（默认：3）
}
```

---

## 多行单元格

格式化时，单元格内容中的换行符会转换为 `<br>` 标签，解析时使用 `convertBr: true` 可以转换回来。

```typescript
// 格式化：换行变为 <br>
formatMarkdown(["Note"], [["Line 1\nLine 2"]]);
// | Note           |
// | -------------- |
// | Line 1<br>Line 2 |

// 解析：<br> 转回换行
parseMarkdown(table, { convertBr: true });
// rows[0] = ["Line 1\nLine 2"]
```

---

## CJK / Unicode 宽度

格式化器在计算列宽时自动考虑 CJK 字符、全角形式和 emoji。无需外部依赖。

```typescript
formatMarkdown(["Name", "名前"], [["Alice", "太郎"]]);
// | Name  | 名前 |
// | ----- | ---- |
// | Alice | 太郎 |
```

---

## 错误

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

## 工作簿方法

| 方法                                  | 平台    | 描述                        |
| ------------------------------------- | ------- | --------------------------- |
| `readMarkdown(input, options?)`       | 全平台  | 解析 Markdown 表格 → 工作表 |
| `readMarkdownAll(input, options?)`    | 全平台  | 解析所有表格 → 工作表[]     |
| `writeMarkdown(options?)`             | 全平台  | 工作表 → Markdown 字符串    |
| `writeMarkdownBuffer(options?)`       | 全平台  | 工作表 → Uint8Array (UTF-8) |
| `readMarkdownFile(path, options?)`    | Node.js | 从文件读取                  |
| `readMarkdownAllFile(path, options?)` | Node.js | 从文件读取所有表格          |
| `writeMarkdownFile(path, options?)`   | Node.js | 写入到文件                  |

**工作簿选项（`MarkdownOptions`）** 继承 `MarkdownParseOptions` 和 `MarkdownFormatOptions`，另加：

| 选项               | 类型                         | 描述                                       |
| ------------------ | ---------------------------- | ------------------------------------------ |
| `sheetName`        | `string`                     | 工作表名称（`readMarkdownAll` 时用作前缀） |
| `sheetId`          | `number`                     | 要写入的工作表 ID                          |
| `map`              | `(value, column) => unknown` | 解析时的自定义值映射器                     |
| `dateFormat`       | `string`                     | 写入时的日期格式                           |
| `dateUTC`          | `boolean`                    | 日期使用 UTC                               |
| `includeEmptyRows` | `boolean`                    | 输出中包含空行                             |
