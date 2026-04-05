# Markdown 模块

[English](README.md)

GFM（GitHub 风格 Markdown）表格解析器和格式化器，零依赖。

## 功能特性

- **GFM 兼容** — 完整的 GitHub 风格 Markdown 表格语法
- **零依赖** — 纯 TypeScript
- **往返保持** — 对齐方式保留
- **管道符转义** — 自动处理内容中的 `|` 字符
- **CJK/Emoji 宽度** — 正确计算东亚字符和 Emoji 的显示宽度
- **多行单元格** — 通过 `<br>` 支持
- **多表格提取** — `parseMarkdownAll()` 从文档中提取所有表格
- **工作簿集成** — `readMarkdown()`、`writeMarkdown()` 直接与 Excel 工作簿交互

## 快速开始

```typescript
import { parseMarkdown, formatMarkdown, parseMarkdownAll } from "@cj-tech-master/excelts/markdown";

// 解析 Markdown 表格
const result = parseMarkdown("| 姓名 | 年龄 |\n| --- | --- |\n| Alice | 30 |");
// result.headers = ["姓名", "年龄"]
// result.rows = [["Alice", "30"]]
// result.alignments = ["none", "none"]

// 解析文档中的所有表格
const tables = parseMarkdownAll(markdownDoc);

// 格式化为 Markdown 表格
const markdownText = formatMarkdown(
  ["姓名", "年龄"],
  [
    ["Alice", "30"],
    ["Bob", "25"]
  ],
  { alignment: "left", padding: true }
);
```

## 示例

查看 [examples 目录](examples/) 获取可运行代码。
