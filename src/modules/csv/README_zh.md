# CSV 模块

高性能、RFC 4180 兼容的 CSV 解析器和格式化器，零依赖。

## 功能特性

- **零依赖** — 纯 TypeScript
- **跨平台** — Node.js 和浏览器
- **流式处理** — `CsvParserStream` 和 `CsvFormatterStream`
- **类型安全泛型** — 解析结果自动推导类型
- **自动检测** — 分隔符、换行符、BOM
- **动态类型** — 自动转换数字、布尔值、日期
- **数据生成** — `csvGenerate()` 生成测试数据
- **公式转义** — CSV 注入防护
- **工作线程池** — 浏览器后台处理

## 快速开始

```typescript
import { parseCsv, formatCsv } from "@cj-tech-master/excelts/csv";

// 解析
const rows = parseCsv("姓名,年龄\nAlice,30\nBob,25");

// 解析为对象
const result = parseCsv("姓名,年龄\nAlice,30", { headers: true });
// result.rows = [{ "姓名": "Alice", "年龄": "30" }]

// 格式化
const csv = formatCsv(
  [
    { name: "Alice", age: 30 },
    { name: "Bob", age: 25 }
  ],
  { headers: true }
);

// 流式解析
import { createCsvParserStream } from "@cj-tech-master/excelts/csv";
const parser = createCsvParserStream({ headers: true });
parser.on("data", row => console.log(row));
```

## 详细文档

完整 API 参考请查看 [英文文档](README.md)。

## 示例

查看 [examples 目录](examples/) 获取可运行代码。
