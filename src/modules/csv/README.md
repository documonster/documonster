# CSV Module

[中文](README_zh.md)

A high-performance, RFC 4180 compliant CSV parser and formatter with zero dependencies.

```typescript
import { parseCsv, formatCsv, CsvParserStream } from "@cj-tech-master/excelts/csv";
```

## Features

- **RFC 4180 Compliant** - Full spec compliance with multi-line fields, quoted fields, and edge cases
- **Zero Dependencies** - Pure TypeScript, no external packages
- **Cross-Platform** - Same API in Node.js and browsers
- **High Performance** - indexOf-based scanner, fast-mode detection, batch processing
- **Streaming** - `CsvParserStream` / `CsvFormatterStream` for large files
- **Type-Safe** - Full TypeScript generics with overloaded signatures
- **Auto-Detection** - Delimiter, linebreak, and BOM detection
- **Dynamic Typing** - Automatic type coercion (numbers, booleans, dates)
- **Data Generation** - Built-in CSV test data generator with seeded PRNG

---

## Quick Start

### Parsing

```typescript
import { parseCsv } from "@cj-tech-master/excelts/csv";

// Simple: returns string[][]
const rows = parseCsv("name,age\nAlice,30\nBob,25");
// [["name","age"], ["Alice","30"], ["Bob","25"]]

// With headers: returns { rows, headers, meta, errors }
const result = parseCsv("name,age\nAlice,30\nBob,25", { headers: true });
// result.rows = [{ name: "Alice", age: "30" }, { name: "Bob", age: "25" }]
// result.headers = ["name", "age"]

// With dynamic typing: auto-converts numbers, booleans, dates
const typed = parseCsv("name,age,active\nAlice,30,true", {
  headers: true,
  dynamicTyping: true
});
// typed.rows = [{ name: "Alice", age: 30, active: true }]
```

### Formatting

```typescript
import { formatCsv } from "@cj-tech-master/excelts/csv";

// From arrays
formatCsv([
  ["name", "age"],
  ["Alice", "30"]
]);
// "name,age\nAlice,30"

// From objects (auto-derives headers)
formatCsv([
  { name: "Alice", age: 30 },
  { name: "Bob", age: 25 }
]);
// "name,age\nAlice,30\nBob,25"

// With options
formatCsv(data, {
  delimiter: ";",
  bom: true,
  escapeFormulae: true,
  quoteColumns: { name: true }
});
```

---

## Parsing API

### `parseCsv(input, options?)`

Synchronous CSV parser with type-safe overloads.

```typescript
// No options -> string[][]
parseCsv(csvString): string[][];

// headers: true -> CsvParseResult<Record<string, unknown>>
parseCsv(csvString, { headers: true }): CsvParseResult;

// headers: true + dynamicTyping -> auto-coerced values
parseCsv(csvString, { headers: true, dynamicTyping: true }): CsvParseResult;

// info: true -> rows include metadata (line number, raw input, etc.)
parseCsv(csvString, { headers: true, info: true }): CsvParseResult<RecordWithInfo>;
```

**Parse Options (`CsvParseOptions`):**

| Option           | Type                                             | Default       | Description                                             |
| ---------------- | ------------------------------------------------ | ------------- | ------------------------------------------------------- |
| `delimiter`      | `string`                                         | auto-detect   | Field delimiter (`,`, `;`, `\t`, `\|`)                  |
| `headers`        | `boolean \| string[] \| HeaderTransformFunction` | `false`       | Use first row as headers, or provide custom             |
| `quote`          | `string \| false`                                | `"`           | Quote character (`false` to disable)                    |
| `escape`         | `string \| false`                                | same as quote | Escape character inside quoted fields                   |
| `comment`        | `string`                                         | -             | Comment line prefix (e.g. `"#"`)                        |
| `skipEmptyLines` | `boolean \| "greedy"`                            | `false`       | Skip empty lines; `"greedy"` also skips whitespace-only |
| `skipLines`      | `number`                                         | `0`           | Number of lines to skip before parsing                  |
| `maxRows`        | `number`                                         | -             | Maximum rows to parse                                   |
| `dynamicTyping`  | `boolean \| DynamicTypingConfig`                 | `false`       | Auto-convert to numbers/booleans/dates                  |
| `castDate`       | `boolean \| CastDateConfig`                      | `false`       | Parse date strings to Date objects                      |
| `transform`      | `RowTransformFunction`                           | -             | Transform each row before output                        |
| `validate`       | `RowValidateFunction`                            | -             | Validate rows; invalid rows go to `errors`              |
| `columns`        | `ColumnConfig[]`                                 | -             | Per-column type/transform configuration                 |
| `objname`        | `string`                                         | -             | Key rows by a specific column value                     |
| `encoding`       | `string`                                         | `"utf-8"`     | Input encoding                                          |
| `info`           | `boolean`                                        | `false`       | Include line/record metadata per row                    |
| `fastMode`       | `boolean`                                        | auto          | Force fast-mode parsing (no quote handling)             |
| `renameHeaders`  | `boolean`                                        | `false`       | Auto-rename duplicate headers                           |
| `columnMismatch` | `ColumnMismatchConfig`                           | -             | Handle rows with more/fewer columns than headers        |

**Result (`CsvParseResult<T>`):**

```typescript
interface CsvParseResult<T> {
  rows: T[];
  headers: string[];
  meta: CsvParseMeta; // { delimiter, linebreak, rowCount, truncated }
  errors: CsvRecordError[];
  invalidRows: T[];
}
```

### `parseCsvAsync(input, options?)`

Async parser supporting strings, `AsyncIterable<string | Uint8Array>`, and `ReadableStream`.

```typescript
import { parseCsvAsync } from "@cj-tech-master/excelts/csv";

// From string
const result = await parseCsvAsync(csvString, { headers: true });

// From ReadableStream (browser fetch)
const response = await fetch("/data.csv");
const result = await parseCsvAsync(response.body, { headers: true });

// From async iterable
const result = await parseCsvAsync(asyncChunks, { headers: true });
```

### `parseCsvRows(input, options?)`

True streaming async generator -- yields rows one at a time. Memory-efficient for large files.

```typescript
import { parseCsvRows } from "@cj-tech-master/excelts/csv";

for await (const row of parseCsvRows(hugeFile, { headers: true })) {
  console.log(row); // { name: "...", age: "..." }
}
```

### `parseCsvWithProgress(input, options?, onProgress?)`

Async parser with progress callback for large files.

```typescript
import { parseCsvWithProgress } from "@cj-tech-master/excelts/csv";

const result = await parseCsvWithProgress(
  largeCsvString,
  { headers: true },
  ({ rowsProcessed, bytesProcessed }) => {
    console.log(`Parsed ${rowsProcessed} rows (${bytesProcessed} bytes)`);
  }
);
```

---

## Formatting API

### `formatCsv(data, options?)`

Batch CSV formatter. Accepts arrays of arrays or arrays of objects.

```typescript
import { formatCsv } from "@cj-tech-master/excelts/csv";

// Array of arrays
formatCsv([
  ["a", "b"],
  [1, 2]
]);

// Array of objects
formatCsv([{ name: "Alice", age: 30 }]);

// RowHashArray format
formatCsv([
  [
    ["name", "Alice"],
    ["age", "30"]
  ]
]);
```

**Format Options (`CsvFormatOptions`):**

| Option            | Type                                              | Default       | Description                                        |
| ----------------- | ------------------------------------------------- | ------------- | -------------------------------------------------- |
| `delimiter`       | `string`                                          | `,`           | Field delimiter                                    |
| `headers`         | `boolean \| string[]`                             | auto          | Include headers; auto-derived from objects         |
| `quote`           | `string \| false`                                 | `"`           | Quote character                                    |
| `escape`          | `string`                                          | same as quote | Escape character                                   |
| `quoteColumns`    | `boolean \| boolean[] \| Record<string, boolean>` | auto          | Force quoting per column                           |
| `escapeFormulae`  | `boolean`                                         | `false`       | Prefix formula chars (`=`, `+`, `-`, `@`) with `'` |
| `bom`             | `boolean`                                         | `false`       | Prepend UTF-8 BOM                                  |
| `columns`         | `ColumnConfig[]`                                  | -             | Per-column formatting configuration                |
| `transform`       | `RowTransformFunction`                            | -             | Transform rows before formatting                   |
| `trailingNewline` | `boolean`                                         | `false`       | Add trailing newline                               |

---

## Streaming API

### `CsvParserStream`

Transform stream that parses CSV data chunk-by-chunk. Cross-platform (Node.js + browser).

```typescript
import { CsvParserStream, createCsvParserStream } from "@cj-tech-master/excelts/csv";
import { pipeline } from "@cj-tech-master/excelts/stream";

// Using factory function
const parser = createCsvParserStream({ headers: true, dynamicTyping: true });

// Using class constructor
const parser = new CsvParserStream({ headers: true });

// With transform and validation
parser.transform(row => ({ ...row, age: Number(row.age) })).validate(row => row.age > 0);

// Events
parser.on("headers", headers => console.log("Headers:", headers));
parser.on("data", row => console.log("Row:", row));
parser.on("data-invalid", row => console.log("Invalid:", row));
parser.on("end", () => console.log("Done"));

// Pipeline usage
await pipeline(readableStream, parser, writable);
```

### `CsvFormatterStream`

Transform stream that formats rows to CSV text. Cross-platform (Node.js + browser).

```typescript
import { CsvFormatterStream, createCsvFormatterStream } from "@cj-tech-master/excelts/csv";

const formatter = createCsvFormatterStream({
  headers: ["name", "age"],
  delimiter: ";",
  bom: true
});

formatter.write(["Alice", 30]);
formatter.write(["Bob", 25]);
formatter.end();

// Pipe output to a file or writable stream
formatter.pipe(writable);
```

---

## Detection Utilities

```typescript
import { detectDelimiter, detectLinebreak, stripBom } from "@cj-tech-master/excelts/csv";

// Auto-detect delimiter from CSV content
detectDelimiter("a,b,c\n1,2,3"); // ","
detectDelimiter("a;b;c\n1;2;3"); // ";"
detectDelimiter("a\tb\tc\n1\t2\t3"); // "\t"

// Detect line terminator (quote-aware)
detectLinebreak("a,b\r\nc,d"); // "\r\n"
detectLinebreak("a,b\nc,d"); // "\n"

// Strip UTF-8 BOM
stripBom("\ufeffname,age"); // "name,age"
```

---

## Row Utilities

```typescript
import { isRowHashArray, deduplicateHeaders, processColumns } from "@cj-tech-master/excelts/csv";

// Check if row is a RowHashArray
isRowHashArray([["key", "value"]]); // true

// Deduplicate header names
deduplicateHeaders(["id", "name", "name", "name"]);
// ["id", "name", "name_1", "name_2"]
```

---

## Dynamic Typing

```typescript
import { applyDynamicTyping } from "@cj-tech-master/excelts/csv";

// Auto-convert string values to native types
applyDynamicTyping("42"); // 42 (number)
applyDynamicTyping("3.14"); // 3.14 (number)
applyDynamicTyping("true"); // true (boolean)
applyDynamicTyping("hello"); // "hello" (string, unchanged)
```

---

## Number Utilities

```typescript
import { formatNumberForCsv, parseNumberFromCsv } from "@cj-tech-master/excelts/csv";

// Format numbers with locale-specific decimal separator
formatNumberForCsv(3.14, "."); // "3.14"
formatNumberForCsv(3.14, ","); // "3,14"

// Parse numbers with locale-specific decimal separator
parseNumberFromCsv("3,14", ","); // 3.14
```

---

## CSV Data Generator

Generate test CSV data with built-in column types and seeded PRNG for reproducibility.

```typescript
import {
  csvGenerate,
  csvGenerateRows,
  csvGenerateAsync,
  csvGenerateData,
  createCsvGenerator
} from "@cj-tech-master/excelts/csv";

// Generate CSV string
const { csv, headers, data } = csvGenerate({
  columns: ["name", "email", "int", "bool", "date"],
  rows: 100,
  seed: 42
});

// Generate with custom column types
const { csv } = csvGenerate({
  columns: [
    { type: "int", min: 18, max: 65, name: "age" },
    { type: "float", min: 0, max: 100, name: "score" },
    ctx => `row-${ctx.rowIndex}`
  ],
  rows: 50
});

// Memory-efficient: yields rows one at a time
for (const row of csvGenerateRows({ columns: 5, rows: 1_000_000 })) {
  process.stdout.write(row + "\n");
}

// Generate for a duration (unlimited rows)
for (const row of csvGenerateRows({ columns: 3, duration: 5000 })) {
  // Generates for 5 seconds
}

// Async generator with delay between rows
for await (const row of csvGenerateAsync({ columns: 5, rows: 100, delay: 10 })) {
  console.log(row);
}

// Generate raw data (not CSV strings)
const rawRows = csvGenerateData({ columns: ["name", "int"], rows: 10 });
// [[name, number], ...]

const objects = csvGenerateData({
  columns: [{ type: "name", name: "fullName" }],
  rows: 10,
  objectMode: true
});
// [{ fullName: "..." }, ...]

// Reusable generator with preset config
const gen = createCsvGenerator({ columns: ["name", "email"], seed: 42 });
const batch1 = gen.generate(100);
const batch2 = gen.generate(100);
```

**Built-in Column Types:**
`string`, `int`, `float`, `bool`, `date`, `uuid`, `email`, `name`, `url`, `phone`, `address`, `ip`, `hex`, `lorem`

---

## Error Classes

```typescript
import { CsvError, CsvWorkerError } from "@cj-tech-master/excelts/csv";

try {
  parseCsv(badInput, { headers: true });
} catch (e) {
  if (e instanceof CsvError) {
    console.error(e.message);
    console.error(e.cause); // Original error if any
  }
}
```

---

## API Reference

### Core Functions

| Function                                             | Description                                      |
| ---------------------------------------------------- | ------------------------------------------------ |
| `parseCsv(input, options?)`                          | Synchronous CSV parser                           |
| `parseCsvAsync(input, options?)`                     | Async parser (strings, streams, async iterables) |
| `parseCsvRows(input, options?)`                      | Async generator yielding rows one at a time      |
| `parseCsvWithProgress(input, options?, onProgress?)` | Parser with progress reporting                   |
| `formatCsv(data, options?)`                          | Batch CSV formatter                              |

### Stream Classes

| Class                                | Description                                |
| ------------------------------------ | ------------------------------------------ |
| `CsvParserStream`                    | Transform stream: CSV bytes -> parsed rows |
| `CsvFormatterStream`                 | Transform stream: rows -> CSV text         |
| `createCsvParserStream(options?)`    | Factory for `CsvParserStream`              |
| `createCsvFormatterStream(options?)` | Factory for `CsvFormatterStream`           |

### Utilities

| Function                         | Description                   |
| -------------------------------- | ----------------------------- |
| `detectDelimiter(input)`         | Auto-detect CSV delimiter     |
| `detectLinebreak(input)`         | Auto-detect line terminator   |
| `stripBom(input)`                | Strip UTF-8 BOM               |
| `applyDynamicTyping(value)`      | Convert string to native type |
| `formatNumberForCsv(value, sep)` | Format number for locale      |
| `parseNumberFromCsv(value, sep)` | Parse locale-formatted number |
| `deduplicateHeaders(headers)`    | Rename duplicate headers      |

### Generator Functions

| Function                       | Description                         |
| ------------------------------ | ----------------------------------- |
| `csvGenerate(options?)`        | Generate CSV string + data          |
| `csvGenerateRows(options?)`    | Sync generator yielding CSV rows    |
| `csvGenerateAsync(options?)`   | Async generator with delay support  |
| `csvGenerateData(options?)`    | Generate raw data (not CSV strings) |
| `createCsvGenerator(options?)` | Reusable generator factory          |
