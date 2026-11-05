# CSV Module

[中文](README_zh.md)

A high-performance, RFC 4180 compliant CSV parser and formatter with zero dependencies.

```typescript
import { Csv } from "documonster/csv";
// Csv.parse, Csv.format, Csv.ParserStream, …
```

## Features

- **RFC 4180 Compliant** - Full spec compliance with multi-line fields, quoted fields, and edge cases
- **Zero Dependencies** - Pure TypeScript, no external packages
- **Cross-Platform** - Same API in Node.js and browsers
- **High Performance** - indexOf-based scanner, fast-mode detection, batch processing
- **Streaming** - `Csv.ParserStream` / `Csv.FormatterStream` for large files
- **Type-Safe** - Full TypeScript generics with overloaded signatures
- **Auto-Detection** - Delimiter, linebreak, and BOM detection
- **Dynamic Typing** - Automatic type coercion (numbers, booleans, dates)
- **Data Generation** - Built-in CSV test data generator with seeded PRNG

---

## Quick Start

### Parsing

```typescript
import { Csv } from "documonster/csv";

// Simple: returns string[][]
const rows = Csv.parse("name,age\nAlice,30\nBob,25");
// [["name","age"], ["Alice","30"], ["Bob","25"]]

// With headers: returns { rows, headers, meta, errors }
const result = Csv.parse("name,age\nAlice,30\nBob,25", { headers: true });
// result.rows = [{ name: "Alice", age: "30" }, { name: "Bob", age: "25" }]
// result.headers = ["name", "age"]

// With dynamic typing: auto-converts numbers, booleans, dates
const typed = Csv.parse("name,age,active\nAlice,30,true", {
  headers: true,
  dynamicTyping: true
});
// typed.rows = [{ name: "Alice", age: 30, active: true }]
```

### Formatting

```typescript
import { Csv } from "documonster/csv";

// From arrays
Csv.format([
  ["name", "age"],
  ["Alice", "30"]
]);
// "name,age\nAlice,30"

// From objects (auto-derives headers)
Csv.format([
  { name: "Alice", age: 30 },
  { name: "Bob", age: 25 }
]);
// "name,age\nAlice,30\nBob,25"

// With options
Csv.format(data, {
  delimiter: ";",
  bom: true,
  escapeFormulae: true,
  quoteColumns: { name: true }
});
```

---

## Parsing API

### `Csv.parse(input, options?)`

Synchronous CSV parser with type-safe overloads.

```text
// No options -> string[][]
Csv.parse(csvString): string[][];

// headers: true -> CsvParseResult<Record<string, unknown>>
Csv.parse(csvString, { headers: true }): CsvParseResult;

// headers: true + dynamicTyping -> auto-coerced values
Csv.parse(csvString, { headers: true, dynamicTyping: true }): CsvParseResult;

// info: true -> rows include metadata (line number, raw input, etc.)
Csv.parse(csvString, { headers: true, info: true }): CsvParseResult<CsvRecordWithInfo>;
```

**Parse Options (`CsvParseOptions`):**

| Option             | Type                                                | Default       | Description                                               |
| ------------------ | --------------------------------------------------- | ------------- | --------------------------------------------------------- |
| `delimiter`        | `string`                                            | `,`           | Field delimiter; `""` auto-detects (`,`, `;`, `\t`, `\|`) |
| `headers`          | `boolean \| string[] \| CsvHeaderTransformFunction` | `false`       | Use first row as headers, or provide custom               |
| `quote`            | `string \| false`                                   | `"`           | Quote character (`false` to disable)                      |
| `escape`           | `string \| false`                                   | same as quote | Escape character inside quoted fields                     |
| `comment`          | `string`                                            | -             | Comment line prefix (e.g. `"#"`)                          |
| `skipEmptyLines`   | `boolean \| "greedy"`                               | `false`       | Skip empty lines; `"greedy"` also skips whitespace-only   |
| `skipLines`        | `number`                                            | `0`           | Number of lines to skip before parsing                    |
| `maxRows`          | `number`                                            | -             | Maximum rows to parse                                     |
| `dynamicTyping`    | `boolean \| CsvDynamicTypingConfig`                 | `false`       | Auto-convert to numbers/booleans/dates                    |
| `decimalSeparator` | `"." \| ","`                                        | `"."`         | Decimal separator `dynamicTyping` reads numbers with      |
| `castDate`         | `boolean \| CsvCastDateConfig`                      | `false`       | Parse ISO dates; a date-only value is UTC midnight        |
| `rowTransform`     | `(row) => row \| null`                              | -             | Transform each row before output                          |
| `validate`         | `(row) => boolean \| { isValid, reason }`           | -             | Validate rows; rejected rows go to `invalidRows`          |
| `objname`          | `string`                                            | -             | Key rows by a header column (needs `headers`)             |
| `onSkip`           | `CsvOnSkipCallback`                                 | -             | Called for every record left out because of an error      |
| `encoding`         | `string`                                            | `"utf-8"`     | Input encoding                                            |
| `info`             | `boolean`                                           | `false`       | Include line/record metadata per row                      |
| `fastMode`         | `boolean`                                           | `false`       | Fast-mode parsing (no quote handling)                     |
| `columnMismatch`   | `CsvColumnMismatchConfig`                           | -             | Handle rows with more/fewer columns than headers          |

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

### `Csv.parseAsync(input, options?)`

Async parser supporting strings, bytes (`Uint8Array` or any `ArrayBufferView`, `ArrayBuffer`),
a `Blob` / `File`, `AsyncIterable<string | Uint8Array>`, and `ReadableStream`. Bytes are decoded as
the same bytes streamed would be (`encoding`, BOM). `Csv.parseRows` takes the same inputs.

```typescript
import { Csv } from "documonster/csv";

// From string
const result = await Csv.parseAsync(csvString, { headers: true });

// From ReadableStream (browser fetch)
const response = await fetch("/data.csv");
const result = await Csv.parseAsync(response.body, { headers: true });

// From async iterable
const result = await Csv.parseAsync(asyncChunks, { headers: true });

// From bytes or a File (an <input type="file">)
const result = await Csv.parseAsync(fileInput.files[0], { headers: true });
```

### `Csv.parseRows(input, options?)`

True streaming async generator -- yields rows one at a time. Memory-efficient for large files.

```typescript
import { Csv } from "documonster/csv";

for await (const row of Csv.parseRows(hugeFile, { headers: true })) {
  console.log(row); // { name: "...", age: "..." }
}
```

### `Csv.parseAsync(input, { onProgress })`

Async parser with progress callback for large files.

```typescript
import { Csv } from "documonster/csv";

const result = await Csv.parseAsync(largeCsvString, {
  headers: true,
  onProgress: ({ rowsProcessed, bytesProcessed }) => {
    console.log(`Parsed ${rowsProcessed} rows (${bytesProcessed} bytes)`);
  }
});
```

---

## Formatting API

### `Csv.format(data, options?)`

Batch CSV formatter. Accepts arrays of arrays or arrays of objects.

```typescript
import { Csv } from "documonster/csv";

// Array of arrays
Csv.format([
  ["a", "b"],
  [1, 2]
]);

// Array of objects
Csv.format([{ name: "Alice", age: 30 }]);

// RowHashArray format
Csv.format([
  [
    ["name", "Alice"],
    ["age", "30"]
  ]
]);
```

**Format Options (`CsvFormatOptions`):**

| Option             | Type                                              | Default       | Description                                        |
| ------------------ | ------------------------------------------------- | ------------- | -------------------------------------------------- |
| `delimiter`        | `string`                                          | `,`           | Field delimiter                                    |
| `headers`          | `boolean \| string[]`                             | auto          | Header row; object rows default to their keys      |
| `quote`            | `string \| false`                                 | `"`           | Quote character                                    |
| `escape`           | `string`                                          | same as quote | Escape character                                   |
| `quoteColumns`     | `boolean \| boolean[] \| Record<string, boolean>` | auto          | Force quoting per column                           |
| `escapeFormulae`   | `boolean`                                         | `false`       | Prefix formula chars (`=`, `+`, `-`, `@`) with `'` |
| `bom`              | `boolean`                                         | `false`       | Prepend UTF-8 BOM                                  |
| `columns`          | `(string \| CsvColumnConfig)[]`                   | -             | Keys to read, with optional header text            |
| `typeTransform`    | `CsvTypeTransformMap`                             | -             | Per-type value formatting, and a `row` transform   |
| `decimalSeparator` | `"." \| ","`                                      | `"."`         | Decimal separator for numbers                      |
| `trailingNewline`  | `boolean`                                         | `false`       | Add trailing newline                               |

---

## Streaming API

### `new Csv.ParserStream(options?)`

Transform stream that parses CSV data chunk-by-chunk. Cross-platform (Node.js + browser).

```typescript
import { Csv } from "documonster/csv";
import { pipeline } from "documonster/stream";

// Construct the stream
const parser = new Csv.ParserStream({ headers: true, dynamicTyping: true });

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

### `new Csv.FormatterStream(options?)`

Transform stream that formats rows to CSV text. Cross-platform (Node.js + browser).

```typescript
import { Csv } from "documonster/csv";

const formatter = new Csv.FormatterStream({
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
import { Csv } from "documonster/csv";

// Auto-detect delimiter from CSV content
Csv.detectDelimiter("a,b,c\n1,2,3"); // ","
Csv.detectDelimiter("a;b;c\n1;2;3"); // ";"
Csv.detectDelimiter("a\tb\tc\n1\t2\t3"); // "\t"
Csv.detectDelimiter("a:b\n1:2", { delimitersToGuess: [":"] }); // ":"

// Detect line terminator (quote-aware)
Csv.detectLinebreak("a,b\r\nc,d"); // "\r\n"
Csv.detectLinebreak("a,b\nc,d"); // "\n"

// Strip UTF-8 BOM
Csv.stripBom("\ufeffname,age"); // "name,age"
```

---

## Row Utilities

```typescript
import { Csv } from "documonster/csv";

// Check if row is a RowHashArray
Csv.isRowHashArray([["key", "value"]]); // true

// Deduplicate header names
Csv.deduplicateHeaders(["id", "name", "name", "name"]);
// ["id", "name", "name_1", "name_2"]
```

---

## Dynamic Typing

```typescript
import { Csv } from "documonster/csv";

// Auto-convert string values to native types
Csv.applyDynamicTyping("42"); // 42 (number)
Csv.applyDynamicTyping("3.14"); // 3.14 (number)
Csv.applyDynamicTyping("true"); // true (boolean)
Csv.applyDynamicTyping("hello"); // "hello" (string, unchanged)
Csv.applyDynamicTyping("007", value => Number(value)); // 7 (your own converter)
```

---

## Number Utilities

```typescript
import { Csv } from "documonster/csv";

// Format numbers with locale-specific decimal separator
Csv.formatNumber(3.14); // "3.14"
Csv.formatNumber(3.14, { decimalSeparator: "," }); // "3,14"

// Parse numbers with locale-specific decimal separator
Csv.parseNumber("3,14", { decimalSeparator: "," }); // 3.14
```

---

## CSV Data Generator

Generate test CSV data with built-in column types and seeded PRNG for reproducibility.

```typescript
import { Csv } from "documonster/csv";

// Generate CSV string
const { csv, headers, data } = Csv.generate({
  columns: ["name", "email", "int", "bool", "date"],
  rows: 100,
  seed: 42
});

// Generate with custom column types
const { csv } = Csv.generate({
  columns: [
    { type: "int", min: 18, max: 65, name: "age" },
    { type: "float", min: 0, max: 100, name: "score" },
    ctx => `row-${ctx.rowIndex}`
  ],
  rows: 50
});

// Memory-efficient: yields rows one at a time
for (const row of Csv.generateRows({ columns: 5, rows: 1_000_000 })) {
  process.stdout.write(row + "\n");
}

// Generate for a duration (unlimited rows)
for (const row of Csv.generateRows({ columns: 3, duration: 5000 })) {
  // Generates for 5 seconds
}

// Async generator with delay between rows
for await (const row of Csv.generateRowsAsync({ columns: 5, rows: 100, delay: 10 })) {
  console.log(row);
}

// Generate raw data (not CSV strings)
const rawRows = Csv.generateData({ columns: ["name", "int"], rows: 10 });
// [[name, number], ...]

const objects = Csv.generateData({
  columns: [{ type: "name", name: "fullName" }],
  rows: 10,
  objectMode: true
});
// [{ fullName: "..." }, ...]

// Reusable generator with preset config
const gen = Csv.createGenerator({ columns: ["name", "email"], seed: 42 });
const batch1 = gen.generate(100);
const batch2 = gen.generate(100);
```

**Built-in Column Types:**
`string`, `int`, `float`, `bool`, `date`, `uuid`, `email`, `name`, `url`, `phone`, `address`, `ip`, `hex`, `lorem`

---

## Error Classes

```typescript
import { Csv } from "documonster/csv";
import { CsvError } from "documonster/csv";

try {
  Csv.parse(badInput, { headers: true });
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

| Function                                | Description                                                    |
| --------------------------------------- | -------------------------------------------------------------- |
| `Csv.parse(input, options?)`            | Synchronous CSV parser                                         |
| `Csv.parseAsync(input, options?)`       | Async parser (strings, bytes, Blobs, streams, async iterables) |
| `Csv.parseRows(input, options?)`        | Async generator yielding rows one at a time                    |
| `Csv.parseAsync(input, { onProgress })` | Parser with progress reporting                                 |
| `Csv.format(data, options?)`            | Batch CSV formatter                                            |

### Stream Classes

| Class                               | Description                                |
| ----------------------------------- | ------------------------------------------ |
| `new Csv.ParserStream(options?)`    | Transform stream: CSV bytes -> parsed rows |
| `new Csv.FormatterStream(options?)` | Transform stream: rows -> CSV text         |

### Utilities

| Function                                       | Description                   |
| ---------------------------------------------- | ----------------------------- |
| `Csv.detectDelimiter(input)`                   | Auto-detect CSV delimiter     |
| `Csv.detectLinebreak(input)`                   | Auto-detect line terminator   |
| `Csv.stripBom(input)`                          | Strip UTF-8 BOM               |
| `Csv.applyDynamicTyping(value, columnConfig?)` | Convert string to native type |
| `Csv.formatNumber(value, options?)`            | Format number for locale      |
| `Csv.parseNumber(value, options?)`             | Parse locale-formatted number |
| `Csv.deduplicateHeaders(headers)`              | Rename duplicate headers      |

### Generator Functions

| Function                          | Description                         |
| --------------------------------- | ----------------------------------- |
| `Csv.generate(options?)`          | Generate CSV string + data          |
| `Csv.generateRows(options?)`      | Sync generator yielding CSV rows    |
| `Csv.generateRowsAsync(options?)` | Async generator with delay support  |
| `Csv.generateData(options?)`      | Generate raw data (not CSV strings) |
| `Csv.createGenerator(options?)`   | Reusable generator factory          |
