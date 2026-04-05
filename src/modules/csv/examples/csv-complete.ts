/**
 * Example: CSV Module — Complete Guide
 *
 * Covers:
 * - Parsing CSV strings (sync and async)
 * - Output modes: arrays vs objects (headers option)
 * - Delimiters, quoting, escaping, BOM
 * - Header transforms and row transforms
 * - Dynamic typing (numbers, booleans, dates)
 * - Column mismatch handling
 * - Row validation and filtering
 * - Formatting arrays and objects to CSV
 * - Custom quoting, type transforms, formula escaping
 * - Streaming: CsvParserStream and CsvFormatterStream
 * - Data generation with csvGenerate
 * - Utility functions: detectDelimiter, detectLinebreak, stripBom
 * - Number formatting with decimal separators
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Readable, pipeline } from "node:stream";
import { promisify } from "node:util";
import {
  parseCsv,
  parseCsvAsync,
  formatCsv,
  createCsvParserStream,
  createCsvFormatterStream,
  csvGenerate,
  detectDelimiter,
  detectLinebreak,
  stripBom,
  formatNumberForCsv,
  parseNumberFromCsv
} from "../index";

const pipelineAsync = promisify(pipeline);
const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/csv-examples"
);
fs.mkdirSync(outDir, { recursive: true });

// =============================================================================
// 1. Basic parsing — string[][] output
// =============================================================================

const csv1 = `Name,Age,City
Alice,30,New York
Bob,25,London
Carol,35,Tokyo`;

const rows1 = parseCsv(csv1);
console.log("=== 1. Basic Parsing (arrays) ===");
console.log(rows1);
// [["Name","Age","City"], ["Alice","30","New York"], ...]

// =============================================================================
// 2. Parsing with headers — object output
// =============================================================================

const result2 = parseCsv(csv1, { headers: true });
console.log("\n=== 2. Object Output (headers: true) ===");
console.log(result2.rows);
// [{ Name: "Alice", Age: "30", City: "New York" }, ...]
console.log("Meta:", result2.meta);

// =============================================================================
// 3. Custom delimiters — TSV, semicolon, pipe
// =============================================================================

const tsv = `Name\tAge\tCity\nAlice\t30\tNew York`;
const rows3 = parseCsv(tsv, { delimiter: "\t" });
console.log("\n=== 3. TSV Parsing ===");
console.log(rows3);

// Auto-detect delimiter
const semicolonCsv = `Name;Age;City\nAlice;30;New York`;
const detected = detectDelimiter(semicolonCsv);
console.log("Detected delimiter:", JSON.stringify(detected));

const rows3b = parseCsv(semicolonCsv, { delimiter: detected });
console.log(rows3b);

// =============================================================================
// 4. Dynamic typing — numbers, booleans, dates
// =============================================================================

const csv4 = `Name,Score,Active,Joined
Alice,98.5,true,2023-01-15
Bob,87,false,2022-06-30
Carol,92.3,true,2024-03-01`;

const result4 = parseCsv(csv4, {
  headers: true,
  dynamicTyping: true
});
console.log("\n=== 4. Dynamic Typing ===");
console.log(result4.rows);
// Values are now number/boolean instead of strings

// Selective dynamic typing — only specific columns
const result4b = parseCsv(csv4, {
  headers: true,
  dynamicTyping: { Score: true, Active: true }
});
console.log("Selective:", result4b.rows[0]);

// =============================================================================
// 5. Header transforms — rename, deduplicate
// =============================================================================

const csv5 = `first name,last name,first name
Alice,Smith,Bob`;

const result5 = parseCsv(csv5, {
  headers: h => h.map(name => name.trim().replace(/\s+/g, "_").toLowerCase())
});
console.log("\n=== 5. Header Transforms ===");
console.log("Headers:", result5.meta.fields);
// ["first_name", "last_name", "first_name_1"] (auto-deduplicated)
console.log(result5.rows);

// =============================================================================
// 6. Row transforms and validation
// =============================================================================

const csv6 = `Name,Score
Alice,95
Bob,42
Carol,78
Dave,31`;

const result6 = parseCsv(csv6, {
  headers: true,
  dynamicTyping: true,
  rowTransform: row => {
    // Uppercase the Name field
    const rec = row as Record<string, unknown>;
    return { ...rec, Name: (rec.Name as string).toUpperCase() };
  },
  validate: row => {
    // Only keep rows where Score >= 50
    const rec = row as Record<string, unknown>;
    return (rec.Score as number) >= 50;
  }
});
console.log("\n=== 6. Row Transforms and Validation ===");
console.log(result6.rows);
// Only Alice (95), Carol (78) — names uppercased, Bob (42) and Dave (31) filtered

// =============================================================================
// 7. Column mismatch handling
// =============================================================================

const csv7 = `A,B,C
1,2,3
4,5
6,7,8,9`;

const result7 = parseCsv(csv7, {
  headers: true,
  columnMismatch: { less: "pad", more: "truncate" }
});
console.log("\n=== 7. Column Mismatch ===");
console.log(result7.rows);
// Row 2 padded with empty string for C; Row 3 extra column truncated

// =============================================================================
// 8. Quoting, escaping, and special characters
// =============================================================================

const csv8 = `Name,Description
Alice,"Contains ""quotes"" inside"
Bob,"Multi
line value"
Carol,"Value with, commas"`;

const result8 = parseCsv(csv8, { headers: true });
console.log("\n=== 8. Quoting and Escaping ===");
console.log(result8.rows);

// =============================================================================
// 9. BOM handling and linebreak detection
// =============================================================================

const bomCsv = "\uFEFFName,Age\nAlice,30";
const stripped = stripBom(bomCsv);
console.log("\n=== 9. BOM Handling ===");
console.log("Has BOM:", bomCsv.charCodeAt(0) === 0xfeff);
console.log("After strip:", stripped.substring(0, 4));

const linebreak = detectLinebreak("a\r\nb\r\nc");
console.log("Detected linebreak:", JSON.stringify(linebreak));

// =============================================================================
// 10. Formatting arrays to CSV
// =============================================================================

const data10 = [
  ["Name", "Age", "City"],
  ["Alice", 30, "New York"],
  ["Bob", 25, "London"]
];

const csvOut10 = formatCsv(data10);
console.log("\n=== 10. Format Arrays ===");
console.log(csvOut10);

fs.writeFileSync(path.join(outDir, "formatted-arrays.csv"), csvOut10);

// =============================================================================
// 11. Formatting objects to CSV
// =============================================================================

const data11 = [
  { name: "Alice", score: 98.5, active: true },
  { name: "Bob", score: 87, active: false },
  { name: "Carol", score: 92.3, active: true }
];

const csvOut11 = formatCsv(data11, {
  headers: true,
  columns: ["name", "score", "active"]
});
console.log("\n=== 11. Format Objects ===");
console.log(csvOut11);

fs.writeFileSync(path.join(outDir, "formatted-objects.csv"), csvOut11);

// =============================================================================
// 12. Format options — delimiter, quoting, BOM, newline
// =============================================================================

const csvOut12 = formatCsv(data11, {
  headers: true,
  delimiter: ";",
  quote: "'",
  lineEnding: "\r\n",
  bom: true,
  quoteHeaders: true
});
console.log("\n=== 12. Format Options ===");
console.log(csvOut12);

fs.writeFileSync(path.join(outDir, "formatted-options.csv"), csvOut12);

// =============================================================================
// 13. Formula escaping (CSV injection protection)
// =============================================================================

const data13 = [
  { formula: "=SUM(A1:A10)", value: "Normal" },
  { formula: "+cmd|'/C calc'!Z0", value: "Malicious" }
];

const csvOut13 = formatCsv(data13, { headers: true, escapeFormulae: true });
console.log("\n=== 13. Formula Escaping ===");
console.log(csvOut13);

// =============================================================================
// 14. Column-level quoting control
// =============================================================================

const data14 = [
  { name: "Alice", city: "New York", score: "95" },
  { name: "Bob", city: "London", score: "87" }
];

// Quote only specific columns by name
const csvOut14 = formatCsv(data14, {
  headers: true,
  quoteColumns: { name: true, city: true, score: false }
});
console.log("\n=== 14. Column-Level Quoting ===");
console.log(csvOut14);

// =============================================================================
// 15. Streaming parser — CsvParserStream
// =============================================================================

console.log("\n=== 15. Streaming Parser ===");

const csvStream15 = `Name,Age,City
Alice,30,New York
Bob,25,London
Carol,35,Tokyo`;

const parserStream = createCsvParserStream({ headers: true });
const rows15: Record<string, string>[] = [];

parserStream.on("data", (row: Record<string, string>) => rows15.push(row));

await new Promise<void>((resolve, reject) => {
  const readable = Readable.from([csvStream15]);
  pipelineAsync(readable, parserStream).then(resolve).catch(reject);
});

console.log("Streamed rows:", rows15.length);
console.log("First row:", rows15[0]);

// =============================================================================
// 16. Streaming formatter — CsvFormatterStream
// =============================================================================

console.log("\n=== 16. Streaming Formatter ===");

const outputPath = path.join(outDir, "streamed-output.csv");
const formatterStream = createCsvFormatterStream({
  headers: true,
  columns: ["id", "name", "value"]
});
const writeStream = fs.createWriteStream(outputPath);

formatterStream.pipe(writeStream);

formatterStream.write({ id: "1", name: "Alpha", value: "100" });
formatterStream.write({ id: "2", name: "Beta", value: "200" });
formatterStream.write({ id: "3", name: "Gamma", value: "300" });
formatterStream.end();

await new Promise<void>(resolve => writeStream.on("finish", resolve));
console.log("Written:", outputPath);

// =============================================================================
// 17. Data generation
// =============================================================================

console.log("\n=== 17. Data Generation ===");

const generated = csvGenerate({
  columns: [
    { name: "id", type: "int", min: 1, max: 10000 },
    { name: "name", type: "name" },
    { name: "email", type: "email" },
    { name: "score", type: "float", min: 0, max: 100 },
    { name: "active", type: "bool" }
  ],
  rows: 5,
  seed: 42,
  headers: true,
  delimiter: ","
});

console.log("Generated CSV:");
console.log(generated.csv.substring(0, 300) + "...");
console.log("Headers:", generated.headers);

fs.writeFileSync(path.join(outDir, "generated.csv"), generated.csv);

// =============================================================================
// 18. Async parsing from stream
// =============================================================================

console.log("\n=== 18. Async Parsing ===");

const csvFile = path.join(outDir, "generated.csv");
const fileStream = fs.createReadStream(csvFile, { encoding: "utf-8" });

const result18 = await parseCsvAsync(fileStream, { headers: true });
console.log("Async parsed rows:", result18.rows.length);
console.log("Fields:", result18.meta.fields?.slice(0, 3));

// =============================================================================
// 19. Number formatting with decimal separators
// =============================================================================

console.log("\n=== 19. Number Formatting ===");

// European format (comma as decimal separator)
const eurFormatted = formatNumberForCsv(1234.56, ",");
console.log("European:", eurFormatted); // "1234,56"

// Parse back
const eurParsed = parseNumberFromCsv("1234,56", ",");
console.log("Parsed:", eurParsed); // 1234.56

// US format (period as decimal separator) — default
const usFormatted = formatNumberForCsv(1234.56, ".");
console.log("US:", usFormatted); // "1234.56"

// =============================================================================
// 20. Skip/comment lines and parse info
// =============================================================================

console.log("\n=== 20. Comments and Skip Lines ===");

const csv20 = `# This is a comment
Name,Score
Alice,95
# Another comment
Bob,87`;

const result20 = parseCsv(csv20, {
  headers: true,
  comment: "#"
});
console.log("Rows (comments skipped):", result20.rows);

// Parse with record info
const result20b = parseCsv("a,b\n1,2\n3,4", { info: true });
console.log("With info:", result20b.rows[0]);

console.log("\n=== CSV Examples Complete ===");
