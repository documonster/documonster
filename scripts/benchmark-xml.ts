/**
 * XML Performance Benchmark: excelts vs fast-xml-parser
 *
 * Compares:
 * 1. DOM parse (non-streaming): parseXml() vs XMLParser.parse()
 * 2. SAX parse (streaming): SaxParser callback vs (no equivalent in fxp)
 * 3. Write (buffered): XmlWriter vs XMLBuilder
 *
 * Run: npx tsx scripts/xml-benchmark.ts
 */

import { SaxParser, parseXml, XmlWriter } from "../dist/esm/modules/xml/index.js";
import { XMLParser, XMLBuilder } from "fast-xml-parser";

// =============================================================================
// Test Data Generation
// =============================================================================

function generateXml(rowCount: number, colCount: number): string {
  const parts: string[] = [];
  parts.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  parts.push('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
  parts.push("<sheetData>");
  for (let r = 1; r <= rowCount; r++) {
    parts.push(`<row r="${r}">`);
    for (let c = 1; c <= colCount; c++) {
      const col = String.fromCharCode(64 + c);
      parts.push(`<c r="${col}${r}" t="s"><v>${(r - 1) * colCount + c}</v></c>`);
    }
    parts.push("</row>");
  }
  parts.push("</sheetData>");
  parts.push("</worksheet>");
  return parts.join("");
}

interface RowData {
  row: number;
  cells: Array<{ col: string; ref: string; type: string; value: string }>;
}

function generateRowData(rowCount: number, colCount: number): RowData[] {
  const rows: RowData[] = [];
  for (let r = 1; r <= rowCount; r++) {
    const cells: RowData["cells"] = [];
    for (let c = 1; c <= colCount; c++) {
      const col = String.fromCharCode(64 + c);
      cells.push({
        col,
        ref: `${col}${r}`,
        type: "s",
        value: String((r - 1) * colCount + c)
      });
    }
    rows.push({ row: r, cells });
  }
  return rows;
}

// =============================================================================
// Benchmark Helpers
// =============================================================================

function bench(name: string, fn: () => void, iterations: number): { avg: number; min: number } {
  // Warmup
  for (let i = 0; i < Math.min(3, iterations); i++) {
    fn();
  }

  const times: number[] = [];
  for (let i = 0; i < iterations; i++) {
    const start = performance.now();
    fn();
    times.push(performance.now() - start);
  }

  times.sort((a, b) => a - b);
  const avg = times.reduce((s, t) => s + t, 0) / times.length;
  const min = times[0];
  return { avg, min };
}

function formatMs(ms: number): string {
  return ms < 1 ? `${(ms * 1000).toFixed(0)}µs` : `${ms.toFixed(2)}ms`;
}

function printResult(
  label: string,
  ours: { avg: number; min: number },
  theirs: { avg: number; min: number }
) {
  const ratio = theirs.min / ours.min;
  const winner = ratio > 1 ? "excelts" : "fxp";
  const factor = ratio > 1 ? ratio : 1 / ratio;
  console.log(`  ${label}`);
  console.log(`    excelts:  avg ${formatMs(ours.avg)}, min ${formatMs(ours.min)}`);
  console.log(`    fxp:      avg ${formatMs(theirs.avg)}, min ${formatMs(theirs.min)}`);
  console.log(`    → ${winner} is ${factor.toFixed(2)}x faster (by min)`);
  console.log();
}

function printSoloResult(label: string, result: { avg: number; min: number }) {
  console.log(`  ${label}`);
  console.log(`    excelts:  avg ${formatMs(result.avg)}, min ${formatMs(result.min)}`);
  console.log(`    fxp:      N/A (no SAX/streaming mode)`);
  console.log();
}

// =============================================================================
// Benchmarks
// =============================================================================

console.log("=".repeat(70));
console.log("XML Performance Benchmark: excelts vs fast-xml-parser");
console.log("=".repeat(70));
console.log();

const sizes = [
  { name: "Small (100 rows × 5 cols)", rows: 100, cols: 5, iters: 100 },
  { name: "Medium (1K rows × 10 cols)", rows: 1000, cols: 10, iters: 20 },
  { name: "Large (10K rows × 10 cols)", rows: 10000, cols: 10, iters: 5 },
  { name: "XL (50K rows × 10 cols)", rows: 50000, cols: 10, iters: 3 }
];

for (const size of sizes) {
  const xml = generateXml(size.rows, size.cols);
  const xmlSizeKB = (Buffer.byteLength(xml) / 1024).toFixed(0);
  console.log(`--- ${size.name} (${xmlSizeKB} KB) ---`);
  console.log();

  // -------------------------------------------------------------------------
  // 1. DOM Parse (non-streaming)
  // -------------------------------------------------------------------------

  const fxpParser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "",
    parseAttributeValue: false
  });

  const domOurs = bench(
    "excelts parseXml",
    () => {
      parseXml(xml);
    },
    size.iters
  );

  const domTheirs = bench(
    "fxp XMLParser.parse",
    () => {
      fxpParser.parse(xml);
    },
    size.iters
  );

  printResult("DOM Parse", domOurs, domTheirs);

  // -------------------------------------------------------------------------
  // 2. SAX Parse (streaming / callback)
  // -------------------------------------------------------------------------

  const saxOurs = bench(
    "excelts SaxParser",
    () => {
      let tagCount = 0;
      const parser = new SaxParser({ position: false });
      parser.on("opentag", () => {
        tagCount++;
      });
      parser.on("text", () => {});
      parser.on("closetag", () => {});
      parser.write(xml);
      parser.close();
      // Prevent dead-code elimination
      if (tagCount === 0) throw new Error("unreachable");
    },
    size.iters
  );

  printSoloResult("SAX Parse (callback)", saxOurs);

  // -------------------------------------------------------------------------
  // 3. Write (buffered)
  // -------------------------------------------------------------------------

  const rowData = generateRowData(size.rows, size.cols);

  const writeOurs = bench(
    "excelts XmlWriter",
    () => {
      const w = new XmlWriter();
      w.openXml();
      w.openNode("worksheet", {
        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      });
      w.openNode("sheetData");
      for (const row of rowData) {
        w.openNode("row", { r: row.row });
        for (const cell of row.cells) {
          w.openNode("c", { r: cell.ref, t: cell.type });
          w.leafNode("v", undefined, cell.value);
          w.closeNode();
        }
        w.closeNode();
      }
      w.closeNode();
      w.closeNode();
      const result = w.xml;
      if (result.length === 0) throw new Error("unreachable");
    },
    size.iters
  );

  const writeTheirs = bench(
    "fxp XMLBuilder",
    () => {
      const rows = rowData.map(row => ({
        "@_r": row.row,
        c: row.cells.map(cell => ({
          "@_r": cell.ref,
          "@_t": cell.type,
          v: cell.value
        }))
      }));
      const builder = new XMLBuilder({
        ignoreAttributes: false,
        attributeNamePrefix: "@_",
        format: false
      });
      const result = builder.build({
        "?xml": { "@_version": "1.0", "@_encoding": "UTF-8", "@_standalone": "yes" },
        worksheet: {
          "@_xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          sheetData: { row: rows }
        }
      }) as string;
      if (result.length === 0) throw new Error("unreachable");
    },
    size.iters
  );

  printResult("Write (buffered)", writeOurs, writeTheirs);
}

// =============================================================================
// Summary
// =============================================================================

console.log("=".repeat(70));
console.log("Notes:");
console.log("  - fast-xml-parser has no SAX/streaming parse mode");
console.log("  - excelts SaxParser shown solo for reference");
console.log("  - 'min' is the best single run (least noise)");
console.log("  - DOM parse comparison is parseXml() vs XMLParser.parse()");
console.log("  - Write comparison is XmlWriter vs XMLBuilder");
console.log("=".repeat(70));
