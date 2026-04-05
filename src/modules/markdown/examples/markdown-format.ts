/**
 * Example: Markdown Formatting
 *
 * Covers:
 * - Basic formatting with padding
 * - Compact mode (padding: false)
 * - Column alignment: global and per-column
 * - Column config: minWidth, header rename
 * - Custom stringify (currency, dates)
 * - Auto-escaping of pipes and backslashes
 * - Multiline cells (newlines → <br>)
 * - CJK / emoji display width alignment
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { formatMarkdown } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/markdown-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const sections: string[] = [];

// =============================================================================
// 1. Basic padded table
// =============================================================================

const table1 = formatMarkdown(
  ["Name", "Age", "City"],
  [
    ["Alice", "30", "New York"],
    ["Bob", "25", "London"],
    ["Carol", "35", "Tokyo"]
  ]
);
sections.push("## 1. Basic Padded Table\n\n" + table1);
console.log("=== 1. Basic ===");
console.log(table1);

// =============================================================================
// 2. Compact mode
// =============================================================================

const table2 = formatMarkdown(
  ["Name", "Age"],
  [
    ["Alice", "30"],
    ["Bob", "25"]
  ],
  { padding: false }
);
sections.push("## 2. Compact Mode\n\n" + table2);
console.log("=== 2. Compact ===");
console.log(table2);

// =============================================================================
// 3. Alignment
// =============================================================================

const table3 = formatMarkdown(
  ["Name", "Amount", "Status"],
  [
    ["Alice", "$1,234.56", "Active"],
    ["Bob", "$567.89", "Pending"],
    ["Carol", "$12,345.00", "Active"]
  ],
  {
    columns: [
      { header: "Name", alignment: "left" },
      { header: "Amount", alignment: "right" },
      { header: "Status", alignment: "center" }
    ]
  }
);
sections.push("## 3. Mixed Alignment\n\n" + table3);
console.log("=== 3. Alignment ===");
console.log(table3);

// =============================================================================
// 4. MinWidth + header rename
// =============================================================================

const table4 = formatMarkdown(
  ["id", "x"],
  [
    ["1", "a"],
    ["2", "b"]
  ],
  {
    columns: [
      { header: "ID", alignment: "right", minWidth: 8 },
      { header: "Value", alignment: "left", minWidth: 10 }
    ]
  }
);
sections.push("## 4. MinWidth + Rename\n\n" + table4);
console.log("=== 4. MinWidth ===");
console.log(table4);

// =============================================================================
// 5. Custom stringify
// =============================================================================

const table5 = formatMarkdown(
  ["Item", "Price", "Qty"],
  [
    ["Widget", 9.99, 100],
    ["Gadget", 24.5, 50]
  ],
  {
    stringify: (value: unknown) => {
      if (typeof value === "number") {
        return value.toLocaleString("en-US", { style: "currency", currency: "USD" });
      }
      return String(value ?? "");
    }
  }
);
sections.push("## 5. Custom Stringify (Currency)\n\n" + table5);
console.log("=== 5. Custom Stringify ===");
console.log(table5);

// =============================================================================
// 6. Auto-escaping
// =============================================================================

const table6 = formatMarkdown(
  ["Expression", "Path"],
  [
    ["a | b", "C:\\Users\\Alice"],
    ["x || y", "D:\\Data\\test"]
  ]
);
sections.push("## 6. Auto-Escaping\n\n" + table6);
console.log("=== 6. Escaping ===");
console.log(table6);

// =============================================================================
// 7. Multiline cells
// =============================================================================

const table7 = formatMarkdown(
  ["Name", "Address"],
  [
    ["Alice", "123 Main St\nApt 4\nNew York"],
    ["Bob", "456 Oak Ave\nLondon"]
  ]
);
sections.push("## 7. Multiline Cells\n\n" + table7);
console.log("=== 7. Multiline ===");
console.log(table7);

// =============================================================================
// 8. CJK / Emoji width
// =============================================================================

const table8 = formatMarkdown(
  ["Name", "名前", "Status"],
  [
    ["Alice", "アリス", "✅ Pass"],
    ["太郎", "太郎", "❌ Fail"]
  ]
);
sections.push("## 8. CJK / Emoji\n\n" + table8);
console.log("=== 8. CJK / Emoji ===");
console.log(table8);

// =============================================================================
// 9. Mixed types
// =============================================================================

const table9 = formatMarkdown(
  ["Name", "Age", "Active", "Score", "Joined"],
  [
    ["Alice", 30, true, 95.5, new Date("2024-01-15")],
    ["Bob", 25, false, null, undefined]
  ]
);
sections.push("## 9. Mixed Types\n\n" + table9);
console.log("=== 9. Mixed Types ===");
console.log(table9);

// Write all formatted tables to a single file
const allOutput = "# Markdown Format Examples\n\n" + sections.join("\n---\n\n");
const outFile = path.join(outDir, "formatted-tables.md");
fs.writeFileSync(outFile, allOutput, "utf8");
console.log(`Wrote: ${outFile}`);
