/**
 * Converting between the two containers, in both directions.
 *
 * Run: `pnpm example --filter xlsb-convert`
 *
 * Nothing about the *call* changes between formats. `.xlsx` and `.xlsb` are chosen by the file
 * extension when there is one, by `{ format }` when writing to a buffer, and by **looking at the
 * bytes** when reading one — so a consumer handed an unnamed buffer does not have to guess, and
 * cannot guess wrong.
 *
 * The one asymmetry worth stating: XLSB expresses less than XLSX. Content it cannot carry is
 * reported by address, so a conversion tells you what it cost rather than quietly costing it.
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Cell, Workbook, Worksheet } from "@excel/index";

const exampleDir = path.dirname(fileURLToPath(import.meta.url));
const outDir = path.resolve(exampleDir, "../../../../tmp/excel-examples");
fs.mkdirSync(outDir, { recursive: true });
const prefix = process.argv[2] ?? path.join(outDir, "xlsb-convert");

// ---------------------------------------------------------------------------
// Build once, write both ways.
// ---------------------------------------------------------------------------

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "Report");
Worksheet.addRow(ws, ["Region", "Units", "Revenue", "Booked"]);
Worksheet.addRow(ws, ["North", 120, 2398.8, new Date(Date.UTC(2024, 0, 15))]);
Worksheet.addRow(ws, ["South", 84, 2058.0, new Date(Date.UTC(2024, 1, 3))]);
Cell.setValue(ws, "C4", { formula: "SUM(C2:C3)", result: 4456.8 });
for (const row of [2, 3, 4]) {
  Cell.setStyle(ws, `C${row}`, { numFmt: '"$"#,##0.00' });
}
for (const row of [2, 3]) {
  Cell.setStyle(ws, `D${row}`, { numFmt: "yyyy-mm-dd" });
}
Cell.setStyle(ws, "A1", { font: { bold: true } });

const xlsxPath = `${prefix}.xlsx`;
const xlsbPath = `${prefix}.xlsb`;
await Workbook.writeFile(wb, xlsxPath);
await Workbook.writeFile(wb, xlsbPath);
// Read each file back once, and use those bytes for both the size report here and the
// format-detection section below. A `statSync` in front of a `readFileSync` asks the file system the
// same question twice and races between the two answers; the length of what was read is the size.
const xlsxBytes = fs.readFileSync(xlsxPath);
const xlsbBytes = fs.readFileSync(xlsbPath);
console.log(`wrote ${xlsxPath} (${xlsxBytes.byteLength} bytes)`);
console.log(`wrote ${xlsbPath} (${xlsbBytes.byteLength} bytes)`);

// ---------------------------------------------------------------------------
// xlsx → xlsb, and back again.
// ---------------------------------------------------------------------------

const fromXlsx = Workbook.create();
await Workbook.readFile(fromXlsx, xlsxPath);
const asXlsb = `${prefix}-from-xlsx.xlsb`;
await Workbook.writeFile(fromXlsx, asXlsb);

const fromXlsb = Workbook.create();
await Workbook.readFile(fromXlsb, asXlsb);
const backToXlsx = `${prefix}-round-tripped.xlsx`;
await Workbook.writeFile(fromXlsb, backToXlsx);
console.log(`wrote ${asXlsb}`);
console.log(`wrote ${backToXlsx}`);

const read = Workbook.getWorksheets(fromXlsb)[0]!;
console.log(`  C4 after xlsx → xlsb: ${JSON.stringify(Cell.getValue(read, "C4"))}`);
console.log(`  D2 is a Date: ${Cell.getValue(read, "D2") instanceof Date}`);
console.log(`  C2 number format: ${JSON.stringify(Cell.getStyle(read, "C2")?.numFmt)}`);

// ---------------------------------------------------------------------------
// Reading bytes with no name attached.
// ---------------------------------------------------------------------------

// The format is detected from the content, so the same call reads either container.
for (const [label, bytes] of [
  ["xlsx bytes", xlsxBytes],
  ["xlsb bytes", xlsbBytes]
] as const) {
  const detected = Workbook.create();
  await Workbook.read(detected, bytes);
  const sheet = Workbook.getWorksheets(detected)[0]!;
  console.log(
    `read ${label} with no filename: ${Worksheet.getName(sheet)}, C4 = ${JSON.stringify(Cell.getValue(sheet, "C4"))}`
  );
}

// Asking for the wrong format is refused rather than silently producing an empty workbook, which is
// what a reader that trusted the caller over the bytes would do.
try {
  const wrong = Workbook.create();
  await Workbook.read(wrong, xlsbBytes, { format: "xlsx" });
  console.log("unexpected: xlsb read as xlsx without complaint");
} catch (error) {
  console.log(`refused to read xlsb as xlsx: ${(error as Error).message.slice(0, 90)}`);
}
