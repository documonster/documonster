import fs from "node:fs";
import path from "node:path";

import { rowCommit } from "@excel/worksheet";

/**
 * Using WorkbookWriter (streaming) for large data
 *
 * WorkbookWriter writes data to disk incrementally, reducing peak memory use
 */
import { WorkbookWriter } from "../../../index";

async function main() {
  console.time("xlsx");

  const ROW_COUNT = 500000;
  const COL_COUNT = 20;

  console.log(`Writing ${ROW_COUNT} rows x ${COL_COUNT} columns = ${ROW_COUNT * COL_COUNT} cells`);

  // Create workbook writer (streaming mode)
  const outDir = path.resolve(process.cwd(), "out");
  fs.mkdirSync(outDir, { recursive: true });

  const book = new WorkbookWriter({
    filename: path.join(outDir, "streaming-writer-large-benchmark.xlsx"),
    useStyles: false,
    useSharedStrings: false,
    zip: { zlib: { level: 1 } }
  });

  console.timeLog("xlsx", "Created workbook writer");

  const sheet = book.addWorksheet("data");

  // Generate header
  const keys: string[] = [];
  for (let i = 0; i < COL_COUNT; i++) {
    keys.push("col" + i);
  }

  // Write header
  rowCommit(sheet.addRow(keys));

  // Write data rows
  console.timeLog("xlsx", "Starting to write rows...");

  for (let i = 0; i < ROW_COUNT; i++) {
    const row = keys.map(() => Math.random().toString());
    rowCommit(sheet.addRow(row));

    // Progress indicator every 100k rows
    if ((i + 1) % 100000 === 0) {
      console.timeLog("xlsx", `Wrote ${i + 1} rows`);
    }
  }

  console.timeLog("xlsx", "Committing worksheet...");
  sheet.commit();

  console.timeLog("xlsx", "Committing workbook...");
  await book.commit();

  console.timeLog("xlsx", "Done!");
}

main().catch(console.error);
