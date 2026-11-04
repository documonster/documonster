#!/usr/bin/env node
/**
 * Encrypted workbook example
 *
 * Writes a workbook encrypted with a password — the Agile Encryption Excel writes for "Encrypt with Password", which
 * Excel opens with that password — reads it back with the password, and shows what a read without it, or with a wrong
 * one, reports.
 *
 * Usage:
 *   pnpm example --filter encrypted-workbook [outputPath]
 */

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Cell, ExcelDecryptionError, ExcelEncryptedError, Workbook } from "@excel/index";

async function main(): Promise<void> {
  const outDir = path.resolve(
    path.dirname(fileURLToPath(import.meta.url)),
    "../../../../tmp/excel-examples"
  );
  fs.mkdirSync(outDir, { recursive: true });
  const outputPath = process.argv[2] || path.join(outDir, "encrypted-workbook.xlsx");
  const password = "s3cret";

  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "Salaries");
  Cell.setValue(ws, "A1", "Name");
  Cell.setValue(ws, "B1", "Salary");
  Cell.setValue(ws, "A2", "Ada");
  Cell.setValue(ws, "B2", 123456);

  // Encrypt on write: Office's defaults (AES-256, SHA-512, 100,000 iterations) unless `encryption` says otherwise.
  await Workbook.writeFile(wb, outputPath, { password });
  console.log(`Wrote ${outputPath} (open it in Excel with the password "${password}")`);

  // Read it back with the password.
  const again = await Workbook.readFile(outputPath, { password });
  console.log("B2 =", Cell.getValue(Workbook.getWorksheet(again, "Salaries")!, "B2"));

  // The same, in memory and as XLSB, with tuned encryption.
  const bytes = await Workbook.toBuffer(wb, {
    format: "xlsb",
    password,
    encryption: { keyBits: 128, hashAlgorithm: "SHA-256", spinCount: 50_000 }
  });
  const fromXlsb = await Workbook.read(bytes, { password });
  console.log("XLSB A2 =", Cell.getValue(Workbook.getWorksheet(fromXlsb, "Salaries")!, "A2"));

  // Without the password, and with a wrong one.
  for (const options of [{}, { password: "wrong" }]) {
    try {
      await Workbook.read(bytes, options);
    } catch (error) {
      if (error instanceof ExcelEncryptedError) {
        console.log("No password:", error.message);
      } else if (error instanceof ExcelDecryptionError) {
        console.log(`Wrong password (reason "${error.reason}"):`, error.message);
      } else {
        throw error;
      }
    }
  }
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
