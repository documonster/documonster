/**
 * CSV ↔ Workbook bridge — Node-only file-path variants.
 *
 * Separated from `./csv-bridge` so the browser bundle never pulls in
 * `@utils/fs`. These functions read/write CSV directly from/to a file path.
 */

import { ExcelFileError } from "@excel/errors";
import type { Workbook } from "@excel/workbook.browser";
import type { Worksheet } from "@excel/worksheet";
import { fileExists, createReadStream, createWriteStream } from "@utils/fs";

import { readCsvStream, writeCsvStream, type CsvOptions } from "./csv-bridge";

/**
 * Read CSV from a file path into a new worksheet (Node.js only).
 *
 * @example
 * ```ts
 * import { readCsvFile } from "documonster/excel/csv";
 * await readCsvFile(workbook, "data.csv", { delimiter: ";", sheetName: "Data" });
 * ```
 */
export async function readCsvFile(
  workbook: Workbook,
  filename: string,
  options?: CsvOptions
): Promise<Worksheet> {
  if (!(await fileExists(filename))) {
    throw new ExcelFileError(filename, "read", "file not found");
  }

  const readStream = createReadStream(filename, {
    encoding: "utf8",
    highWaterMark: options?.highWaterMark ?? 64 * 1024
  });

  return readCsvStream(workbook, readStream as never, options);
}

/**
 * Write a worksheet to a CSV file path (Node.js only). Supports append mode.
 */
export async function writeCsvFile(
  workbook: Workbook,
  filename: string,
  options?: CsvOptions
): Promise<void> {
  const isAppend = options?.append && (await fileExists(filename));

  const writeStream = createWriteStream(filename, {
    encoding: (options?.encoding || "utf8") as BufferEncoding,
    highWaterMark: options?.highWaterMark ?? 64 * 1024,
    flags: options?.append ? "a" : "w"
  });

  if (isAppend) {
    const lineEnding = options?.lineEnding ?? "\n";
    writeStream.write(lineEnding);
    return writeCsvStream(workbook, writeStream as never, {
      ...options,
      writeHeaders: false
    });
  }

  return writeCsvStream(workbook, writeStream as never, options);
}
