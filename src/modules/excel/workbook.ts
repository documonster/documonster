/**
 * Workbook - Node.js Version
 *
 * Extends browser Workbook with Node.js file system support for CSV operations.
 */

import { fileExists, createReadStream, createWriteStream } from "@utils/fs";
import { Workbook as WorkbookBrowser, type CsvOptions } from "@excel/workbook.browser";
import { ExcelFileError } from "@excel/errors";
import type { Worksheet } from "@excel/worksheet";

// =============================================================================
// Node.js Workbook Class
// =============================================================================

class Workbook extends WorkbookBrowser {
  /**
   * Read CSV from file (Node.js only)
   *
   * @example
   * ```ts
   * await workbook.readCsvFile("data.csv");
   * await workbook.readCsvFile("data.csv", { delimiter: ";", sheetName: "Data" });
   * ```
   */
  override async readCsvFile(filename: string, options?: CsvOptions): Promise<Worksheet> {
    if (!(await fileExists(filename))) {
      throw new ExcelFileError(filename, "read", "file not found");
    }

    const readStream = createReadStream(filename, {
      encoding: "utf8",
      highWaterMark: options?.highWaterMark ?? 64 * 1024
    });

    return this._readCsvStream(readStream, options);
  }

  /**
   * Write CSV to file (Node.js only)
   *
   * @example
   * ```ts
   * await workbook.writeCsvFile("output.csv");
   * await workbook.writeCsvFile("output.csv", { delimiter: ";", sheetName: "Data" });
   * await workbook.writeCsvFile("output.csv", { append: true }); // Append mode
   * ```
   */
  override async writeCsvFile(filename: string, options?: CsvOptions): Promise<void> {
    const isAppend = options?.append && (await fileExists(filename));

    const writeStream = createWriteStream(filename, {
      encoding: (options?.encoding || "utf8") as BufferEncoding,
      highWaterMark: options?.highWaterMark ?? 64 * 1024,
      flags: options?.append ? "a" : "w"
    });

    // Append mode to existing file: write leading newline and skip headers
    if (isAppend) {
      const lineEnding = options?.lineEnding ?? "\n";
      writeStream.write(lineEnding);
      return this._writeCsvStream(writeStream, {
        ...options,
        writeHeaders: false
      });
    }

    return this._writeCsvStream(writeStream, options);
  }
}

export { Workbook };
export type {
  CsvOptions,
  CsvInput,
  WorkbookModel,
  WorkbookMedia,
  ToPdfOptions
} from "@excel/workbook.browser";
