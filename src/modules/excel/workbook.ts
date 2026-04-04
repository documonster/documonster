/**
 * Workbook - Node.js Version
 *
 * Extends browser Workbook with Node.js file system support for CSV and Markdown operations.
 */

import {
  fileExists,
  createReadStream,
  createWriteStream,
  readFileText,
  writeFileText
} from "@utils/fs";
import { Workbook as WorkbookBrowser, type CsvOptions } from "@excel/workbook.browser";
import { ExcelFileError } from "@excel/errors";
import type { Worksheet } from "@excel/worksheet";
import type { MdOptions } from "@md/types";

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

  /**
   * Read Markdown table from file (Node.js only)
   *
   * @example
   * ```ts
   * await workbook.readMdFile("table.md");
   * await workbook.readMdFile("table.md", { sheetName: "Data" });
   * ```
   */
  override async readMdFile(filename: string, options?: MdOptions): Promise<Worksheet> {
    if (!(await fileExists(filename))) {
      throw new ExcelFileError(filename, "read", "file not found");
    }

    const content = await readFileText(filename);
    return this.readMd(content, options);
  }

  /**
   * Read all Markdown tables from file, each as a separate worksheet (Node.js only)
   *
   * @example
   * ```ts
   * await workbook.readMdAllFile("doc.md");
   * await workbook.readMdAllFile("doc.md", { sheetName: "Table" });
   * ```
   */
  override async readMdAllFile(filename: string, options?: MdOptions): Promise<Worksheet[]> {
    if (!(await fileExists(filename))) {
      throw new ExcelFileError(filename, "read", "file not found");
    }

    const content = await readFileText(filename);
    return this.readMdAll(content, options);
  }

  /**
   * Write Markdown table to file (Node.js only)
   *
   * @example
   * ```ts
   * await workbook.writeMdFile("output.md");
   * await workbook.writeMdFile("output.md", { sheetName: "Data", padding: true });
   * ```
   */
  override async writeMdFile(filename: string, options?: MdOptions): Promise<void> {
    const mdString = this.writeMd(options);
    await writeFileText(filename, mdString);
  }
}

export { Workbook };
export type { CsvOptions, CsvInput, WorkbookModel, WorkbookMedia } from "@excel/workbook.browser";
