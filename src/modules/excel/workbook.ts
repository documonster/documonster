/**
 * Workbook - Node.js Version
 *
 * Extends browser Workbook with Node.js file system support for CSV and Markdown operations.
 */

import { ExcelFileError } from "@excel/errors";
import { Workbook as WorkbookBrowser, type CsvOptions } from "@excel/workbook.browser";
import type { Worksheet } from "@excel/worksheet";
import type { MarkdownOptions } from "@markdown/types";
import {
  fileExists,
  createReadStream,
  createWriteStream,
  readFileText,
  writeFileText
} from "@utils/fs";

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
   * await workbook.readMarkdownFile("table.md");
   * await workbook.readMarkdownFile("table.md", { sheetName: "Data" });
   * ```
   */
  override async readMarkdownFile(filename: string, options?: MarkdownOptions): Promise<Worksheet> {
    if (!(await fileExists(filename))) {
      throw new ExcelFileError(filename, "read", "file not found");
    }

    const content = await readFileText(filename);
    return this.readMarkdown(content, options);
  }

  /**
   * Read all Markdown tables from file, each as a separate worksheet (Node.js only)
   *
   * @example
   * ```ts
   * await workbook.readMarkdownAllFile("doc.md");
   * await workbook.readMarkdownAllFile("doc.md", { sheetName: "Table" });
   * ```
   */
  override async readMarkdownAllFile(
    filename: string,
    options?: MarkdownOptions
  ): Promise<Worksheet[]> {
    if (!(await fileExists(filename))) {
      throw new ExcelFileError(filename, "read", "file not found");
    }

    const content = await readFileText(filename);
    return this.readMarkdownAll(content, options);
  }

  /**
   * Write Markdown table to file (Node.js only)
   *
   * @example
   * ```ts
   * await workbook.writeMarkdownFile("output.md");
   * await workbook.writeMarkdownFile("output.md", { sheetName: "Data", padding: true });
   * ```
   */
  override async writeMarkdownFile(filename: string, options?: MarkdownOptions): Promise<void> {
    const markdownString = this.writeMarkdown(options);
    await writeFileText(filename, markdownString);
  }
}

export { Workbook };
export type { CsvOptions, CsvInput, WorkbookModel, WorkbookMedia } from "@excel/workbook.browser";
