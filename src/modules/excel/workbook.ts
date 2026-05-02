/**
 * Workbook - Node.js Version
 *
 * Extends browser Workbook with Node.js file system support for CSV and Markdown operations.
 */

import { ExcelFileError } from "@excel/errors";
import {
  WorkbookReader,
  type WorkbookReaderOptions,
  type NodeInput
} from "@excel/stream/workbook-reader";
import type { WorkbookReader as WorkbookReaderBrowser } from "@excel/stream/workbook-reader.browser";
import { WorkbookWriter, type WorkbookWriterOptions } from "@excel/stream/workbook-writer";
import type { WorkbookWriter as WorkbookWriterBrowser } from "@excel/stream/workbook-writer.browser";
import { Workbook as WorkbookBrowser, type CsvOptions } from "@excel/workbook.browser";
import type { Worksheet } from "@excel/worksheet";
import { XLSX } from "@excel/xlsx/xlsx";
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
  // Declare the overridden private slot so the getter below can stash
  // the Node-typed XLSX instance without TS complaining that we're
  // widening the base class's private member (which would be illegal
  // anyway — the base declares `_xlsx?: XlsxBrowser`).
  private _xlsxNode?: XLSX;

  /**
   * xlsx file format operations — Node.js variant. Exposes the
   * `readFile` / `writeFile` / streaming `read` / `write` methods that
   * the browser XLSX omits. Overriding here (rather than typing the
   * base getter as Node XLSX directly) keeps the browser
   * `Workbook.xlsx` type clean — browser consumers see only the
   * operations the bundle actually supports and get a TS error if
   * they accidentally reach for file-path APIs.
   */
  override get xlsx(): XLSX {
    if (!this._xlsxNode) {
      this._xlsxNode = new XLSX(this);
    }
    return this._xlsxNode;
  }

  /**
   * Create a streaming workbook writer — Node.js variant. Accepts the
   * Node-only `{ filename }` option for direct file-path output in
   * addition to the cross-platform `{ stream }` option. Overriding
   * here (rather than typing the base factory as the Node writer)
   * keeps the browser bundle free of Node-only stream code.
   *
   * The return type is declared as the browser `WorkbookWriter` to
   * preserve static-side Liskov compatibility with the base class;
   * downcast at the call site if the Node subclass API is needed, or
   * use `new WorkbookWriter()` directly.
   */
  static override createStreamWriter(options?: WorkbookWriterOptions): WorkbookWriterBrowser {
    return new WorkbookWriter(options) as unknown as WorkbookWriterBrowser;
  }

  /**
   * Create a streaming workbook reader — Node.js variant. Accepts a
   * Node-only file-path `string` in addition to the cross-platform
   * `CommonInput` types (buffer / readable). Overriding here keeps the
   * browser bundle free of Node-only `fs` imports.
   *
   * The return type is declared as the browser `WorkbookReader` to
   * preserve static-side Liskov compatibility with the base class.
   * The runtime instance is the Node `WorkbookReader` subclass and
   * handles file-path inputs transparently; downcast if the subclass
   * API is needed, or use `new WorkbookReader()` directly.
   */
  static override createStreamReader(
    input: NodeInput,
    options?: WorkbookReaderOptions
  ): WorkbookReaderBrowser {
    return new WorkbookReader(input, options) as unknown as WorkbookReaderBrowser;
  }

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
export type {
  CsvOptions,
  CsvInput,
  WorkbookModel,
  WorkbookMedia,
  WorkbookProtectionModel,
  ExternalLinkModel,
  ExternalLinkCachedSheet
} from "@excel/workbook.browser";
export type {
  AddChartsheetOptions,
  AddPivotChartsheetOptions,
  ChartsheetOptions,
  ChartsheetViewOptions
} from "@excel/chartsheet";
