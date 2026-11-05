/**
 * CSV ↔ Workbook bridge — Node-only file-path variants.
 *
 * Separated from `./csv-bridge` so the browser bundle never pulls in
 * `@utils/fs`. These functions read/write CSV directly from/to a file path.
 */

import { open, stat } from "node:fs/promises";
import type { Readable, Writable } from "node:stream";

import type { ExcelCsvReadOptions, ExcelCsvWriteOptions } from "@excel/bridge/csv-bridge";
import {
  checkCsvReadArguments,
  checkCsvWriteArguments,
  createCsvByteStream,
  createCsvImportParser,
  readCsvStream,
  streamWorksheet
} from "@excel/bridge/csv-bridge";
import { createWorkbook } from "@excel/core/workbook";
import { isReadTarget } from "@excel/core/workbook-io-types";
import type { Workbook } from "@excel/core/workbook.browser";
import type { Worksheet } from "@excel/core/worksheet";
import { ExcelArgumentError } from "@excel/errors";
import { checkFilePath, fileSystemError } from "@excel/utils/file-path";
import type { IReadable, IWritable } from "@stream/types";
import { createReadStream, createWriteStream, safeStats } from "@utils/fs";

/** Read CSV from a file path into a new worksheet on `workbook` (Node.js only). */
export async function readCsvFile(
  workbook: Workbook,
  filename: string,
  options?: ExcelCsvReadOptions
): Promise<Worksheet> {
  checkFilePath("ExcelCsv.readFile", filename);
  checkCsvReadArguments("ExcelCsv.readFile", options);
  // Asked of the file system rather than tested for existence, so its own error — `ENOENT`, `EACCES` — is kept as the
  // cause, as `Workbook.readFile` keeps it.
  try {
    await stat(filename);
  } catch (error) {
    throw fileSystemError(filename, "read", error, "CSV");
  }
  // Bytes, decoded by the parser as `options.encoding` says: the stream used to be opened as UTF-8 text, so the
  // option had no effect on a file.
  const readStream = createReadStream(filename, {
    highWaterMark: options?.highWaterMark ?? 64 * 1024
  });

  return readCsvStream(workbook, readStream as never, options, "ExcelCsv.readFile");
}

/**
 * Read CSV from a file path into a new workbook holding one worksheet, returned (Node.js only).
 *
 * @example
 * ```ts
 * import { ExcelCsv } from "documonster/excel/csv";
 * const workbook = await ExcelCsv.readFile("data.csv", { delimiter: ";", sheetName: "Data" });
 * ```
 */
export async function readCsvFileWorkbook(
  filename: string,
  options?: ExcelCsvReadOptions
): Promise<Workbook>;
/**
 * @deprecated Use `ExcelCsv.addWorksheet(workbook, createReadStream(filename, "utf8"), options)`, which
 * this form does — it adds a worksheet to `workbook` and returns it. `ExcelCsv.readFile(filename)`
 * returns a new workbook. Removed in 1.0.
 */
export async function readCsvFileWorkbook(
  workbook: Workbook,
  filename: string,
  options?: ExcelCsvReadOptions
): Promise<Worksheet>;
export async function readCsvFileWorkbook(
  first: Workbook | string,
  second?: string | ExcelCsvReadOptions,
  third?: ExcelCsvReadOptions
): Promise<Workbook | Worksheet> {
  if (isReadTarget(first)) {
    return readCsvFile(first, second as string, third);
  }
  const workbook = createWorkbook();
  await readCsvFile(workbook, first, second as ExcelCsvReadOptions | undefined);
  return workbook;
}

/**
 * Write a worksheet to a CSV file path (Node.js only). With `append`, a file that exists and is not empty is continued
 * with every row; `writeHeaders: false` leaves out the sheet's first (header) row. See {@link ExcelCsvWriteOptions.append}.
 */
export async function writeCsvFile(
  workbook: Workbook,
  filename: string,
  options?: ExcelCsvWriteOptions
): Promise<void> {
  checkFilePath("ExcelCsv.writeFile", filename);
  // Before the file is opened: a refused option used to leave it created, or truncated.
  checkCsvWriteArguments("ExcelCsv.writeFile", workbook, options);
  const encoding = options?.encoding ?? "utf8";
  if (!isTextEncoding(encoding)) {
    throw new ExcelArgumentError(
      "options.encoding",
      `ExcelCsv.writeFile: options.encoding must be a text encoding Node can write — "utf8", "utf16le", "latin1", "ascii" — got ${JSON.stringify(encoding)}.`
    );
  }
  const existing = options?.append ? await safeStats(filename) : null;
  // A line break first, unless the file already ends with one or holds nothing but a byte-order mark: otherwise its
  // last row and our first would join.
  const breakFirst =
    existing !== null &&
    existing.size > 0 &&
    (await needsLineBreak(filename, existing.size, encoding));

  const writeStream = createWriteStream(filename, {
    encoding,
    highWaterMark: options?.highWaterMark ?? 64 * 1024,
    flags: options?.append ? "a" : "w"
  });
  // Only the file stream's own failure is a file-system error; one from the formatter is passed on as it is.
  let fileFailure: unknown;
  writeStream.once("error", error => {
    fileFailure = error;
  });
  if (breakFirst) {
    writeStream.write(options?.lineEnding ?? "\n");
  }
  try {
    await streamWorksheet("ExcelCsv.writeFile", workbook, writeStream as never, options, true);
  } catch (error) {
    throw error === fileFailure ? fileSystemError(filename, "write", error, "CSV") : error;
  }
}

/**
 * Whether `encoding` is a `Buffer` encoding of text. `hex`, `base64` and `base64url` are `Buffer` encodings too, but
 * of bytes: a CSV "written" in them is not a CSV, and a file in them has no line to continue.
 */
function isTextEncoding(encoding: string): encoding is BufferEncoding {
  return Buffer.isEncoding(encoding) && !/^(?:hex|base64|base64url)$/i.test(encoding);
}

/** The bytes one character takes at the least in `encoding`, so a read of the file's end starts on a character. */
function unitBytes(encoding: BufferEncoding): number {
  return /^(?:utf-?16le|ucs-?2)$/i.test(encoding) ? 2 : 1;
}

/**
 * Whether a file of `size` bytes must be given a line break before rows are appended to it: its last characters,
 * decoded in the encoding the rows are written in, do not end with `\n` or `\r` (so `\r\n` ends a line too), and it
 * holds more than a byte-order mark. Read in that encoding because a line feed is not one byte in all of them: in
 * UTF-16LE it is `0a 00`.
 */
async function needsLineBreak(
  filename: string,
  size: number,
  encoding: BufferEncoding
): Promise<boolean> {
  const unit = unitBytes(encoding);
  const length = Math.min(size - (size % unit), 4 * unit);
  if (length === 0) {
    return false;
  }
  const handle = await open(filename, "r");
  try {
    const tail = Buffer.alloc(length);
    const start = size - (size % unit) - length;
    const { bytesRead } = await handle.read(tail, 0, length, start);
    let text = tail.subarray(0, bytesRead).toString(encoding);
    if (start === 0) {
      text = text.replace(/^\uFEFF/, "");
    }
    return text !== "" && !/[\r\n]$/.test(text);
  } finally {
    await handle.close();
  }
}

/**
 * Create a readable stream that outputs the worksheet as CSV.
 *
 * Same stream as the cross-platform export, typed as the `stream.Readable` Node
 * actually returns, so it goes straight into `stream.pipeline()`,
 * `Readable.toWeb()`, or an SDK upload body — no `as unknown as Readable`.
 *
 * There is no type assertion here: on the Node build `CsvFormatterStream extends
 * Transform extends Readable`, so the compiler verifies the narrowing on the
 * return statement below. Intersecting with `Readable` rather than with the
 * concrete `CsvFormatterStream` is what keeps `write()` / `end()` off the public
 * surface — the stream is already attached to a producer, so a caller who wrote
 * to it would inject rows into the CSV.
 *
 * `IReadable<Uint8Array>` stays first in the intersection so `for await` keeps
 * yielding `Uint8Array` instead of the `any` Node's own declaration returns.
 *
 * Leaving `for await` early has the same Node-level caveat as
 * `Workbook.toStream`: `break` lets Node's async iterator destroy the stream with
 * an `AbortError`, which sets `errored` and reaches any attached `'error'`
 * listener, while calling `destroy()` yourself is silent. Destroy before breaking,
 * or iterate via `stream.iterator({ destroyOnReturn: false })` and destroy in a
 * `finally`.
 */
export function createCsvReadStream(
  workbook: Workbook,
  options?: ExcelCsvWriteOptions
): IReadable<Uint8Array> & Readable {
  return createCsvByteStream(workbook, options);
}

/**
 * Create a writable stream that imports the CSV written into it into a new worksheet on `workbook` — typed as the
 * `stream.Writable` Node actually returns, so `fs.createReadStream("data.csv").pipe(sink)` type-checks. See the
 * cross-platform `ExcelCsv.createImportStream`.
 */
export function createCsvImportStream(
  workbook: Workbook,
  options?: ExcelCsvReadOptions
): IWritable<Uint8Array | string> & Writable {
  return createCsvImportParser(workbook, options);
}
