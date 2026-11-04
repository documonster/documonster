/**
 * Markdown ↔ Workbook bridge — Node-only file-path variants.
 *
 * Separated from `./markdown-bridge` so the browser bundle never pulls in
 * `@utils/fs`.
 */

import type { MarkdownReadOptions, MarkdownWriteOptions } from "@excel/bridge/markdown-bridge";
import {
  formatWorksheet,
  readMarkdown,
  readMarkdownAll,
  readMarkdownInto,
  readNewWorkbook
} from "@excel/bridge/markdown-bridge";
import { isReadTarget } from "@excel/core/workbook-io-types";
import type { Workbook } from "@excel/core/workbook.browser";
import type { Worksheet } from "@excel/core/worksheet";
import { checkFilePath, fileSystemError } from "@excel/utils/file-path";
import type { MarkdownOptions } from "@markdown/types";
import { readFileText, writeFileText } from "@utils/fs";

/** Read a Markdown table from a file path into a new worksheet (Node.js only). */
export async function readMarkdownFile(
  workbook: Workbook,
  filename: string,
  options?: MarkdownOptions
): Promise<Worksheet> {
  return readMarkdown(workbook, await markdownText("readMarkdownFile", filename), options);
}

/** Read all Markdown tables from a file path, each as a worksheet (Node.js only). */
export async function readMarkdownAllFile(
  workbook: Workbook,
  filename: string,
  options?: MarkdownOptions
): Promise<Worksheet[]> {
  return readMarkdownAll(workbook, await markdownText("readMarkdownAllFile", filename), options);
}

/**
 * The text of a Markdown file. A path that is not one is refused with `ExcelArgumentError`; a file the system will not
 * read is an `ExcelFileError` naming a Markdown file, with the system's error — `ENOENT`, `EACCES` — as its cause.
 */
async function markdownText(where: string, filename: string): Promise<string> {
  checkFilePath(where, filename);
  try {
    return await readFileText(filename);
  } catch (error) {
    throw fileSystemError(filename, "read", error, "Markdown");
  }
}

/**
 * Read the first Markdown table in a file into a new workbook, returned (Node.js only) — with `{ tables: "all" }`
 * every table, one worksheet each. Values are read as `ExcelMarkdown.addWorksheet` reads them.
 */
export async function readMarkdownFileWorkbook(
  filename: string,
  options?: MarkdownReadOptions
): Promise<Workbook>;
/**
 * @deprecated Use `ExcelMarkdown.addWorksheet(workbook, await readFile(filename, "utf8"), options)`, which this
 * form does. `ExcelMarkdown.readFile(filename)` returns a new workbook. Removed in 1.0.
 */
export async function readMarkdownFileWorkbook(
  workbook: Workbook,
  filename: string,
  options?: MarkdownOptions
): Promise<Worksheet>;
export async function readMarkdownFileWorkbook(
  first: Workbook | string,
  second?: string | MarkdownOptions,
  third?: MarkdownOptions
): Promise<Workbook | Worksheet> {
  if (isReadTarget(first)) {
    return readMarkdownInto(
      first,
      await markdownText("ExcelMarkdown.readFile", second as string),
      third,
      false,
      "ExcelMarkdown.readFile"
    );
  }
  return readNewWorkbook(
    "ExcelMarkdown.readFile",
    await markdownText("ExcelMarkdown.readFile", first),
    second as MarkdownReadOptions | undefined
  );
}

/** Write a worksheet to a Markdown file path (Node.js only). */
export async function writeMarkdownFile(
  workbook: Workbook,
  filename: string,
  options?: MarkdownWriteOptions
): Promise<void> {
  checkFilePath("ExcelMarkdown.writeFile", filename);
  const markdownString = formatWorksheet("ExcelMarkdown.writeFile", workbook, options);
  await writeFileText(filename, markdownString);
}
