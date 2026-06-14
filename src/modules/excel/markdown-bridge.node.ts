/**
 * Markdown ↔ Workbook bridge — Node-only file-path variants.
 *
 * Separated from `./markdown-bridge` so the browser bundle never pulls in
 * `@utils/fs`.
 */

import { ExcelFileError } from "@excel/errors";
import type { Workbook } from "@excel/workbook.browser";
import type { Worksheet } from "@excel/worksheet";
import type { MarkdownOptions } from "@markdown/types";
import { fileExists, readFileText, writeFileText } from "@utils/fs";

import { readMarkdown, readMarkdownAll, writeMarkdown } from "./markdown-bridge";

/** Read a Markdown table from a file path into a new worksheet (Node.js only). */
export async function readMarkdownFile(
  workbook: Workbook,
  filename: string,
  options?: MarkdownOptions
): Promise<Worksheet> {
  if (!(await fileExists(filename))) {
    throw new ExcelFileError(filename, "read", "file not found");
  }

  const content = await readFileText(filename);
  return readMarkdown(workbook, content, options);
}

/** Read all Markdown tables from a file path, each as a worksheet (Node.js only). */
export async function readMarkdownAllFile(
  workbook: Workbook,
  filename: string,
  options?: MarkdownOptions
): Promise<Worksheet[]> {
  if (!(await fileExists(filename))) {
    throw new ExcelFileError(filename, "read", "file not found");
  }

  const content = await readFileText(filename);
  return readMarkdownAll(workbook, content, options);
}

/** Write a worksheet to a Markdown file path (Node.js only). */
export async function writeMarkdownFile(
  workbook: Workbook,
  filename: string,
  options?: MarkdownOptions
): Promise<void> {
  const markdownString = writeMarkdown(workbook, options);
  await writeFileText(filename, markdownString);
}
