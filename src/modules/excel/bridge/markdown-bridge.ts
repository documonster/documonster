/**
 * Markdown ↔ Workbook bridge — free functions.
 *
 * Tree-shakeable Markdown-table import/export as free functions taking a
 * `Workbook` handle. Consumers who never import this module pay nothing for
 * the markdown parser/formatter — the core `Workbook` no longer references
 * `@markdown`.
 *
 * Layer note: lives inside the excel module (layer 4), so it may import from
 * `@markdown` (layer 1). Node-only file-path variants live in
 * `./markdown-bridge.node.ts`.
 */

import { rowValues } from "@excel/core/row";
import { addWorksheet, getWorksheet } from "@excel/core/workbook";
import type { Workbook } from "@excel/core/workbook.browser";
import type { Worksheet } from "@excel/core/worksheet";
import { addRow, eachRow } from "@excel/core/worksheet";
import { formatMarkdown } from "@markdown/format/index";
import { parseMarkdown, parseMarkdownAll } from "@markdown/parse/index";
import type { MarkdownOptions, MarkdownAlignment, MarkdownParseResult } from "@markdown/types";
import { DateFormatter } from "@utils/datetime";

// =============================================================================
// Stringify
// =============================================================================

function createMarkdownStringify(
  dateFormat?: string,
  dateUTC?: boolean
): (value: unknown) => string {
  const formatter = dateFormat
    ? DateFormatter.create(dateFormat, { utc: dateUTC })
    : DateFormatter.iso(dateUTC);

  return function stringify(value: unknown): string {
    if (value === null || value === undefined) {
      return "";
    }
    if (typeof value === "string") {
      return value;
    }
    if (typeof value === "number" || typeof value === "bigint") {
      return String(value);
    }
    if (typeof value === "boolean") {
      return value ? "true" : "false";
    }
    if (value instanceof Date) {
      return formatter.format(value);
    }
    if (typeof value === "object") {
      const v = value as any;
      if (v.text || v.hyperlink) {
        return v.hyperlink || v.text || "";
      }
      if (v.formula || v.result) {
        return v.result != null ? String(v.result) : "";
      }
      if (v.richText && Array.isArray(v.richText)) {
        return v.richText.map((r: { text: string }) => r.text).join("");
      }
      if (v.error) {
        return v.error;
      }
      try {
        return JSON.stringify(value);
      } catch {
        return "[object Object]";
      }
    }
    return String(value);
  };
}

function populateMarkdownWorksheet(
  worksheet: Worksheet,
  result: MarkdownParseResult,
  map?: (value: string, column: number) => unknown
): void {
  addRow(worksheet, result.headers);
  (worksheet as any)._markdownAlignments = result.alignments;
  for (const row of result.rows) {
    if (map) {
      addRow(
        worksheet,
        row.map((v, i) => map(v, i))
      );
    } else {
      addRow(worksheet, row);
    }
  }
}

// =============================================================================
// Read
// =============================================================================

/**
 * Read a Markdown table into a new worksheet on `workbook`.
 *
 * @example
 * ```ts
 * import { readMarkdown } from "documonster/excel/markdown";
 * readMarkdown(workbook, "| Name | Age |\n| --- | --- |\n| Alice | 30 |");
 * ```
 */
export function readMarkdown(
  workbook: Workbook,
  input: string,
  options?: MarkdownOptions
): Worksheet {
  const parseResult = parseMarkdown(input, {
    trim: options?.trim,
    unescape: options?.unescape,
    skipEmptyRows: options?.skipEmptyRows,
    maxRows: options?.maxRows,
    convertBr: options?.convertBr
  });

  const worksheet = addWorksheet(workbook, options?.sheetName);
  populateMarkdownWorksheet(worksheet, parseResult, options?.map);
  return worksheet;
}

/**
 * Read all Markdown tables from a document, each becoming a separate
 * worksheet (named `sheetName`, `sheetName_2`, ...).
 */
export function readMarkdownAll(
  workbook: Workbook,
  input: string,
  options?: MarkdownOptions
): Worksheet[] {
  const parseResults = parseMarkdownAll(input, {
    trim: options?.trim,
    unescape: options?.unescape,
    skipEmptyRows: options?.skipEmptyRows,
    maxRows: options?.maxRows,
    convertBr: options?.convertBr
  });

  const baseName = options?.sheetName;
  const map = options?.map;
  const worksheets: Worksheet[] = [];

  for (let t = 0; t < parseResults.length; t++) {
    const name = baseName ? (t === 0 ? baseName : `${baseName}_${t + 1}`) : undefined;
    const worksheet = addWorksheet(workbook, name);
    populateMarkdownWorksheet(worksheet, parseResults[t], map);
    worksheets.push(worksheet);
  }

  return worksheets;
}

// =============================================================================
// Write
// =============================================================================

/** Write a worksheet as a Markdown table string. */
export function writeMarkdown(workbook: Workbook, options?: MarkdownOptions): string {
  const worksheet = getWorksheet(workbook, options?.sheetName || options?.sheetId);
  if (!worksheet) {
    return "";
  }

  const dateFormat = options?.dateFormat;
  const dateUTC = options?.dateUTC;
  const includeEmptyRows = options?.includeEmptyRows !== false;

  const stringify = options?.stringify ?? createMarkdownStringify(dateFormat, dateUTC);

  const allRows: unknown[][] = [];
  let lastRow = 1;

  eachRow(worksheet, (row: any, rowNumber: number) => {
    if (includeEmptyRows) {
      while (lastRow++ < rowNumber - 1) {
        allRows.push([]);
      }
    }
    const values = Array.from(rowValues(row) as unknown[]).slice(1);
    allRows.push(values);
    lastRow = rowNumber;
  });

  if (allRows.length === 0) {
    return "";
  }

  const headerRow = allRows[0];
  const headers: string[] = headerRow.map(v => stringify(v));
  const dataRows = allRows.slice(1);

  const storedAlignments: MarkdownAlignment[] | undefined = (worksheet as any)._markdownAlignments;

  const columns = options?.columns;
  let resolvedColumns: { header: string; alignment?: MarkdownAlignment }[] | undefined;

  if (!columns && storedAlignments) {
    resolvedColumns = headers.map((h, i) => ({
      header: h,
      alignment: i < storedAlignments.length ? storedAlignments[i] : undefined
    }));
  }

  return formatMarkdown(headers, dataRows, {
    columns: resolvedColumns ?? columns,
    alignment: options?.alignment,
    padding: options?.padding,
    trailingNewline: options?.trailingNewline,
    escapeContent: options?.escapeContent,
    stringify
  });
}

/** Write a worksheet to a Markdown buffer (`Uint8Array`). */
export function writeMarkdownBuffer(workbook: Workbook, options?: MarkdownOptions): Uint8Array {
  return new TextEncoder().encode(writeMarkdown(workbook, options));
}
