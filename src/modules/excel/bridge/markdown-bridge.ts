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

import {
  DEFAULT_DATE_FORMATS,
  addInputRow,
  createDefaultValueMapper
} from "@excel/bridge/value-reading";
import { rowValues } from "@excel/core/row";
import type { RowData } from "@excel/core/row";
import { addWorksheet, createWorkbook, getWorksheet } from "@excel/core/workbook";
import { refuseLongSheetName, validateSheetName } from "@excel/core/workbook-core";
import { isReadTarget } from "@excel/core/workbook-io-types";
import type { Workbook } from "@excel/core/workbook.browser";
import type { Worksheet } from "@excel/core/worksheet";
import { eachRow } from "@excel/core/worksheet";
import { ExcelArgumentError, ExcelWorksheetNameError } from "@excel/errors";
import type { CellValue, CellValueInput } from "@excel/types";
import { argumentFailure, checkSchema } from "@excel/utils/arguments";
import { checkFormatArguments } from "@markdown/arguments";
import { MarkdownArgumentError } from "@markdown/errors";
import { formatMarkdown } from "@markdown/format/index";
import { parseMarkdown, parseMarkdownAll } from "@markdown/parse/index";
import type {
  MarkdownOptions,
  MarkdownAlignment,
  MarkdownFormatOptions,
  MarkdownParseOptions,
  MarkdownParseResult
} from "@markdown/types";
import { createDateFormatter, createIsoDateFormatter } from "@utils/datetime";
import { sheetNameKey } from "@utils/name-key";
import { s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";

// =============================================================================
// Stringify
// =============================================================================

function createMarkdownStringify(
  dateFormat: string | undefined,
  dateUTC: boolean
): (value: unknown) => string {
  const formatter = dateFormat
    ? createDateFormatter(dateFormat, { utc: dateUTC })
    : createIsoDateFormatter(dateUTC);

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
      const v = value as {
        text?: string;
        hyperlink?: string;
        formula?: unknown;
        result?: unknown;
        richText?: { text: string }[];
        error?: string;
      };
      if (v.text || v.hyperlink) {
        return v.hyperlink || v.text || "";
      }
      if (v.formula || v.result) {
        return v.result != null ? String(v.result) : "";
      }
      if (v.richText && Array.isArray(v.richText)) {
        return v.richText.map(r => r.text).join("");
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
  map?: (text: string, column: number) => CellValueInput
): void {
  addInputRow(worksheet, result.headers, "Markdown");
  (worksheet as { _markdownAlignments?: MarkdownAlignment[] })._markdownAlignments =
    result.alignments;
  for (const row of result.rows) {
    if (map) {
      // Checked as `Cell.setValue` checks a value, by `addInputRow`.
      addInputRow(worksheet, row.map((v, i) => map(v, i)) as CellValue[], "Markdown");
    } else {
      addInputRow(worksheet, row, "Markdown");
    }
  }
}

/**
 * What a Markdown read takes: the text, or its UTF-8 bytes — what `ExcelMarkdown.toBuffer` and `toBufferSync`
 * produce. A byte-order mark is dropped.
 */
export type ExcelMarkdownInput = string | ArrayBufferView | ArrayBuffer;

/** Options for reading a Markdown table into a worksheet — `ExcelMarkdown.read`, `readFile`, `addWorksheet`, … */
export interface ExcelMarkdownReadOptions extends MarkdownParseOptions {
  /** The name of the worksheet the table is written to (`name`, `name_2`, … for several tables). */
  sheetName?: string;
  /**
   * Turn each data cell's text into the cell value written, in place of the default reading — numbers, ISO dates,
   * `true` / `false` and error literals, as `ExcelCsv` reads them. Called with the cell's text and its 0-based
   * column; what it returns is checked as `Cell.setValue` checks a value.
   */
  map?: (text: string, column: number) => CellValueInput;
  /**
   * Which tables `ExcelMarkdown.read` and `readFile` import: `"first"` (the default), or `"all"` — one sheet each,
   * named `sheetName`, `sheetName_2`, …. `addWorksheet` reads the first table and `addWorksheets` every one, as the
   * sheet or sheets they return say; given the other value, they refuse it.
   */
  tables?: "first" | "all";
}

/** Options for writing a worksheet as a Markdown table — `ExcelMarkdown.format`, `toBuffer`, `writeFile`, … */
export interface ExcelMarkdownWriteOptions extends MarkdownFormatOptions {
  /** The worksheet written, by name. Default: the first. */
  sheetName?: string;
  /** The worksheet written, by id, when no `sheetName` is given. */
  sheetId?: number;
  /**
   * The form dates are written in, as `YYYY-MM-DD HH:mm:ss` tokens. Default: ISO 8601 with milliseconds, in UTC
   * (`…Z`). A cell's number format is not applied — `Cell.getDisplayText` is the reading that applies it.
   */
  dateFormat?: string;
  /**
   * Take a date's fields from UTC — the calendar value the cell holds, and what Excel shows — rather than from the
   * host's timezone. Default `true`; `false` writes the host's local time. For the cell `2024-03-10 00:00` written
   * in New York: `2024-03-10T00:00:00.000Z` by default, `2024-03-09T19:00:00.000-05:00` with `dateUTC: false`; with
   * `dateFormat: "YYYY-MM-DD"`, `2024-03-10` and `2024-03-09`. As `ExcelCsv` writes them. (Before this release the
   * default was `false`.)
   */
  dateUTC?: boolean;
  /** Write an empty table row for each empty sheet row between rows that have values. Default `true`. */
  includeEmptyRows?: boolean;
}

/** The bridge's argument error: `where` is the function called (`"ExcelMarkdown.read"`), `field` the argument. */
function fail(where: string, field: string, expected: string, value: unknown): never {
  throw argumentFailure(where, field, expected, value);
}

/**
 * Run the Markdown parser or formatter, raising its argument errors as this bridge's. Every argument is checked here
 * or by the Markdown module before anything is read or written, and its `MarkdownArgumentError` named
 * `Markdown.parse`, which the caller never called, and was missed by `isExcelError`. Content the parser cannot read
 * is not an argument error and keeps the Markdown module's error.
 */
function asBridgeError<T>(where: string, run: () => T): T {
  try {
    return run();
  } catch (error) {
    if (error instanceof MarkdownArgumentError) {
      const detail = error.message.slice(error.message.indexOf(": ") + 2);
      throw new ExcelArgumentError(error.field, `${where}: ${detail}`, { cause: error });
    }
    throw error;
  }
}

// The bridge's own options. Those it hands to the Markdown parser or formatter are checked there (`asBridgeError`).
const readOptionsSchema = s.object({
  sheetName: s.string().optional(),
  map: s.function<(text: string, column: number) => CellValueInput>().optional(),
  tables: s.literal("first", "all").optional()
});
const _readOptionsMatch: SchemaMatches<
  typeof readOptionsSchema,
  Omit<ExcelMarkdownReadOptions, keyof MarkdownParseOptions>
> = true;
void _readOptionsMatch;

const writeOptionsSchema = s.object({
  sheetName: s.string().optional(),
  sheetId: s.number({ integer: true }).optional(),
  dateFormat: s.string().optional(),
  dateUTC: s.boolean().optional(),
  includeEmptyRows: s.boolean().optional()
});
const _writeOptionsMatch: SchemaMatches<
  typeof writeOptionsSchema,
  Omit<ExcelMarkdownWriteOptions, keyof MarkdownFormatOptions>
> = true;
void _writeOptionsMatch;

/**
 * The Markdown text of `input` — the text itself, or its bytes decoded as UTF-8, a byte-order mark dropped. Anything
 * else is refused with this bridge's error, as `ExcelCsv` refuses its input — the parser's own `MarkdownError` named
 * `Markdown.parse`, a function the caller never called. Bytes are recognised structurally: a typed array from another
 * realm fails `instanceof`.
 */
function markdownText(where: string, input: unknown): string {
  if (typeof input === "string") {
    return input;
  }
  if (ArrayBuffer.isView(input)) {
    return UTF8.decode(input);
  }
  if (Object.prototype.toString.call(input) === "[object ArrayBuffer]") {
    return UTF8.decode(input as ArrayBuffer);
  }
  return fail(where, "input", "Markdown text or its UTF-8 bytes", input);
}

const UTF8 = new TextDecoder("utf-8");

/** Refuse a `tables` that is not one of its two values, or that contradicts what the call returns. */
function checkTables(
  where: string,
  options: ExcelMarkdownReadOptions | MarkdownOptions | undefined,
  only?: "first" | "all"
): void {
  const tables = (options as ExcelMarkdownReadOptions | undefined)?.tables;
  if (tables === undefined) {
    return;
  }
  if (only !== undefined && tables !== only) {
    throw new ExcelArgumentError(
      "options.tables",
      `${where}: options.tables must be "${only}" here, got "${tables}" — ${
        only === "first"
          ? "addWorksheet reads one table into one sheet; addWorksheets reads every table"
          : "addWorksheets reads every table; addWorksheet reads the first"
      }.`
    );
  }
}

/**
 * The parse options and value reading of a read. `typed` is the reading of the `ExcelMarkdown` namespace: text that
 * reads as a number, date, boolean or error becomes one — as `ExcelCsv` reads it — and the `<br>` the formatter
 * writes for a newline is a newline again, so a workbook formatted and read back keeps its values. The deprecated
 * flat functions read text only, as they always did.
 */
function readSettings(
  where: string,
  options: ExcelMarkdownReadOptions | MarkdownOptions | undefined,
  typed: boolean
): { parse: MarkdownParseOptions; map?: (text: string, column: number) => CellValueInput } {
  checkSchema(where, "options", readOptionsSchema.optional(), options);
  checkTables(where, options);
  const parse: MarkdownParseOptions = {
    trim: options?.trim,
    unescape: options?.unescape,
    skipEmptyRows: options?.skipEmptyRows,
    maxRows: options?.maxRows,
    convertBr: options?.convertBr ?? (typed ? true : undefined)
  };
  if (options?.map || !typed) {
    // The deprecated flat forms' `MarkdownOptions.map` returns `unknown`; `addInputRow` checks what it returns either way.
    return {
      parse,
      map: options?.map as ((text: string, column: number) => CellValueInput) | undefined
    };
  }
  const read = createDefaultValueMapper(DEFAULT_DATE_FORMATS);
  return { parse, map: value => read(value) };
}

/** {@link readMarkdown}, reading values as the `ExcelMarkdown` namespace does when `typed`. */
export function readMarkdownInto(
  workbook: Workbook,
  input: ExcelMarkdownInput,
  options: ExcelMarkdownReadOptions | MarkdownOptions | undefined,
  typed: boolean,
  where = "ExcelMarkdown.addWorksheet"
): Worksheet {
  if (!isReadTarget(workbook)) {
    fail(where, "workbook", "a workbook", workbook);
  }
  const text = markdownText(where, input);
  const { parse, map } = readSettings(where, options, typed);
  refuseLongSheetName(options?.sheetName);
  const parseResult = asBridgeError(where, () => parseMarkdown(text, parse));
  const worksheet = addWorksheet(workbook, options?.sheetName);
  populateMarkdownWorksheet(worksheet, parseResult, map);
  return worksheet;
}

/**
 * The name of the `t`-th (0-based) table's sheet: `name`, then `name_2`, `name_3`, … — with `name` shortened so the
 * suffix fits Excel's 31 characters. The caller's own name is theirs and is refused past 31; a name this bridge makes
 * up is its own to keep valid, so a 31-character `name` gave `name_2`, 33 characters, and a refusal of a name the
 * caller never wrote.
 */
function derivedSheetName(name: string, t: number): string {
  if (t === 0) {
    return name;
  }
  const suffix = `_${t + 1}`;
  return `${name.slice(0, 31 - suffix.length)}${suffix}`;
}

/** {@link readMarkdownAll}, reading values as the `ExcelMarkdown` namespace does when `typed`. */
export function readMarkdownAllInto(
  workbook: Workbook,
  input: ExcelMarkdownInput,
  options: ExcelMarkdownReadOptions | MarkdownOptions | undefined,
  typed: boolean,
  where = "ExcelMarkdown.addWorksheets"
): Worksheet[] {
  if (!isReadTarget(workbook)) {
    fail(where, "workbook", "a workbook", workbook);
  }
  const text = markdownText(where, input);
  const { parse, map } = readSettings(where, options, typed);
  const parseResults = asBridgeError(where, () => parseMarkdownAll(text, parse));
  const baseName = options?.sheetName;
  const worksheets: Worksheet[] = [];
  refuseLongSheetName(baseName);
  const names = parseResults.map((_, t) => (baseName ? derivedSheetName(baseName, t) : undefined));
  // Every name is checked — against the workbook and against each other — before the first sheet is added, so a
  // refused read adds none. A clash found while adding left the sheets before it behind.
  const seen = new Set<string>();
  for (const name of names) {
    if (name !== undefined) {
      validateSheetName(workbook, name);
      const key = sheetNameKey(name);
      if (seen.has(key)) {
        throw new ExcelWorksheetNameError(`Sheet name already exists: ${name}`);
      }
      seen.add(key);
    }
  }
  for (let t = 0; t < parseResults.length; t++) {
    const name = names[t];
    const worksheet = addWorksheet(workbook, name);
    populateMarkdownWorksheet(worksheet, parseResults[t], map);
    worksheets.push(worksheet);
  }
  return worksheets;
}

/**
 * Read the first Markdown table in `input` into a new worksheet on `workbook`, and return it. The header row stays
 * text; each data cell is read as `ExcelCsv` reads one — numbers, ISO dates, `true` / `false`, error literals — and
 * `<br>` becomes a newline, unless `options.map` or `options.convertBr` says otherwise. A date with no timezone
 * (`2024-03-10`) is the calendar value in UTC, as a cell holds it.
 *
 * @throws {ExcelArgumentError} for an argument or option this bridge cannot use, naming the function called.
 *
 * @example
 * ```ts
 * import { ExcelMarkdown } from "documonster/excel/markdown";
 * const worksheet = ExcelMarkdown.addWorksheet(workbook, "| Name | Age |\n| --- | --- |\n| Alice | 30 |");
 * ```
 */
export function addMarkdownWorksheet(
  workbook: Workbook,
  input: ExcelMarkdownInput,
  options?: ExcelMarkdownReadOptions
): Worksheet {
  checkSchema("ExcelMarkdown.addWorksheet", "options", readOptionsSchema.optional(), options);
  checkTables("ExcelMarkdown.addWorksheet", options, "first");
  return readMarkdownInto(workbook, input, options, true);
}

/** Read every Markdown table in `input`, one new worksheet each, read as {@link addMarkdownWorksheet} reads one. */
export function addMarkdownWorksheets(
  workbook: Workbook,
  input: ExcelMarkdownInput,
  options?: ExcelMarkdownReadOptions
): Worksheet[] {
  checkSchema("ExcelMarkdown.addWorksheets", "options", readOptionsSchema.optional(), options);
  checkTables("ExcelMarkdown.addWorksheets", options, "all");
  return readMarkdownAllInto(workbook, input, options, true);
}

// =============================================================================
// Read
// =============================================================================

/** Read a Markdown table into a new worksheet on `workbook`, every cell as text (the deprecated flat form). */
export function readMarkdown(
  workbook: Workbook,
  input: string,
  options?: MarkdownOptions
): Worksheet {
  return readMarkdownInto(workbook, input, options, false, "readMarkdown");
}

/**
 * Read all Markdown tables from a document, each becoming a separate worksheet (named `sheetName`, `sheetName_2`,
 * ...), every cell as text (the deprecated flat form).
 */
export function readMarkdownAll(
  workbook: Workbook,
  input: string,
  options?: MarkdownOptions
): Worksheet[] {
  return readMarkdownAllInto(workbook, input, options, false, "readMarkdownAll");
}

/**
 * Read the first Markdown table in `input` — text, or its UTF-8 bytes — into a new workbook, returned; with
 * `{ tables: "all" }` every table, one worksheet each. Values are read as `ExcelMarkdown.addWorksheet` reads them.
 *
 * @example
 * ```ts
 * import { ExcelMarkdown } from "documonster/excel/markdown";
 * const workbook = ExcelMarkdown.read("| Name | Age |\n| --- | --- |\n| Alice | 30 |");
 * const everyTable = ExcelMarkdown.read(document, { tables: "all", sheetName: "Table" });
 * ```
 */
export function readMarkdownWorkbook(
  input: ExcelMarkdownInput,
  options?: ExcelMarkdownReadOptions
): Workbook;
/**
 * @deprecated Use `ExcelMarkdown.addWorksheet(workbook, input, options)`, which this form does — it adds a
 * worksheet to `workbook` and returns it. `ExcelMarkdown.read(input)` returns a new workbook. Removed in 1.0.
 */
export function readMarkdownWorkbook(
  workbook: Workbook,
  input: string,
  options?: MarkdownOptions
): Worksheet;
export function readMarkdownWorkbook(
  first: Workbook | ExcelMarkdownInput,
  second?: string | MarkdownOptions,
  third?: MarkdownOptions
): Workbook | Worksheet {
  if (isReadTarget(first)) {
    return readMarkdownInto(first, second as string, third, false, "ExcelMarkdown.read");
  }
  return readNewWorkbook(
    "ExcelMarkdown.read",
    first,
    second as ExcelMarkdownReadOptions | undefined
  );
}

/** A new workbook holding the first table of `input`, or every one with `{ tables: "all" }`. */
export function readNewWorkbook(
  where: string,
  input: ExcelMarkdownInput,
  options: ExcelMarkdownReadOptions | undefined
): Workbook {
  const workbook = createWorkbook();
  if (isRecord(options) && options.tables === "all") {
    readMarkdownAllInto(workbook, input, options, true, where);
  } else {
    readMarkdownInto(workbook, input, options, true, where);
  }
  return workbook;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

// =============================================================================
// Write
// =============================================================================

/**
 * Write a worksheet as a Markdown table string: `sheetName`, else `sheetId`, else the first sheet; `""` for a
 * workbook with no sheets.
 *
 * A date cell is written as the calendar value it holds — ISO 8601 in UTC, or `dateFormat` filled from the UTC
 * fields — unless `dateUTC: false`; see {@link ExcelMarkdownWriteOptions.dateUTC}.
 *
 * @throws {ExcelArgumentError} for an argument or option this bridge cannot use — the message names the function
 * called and `field` the option — including a worksheet name or id the workbook does not have (that used to write
 * `""`, so a misspelt name produced an empty table without a word).
 */
export function writeMarkdown(workbook: Workbook, options?: ExcelMarkdownWriteOptions): string {
  return formatWorksheet("ExcelMarkdown.format", workbook, options);
}

/** @internal {@link writeMarkdown}, with the public name the errors give — for the Node file variant too. */
export function formatWorksheet(
  where: string,
  workbook: Workbook,
  options?: ExcelMarkdownWriteOptions
): string {
  if (!isReadTarget(workbook)) {
    fail(where, "workbook", "a workbook", workbook);
  }
  checkSchema(where, "options", writeOptionsSchema.optional(), options);
  const sheetName = options?.sheetName;
  const sheetId = options?.sheetId;
  // The formatter's own checks of its options, run before the sheet is read so an empty sheet is checked too.
  const formatOptions = {
    columns: options?.columns,
    alignment: options?.alignment,
    padding: options?.padding,
    trailingNewline: options?.trailingNewline,
    escapeContent: options?.escapeContent,
    stringify: options?.stringify
  };
  asBridgeError(where, () => checkFormatArguments([], [], formatOptions));
  const key = sheetName || sheetId;
  const worksheet = getWorksheet(workbook, key);
  if (!worksheet) {
    if (key !== undefined) {
      throw new ExcelArgumentError(
        sheetName ? "options.sheetName" : "options.sheetId",
        `${where}: the workbook has no worksheet ${JSON.stringify(key)}.`
      );
    }
    return "";
  }

  const dateFormat = options?.dateFormat;
  const dateUTC = options?.dateUTC ?? true;
  const includeEmptyRows = options?.includeEmptyRows !== false;

  const stringify = options?.stringify ?? createMarkdownStringify(dateFormat, dateUTC);

  const allRows: unknown[][] = [];
  let lastRow = 1;

  eachRow(worksheet, (row: RowData, rowNumber: number) => {
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

  const storedAlignments: MarkdownAlignment[] | undefined = (
    worksheet as { _markdownAlignments?: MarkdownAlignment[] }
  )._markdownAlignments;

  const columns = options?.columns;
  let resolvedColumns: { header: string; alignment?: MarkdownAlignment }[] | undefined;

  if (!columns && storedAlignments) {
    resolvedColumns = headers.map((h, i) => ({
      header: h,
      alignment: i < storedAlignments.length ? storedAlignments[i] : undefined
    }));
  }

  return asBridgeError(where, () =>
    formatMarkdown(headers, dataRows, {
      columns: resolvedColumns ?? columns,
      alignment: options?.alignment,
      padding: options?.padding,
      trailingNewline: options?.trailingNewline,
      escapeContent: options?.escapeContent,
      stringify
    })
  );
}

/**
 * Write a worksheet to a Markdown buffer (`Uint8Array`) synchronously. Published as
 * `ExcelMarkdown.toBufferSync`; `ExcelMarkdown.toBuffer` is the asynchronous form.
 */
export function writeMarkdownBuffer(
  workbook: Workbook,
  options?: ExcelMarkdownWriteOptions
): Uint8Array {
  return new TextEncoder().encode(formatWorksheet("ExcelMarkdown.toBufferSync", workbook, options));
}

/**
 * Write a worksheet as Markdown to a buffer (`Uint8Array`).
 *
 * Asynchronous like every other `toBuffer` in the excel entries
 * (`Workbook.toBuffer`, `ExcelCsv.toBuffer`), so one verb has one shape.
 */
export async function markdownToBuffer(
  workbook: Workbook,
  options?: ExcelMarkdownWriteOptions
): Promise<Uint8Array> {
  return new TextEncoder().encode(formatWorksheet("ExcelMarkdown.toBuffer", workbook, options));
}
