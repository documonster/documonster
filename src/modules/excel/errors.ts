/**
 * Excel module error types.
 *
 * All Excel-related errors extend ExcelError.
 */

import type { BaseErrorOptions } from "@utils/errors";
import { BaseError } from "@utils/errors";

/**
 * Base class for all Excel-related errors.
 */
export class ExcelError extends BaseError {
  constructor(message: string, options?: BaseErrorOptions) {
    super(message, options);
    this.name = "ExcelError";
  }
}

/**
 * Check if an error is an Excel error.
 */
export function isExcelError(err: unknown): err is ExcelError {
  return err instanceof ExcelError;
}

/**
 * Error thrown when worksheet name validation fails.
 */
export class ExcelWorksheetNameError extends ExcelError {
  override name = "ExcelWorksheetNameError";
}

/**
 * Error thrown when trying to merge already merged cells.
 */
export class ExcelMergeConflictError extends ExcelError {
  override name = "ExcelMergeConflictError";

  constructor(details?: string, options?: BaseErrorOptions) {
    super(details ?? "Cannot merge already merged cells", options);
  }
}

/**
 * Error thrown when an argument has no meaning for the call it was passed to — a negative row height, a column width
 * that is not a number, an image id the workbook does not have — naming the function and the argument.
 *
 * `field` is the path from the parameter name in the public signature: `height`, `options.zip.level`,
 * `values[2]`, `font.size`. The message starts with the public call as the caller wrote it (`Row.setHeight: …`).
 */
export class ExcelArgumentError extends ExcelError {
  override name = "ExcelArgumentError";

  constructor(
    public readonly field: string,
    message: string,
    options?: BaseErrorOptions
  ) {
    super(message, options);
  }
}

/**
 * Options for the position errors — {@link ExcelInvalidAddressError}, {@link ExcelColumnOutOfBoundsError} and
 * {@link ExcelRowOutOfBoundsError}: the base options, the argument path the position was given at, and the public call.
 */
export interface ExcelPositionErrorOptions extends BaseErrorOptions {
  /**
   * The path of the refused position from the parameter name in the public signature — `addr`, `row`, `col`,
   * `cell.r`. Defaults to `addr`, `row` and `col` respectively.
   */
  field?: string;
  /** The public call as the caller wrote it (`Address.encodeCol`), which the message then starts with. */
  where?: string;
}

function positionMessage(text: string, options: ExcelPositionErrorOptions | undefined): string {
  return options?.where ? `${options.where}: ${text}` : text;
}

function positionCause(
  options: ExcelPositionErrorOptions | undefined
): BaseErrorOptions | undefined {
  return options?.cause === undefined ? undefined : { cause: options.cause };
}

/**
 * Error thrown when a cell address or range is invalid.
 *
 * An {@link ExcelArgumentError} (since this release), so the `catch` a caller writes for bad input catches it;
 * `field` names the parameter the address was given as (`addr` unless the call says otherwise).
 */
export class ExcelInvalidAddressError extends ExcelArgumentError {
  override name = "ExcelInvalidAddressError";

  constructor(
    public readonly address: string,
    details?: string,
    options?: ExcelPositionErrorOptions
  ) {
    super(
      options?.field ?? "addr",
      positionMessage(
        details ? `Invalid address "${address}": ${details}` : `Invalid address: ${address}`,
        options
      ),
      positionCause(options)
    );
  }
}

/**
 * Error thrown when a column number or letter is out of bounds.
 *
 * An {@link ExcelArgumentError} (since this release); `field` names the parameter (`col` unless the call says
 * otherwise).
 */
export class ExcelColumnOutOfBoundsError extends ExcelArgumentError {
  override name = "ExcelColumnOutOfBoundsError";

  constructor(
    public readonly column: number | string,
    details?: string,
    options?: ExcelPositionErrorOptions
  ) {
    super(
      options?.field ?? "col",
      positionMessage(
        details
          ? `Column ${column} is out of bounds: ${details}`
          : `Column ${column} is out of bounds. Excel supports columns from 1 to 16384`,
        options
      ),
      positionCause(options)
    );
  }
}

/**
 * Error thrown when a row number is out of bounds.
 *
 * An {@link ExcelArgumentError} (since this release); `field` names the parameter (`row` unless the call says
 * otherwise).
 */
export class ExcelRowOutOfBoundsError extends ExcelArgumentError {
  override name = "ExcelRowOutOfBoundsError";

  constructor(
    public readonly row: number,
    details?: string,
    options?: ExcelPositionErrorOptions
  ) {
    super(
      options?.field ?? "row",
      positionMessage(
        details ? `Row ${row} is out of bounds: ${details}` : `Row ${row} is out of bounds`,
        options
      ),
      positionCause(options)
    );
  }
}

/** Options for {@link ExcelInvalidValueTypeError}: the base options plus the argument path the value was given at. */
export interface ExcelInvalidValueTypeErrorOptions extends BaseErrorOptions {
  /** The path of the refused value from the parameter name (`values[2]`, `value.richText[0].text`). Default `value`. */
  field?: string;
}

/**
 * Error thrown when a cell value cannot be held by a workbook — `NaN`, a function, an invalid date, or an object of a
 * known value shape that is malformed (`{ richText: 42 }`, `{ error: "#BOGUS" }`).
 *
 * It is an {@link ExcelArgumentError}, so one `catch` handles every refused argument; `field` names where the value was
 * given (`values[2]`, `value.richText[0].text`).
 */
export class ExcelInvalidValueTypeError extends ExcelArgumentError {
  override name = "ExcelInvalidValueTypeError";

  constructor(
    public readonly valueType: string,
    details?: string,
    options?: ExcelInvalidValueTypeErrorOptions
  ) {
    super(
      options?.field ?? "value",
      details ?? `Cannot process value of type: ${valueType}`,
      options
    );
  }
}

/**
 * Error thrown when xlsx (OOXML) parsing encounters unexpected content.
 */
export class XlsxParseError extends ExcelError {
  override name = "XlsxParseError";

  constructor(
    public readonly context: string,
    details?: string,
    options?: BaseErrorOptions
  ) {
    super(details ?? `Unexpected XML content in ${context}`, options);
  }
}

/**
 * Error thrown when an operation is not supported.
 */
export class ExcelUnsupportedFeatureError extends ExcelError {
  override name = "ExcelUnsupportedFeatureError";

  /**
   * What could not be expressed or recovered, one entry per item.
   *
   * Present so a caller does not have to parse `message` to find out. The message truncates at ten
   * entries because an error is read by a human; this does not, because it is read by code — a
   * converter reporting "these 340 cells need attention" is a reasonable thing to build, and
   * scraping a sentence to build it is not.
   */
  readonly items: readonly string[];

  constructor(
    public readonly operation: string,
    public readonly reason?: string,
    options?: BaseErrorOptions & { readonly items?: readonly string[] }
  ) {
    super(reason ? `${operation}: ${reason}` : `${operation} is not supported`, options);
    this.items = options?.items ?? [];
  }
}

/**
 * Error thrown when a file operation fails.
 */
export class ExcelFileError extends ExcelError {
  override name = "ExcelFileError";

  constructor(
    public readonly path: string,
    public readonly operation: "read" | "write",
    details?: string,
    options?: ExcelFileErrorOptions
  ) {
    super(
      details
        ? `Failed to ${operation} ${options?.fileKind ?? "Excel"} file "${path}": ${details}`
        : `Failed to ${operation} ${options?.fileKind ?? "Excel"} file "${path}"`,
      options?.cause === undefined ? undefined : { cause: options.cause }
    );
  }
}

/**
 * Error thrown when the workbook read is encrypted (password-protected) and no `password` was given: an OLE compound
 * file holding the `EncryptionInfo` and `EncryptedPackage` streams Excel writes for a file with a password to open,
 * rather than a ZIP package. Read it again with its password — `Workbook.read(bytes, { password })`, and the same
 * option of `readFile`, `readStream`, `readWithDiagnostics` and `Stream.WorkbookReader` — as Word's `Document.read` takes it.
 * A wrong password is an {@link ExcelDecryptionError}.
 *
 * An {@link ExcelFileError}, so a `catch` for a file that cannot be read catches it too.
 */
export class ExcelEncryptedError extends ExcelFileError {
  override name = "ExcelEncryptedError";

  constructor(path = "<input>", options?: BaseErrorOptions) {
    super(
      path,
      "read",
      "the workbook is encrypted (password-protected) — pass its password as the read's `password` option",
      options
    );
  }
}

/**
 * Error thrown when an encrypted workbook cannot be decrypted with the `password` given: the password is wrong, the
 * file uses an encryption other than the Agile Encryption Excel writes (Standard Encryption from Excel 2007, RC4),
 * or its encryption data is damaged. `reason` says which — `"password"` means ask again. The counterpart of Word's
 * `DocxDecryptionError`, with the same `reason` values.
 *
 * An {@link ExcelFileError}, so a `catch` for a file that cannot be read catches it too.
 */
export class ExcelDecryptionError extends ExcelFileError {
  override name = "ExcelDecryptionError";
  /**
   * `"password"` — the password is wrong; `"unsupported"` — an encryption other than Agile; `"malformed"` — the
   * encryption metadata or package is damaged; `"integrity"` — the package fails its data-integrity check (altered or
   * damaged); `"limit"` — the package is larger than the read allows (`options.zip.maxTotalUncompressedSize`).
   */
  readonly reason: "password" | "unsupported" | "malformed" | "integrity" | "limit";

  constructor(
    path: string,
    reason: "password" | "unsupported" | "malformed" | "integrity" | "limit",
    details: string,
    options?: BaseErrorOptions
  ) {
    super(path, "read", details, options);
    this.reason = reason;
  }
}

/** Options for {@link ExcelFileError}: the base options plus what kind of file it was, for the message. */
export interface ExcelFileErrorOptions extends BaseErrorOptions {
  /** The kind of file, as the message names it — `"CSV"`, `"Markdown"`. Default `"Excel"`. */
  fileKind?: string;
}

/**
 * Error thrown when a streaming operation fails due to invalid state.
 */
export class ExcelStreamStateError extends ExcelError {
  override name = "ExcelStreamStateError";

  constructor(
    public readonly operation: string,
    public readonly state: string,
    options?: BaseErrorOptions
  ) {
    super(`Cannot ${operation}: ${state}`, options);
  }
}

/**
 * Error thrown when an HTTP download fails.
 */
export class ExcelDownloadError extends ExcelError {
  override name = "ExcelDownloadError";

  constructor(
    public readonly url: string,
    public readonly status: number,
    public readonly statusText: string,
    options?: BaseErrorOptions
  ) {
    super(`Failed to download from "${url}": HTTP ${status} ${statusText}`, options);
  }
}

/**
 * Error thrown when pivot table configuration is invalid.
 */
export class ExcelPivotTableError extends ExcelError {
  override name = "ExcelPivotTableError";
}

/**
 * Error thrown when chart options, a reference, a preset or a topology is not valid.
 *
 * An {@link ExcelArgumentError}, so it carries the `field` at fault (a path from the parameter
 * name, e.g. `options.series[0].values`) and every `catch` for either class, and
 * `isExcelError`, still answers.
 */
export class ExcelChartOptionsError extends ExcelArgumentError {
  override name = "ExcelChartOptionsError";

  constructor(field: string, message: string, options?: BaseErrorOptions);
  /**
   * @deprecated Pass the field first: `new ExcelChartOptionsError(field, message, options)`. This
   * form leaves `field` empty. Removed in 1.0.
   */
  constructor(message: string, options?: BaseErrorOptions);
  constructor(
    fieldOrMessage: string,
    messageOrOptions?: string | BaseErrorOptions,
    options?: BaseErrorOptions
  ) {
    const legacy = typeof messageOrOptions !== "string";
    super(
      legacy ? "" : fieldOrMessage,
      legacy ? fieldOrMessage : (messageOrOptions as string),
      legacy ? (messageOrOptions as BaseErrorOptions | undefined) : options
    );
  }
}

/**
 * Error thrown when table configuration or operation is invalid.
 */
export class ExcelTableError extends ExcelError {
  override name = "ExcelTableError";
}

/**
 * Error thrown when a defined name cannot be added — it is a table's name, and tables and defined names share one
 * namespace in Excel.
 */
export class ExcelDefinedNameError extends ExcelError {
  override name = "ExcelDefinedNameError";
}

/**
 * Error thrown when image processing fails.
 */
export class ExcelImageError extends ExcelError {
  override name = "ExcelImageError";
}

/**
 * A read option an {@link ExcelLimitExceededError} reports: `maxRows` / `maxCols` of the read, or one of the
 * container limits under `zip` (`maxEntrySize`, `maxTotalUncompressedSize`, `maxEntries`).
 */
export type ExcelLimit =
  | "maxRows"
  | "maxCols"
  | "maxEntrySize"
  | "maxTotalUncompressedSize"
  | "maxEntries";

/** Options for {@link ExcelLimitExceededError}. */
export interface ExcelLimitExceededErrorOptions extends BaseErrorOptions {
  /** The read option that set the bound; becomes {@link ExcelLimitExceededError.limit}. */
  limit?: ExcelLimit;
}

/**
 * Error thrown when a read exceeds a declared limit — `maxRows` / `maxCols`, or a container limit under `zip` — the
 * counterpart of `DocxLimitExceededError` and `PdfLimitExceededError`. `limit` names the option and `allowed` is its
 * value, as on the other modules' limit errors; a container limit keeps the archive's error as `cause`.
 */
export class ExcelLimitExceededError extends ExcelError {
  override name = "ExcelLimitExceededError";

  /** The read option that set the bound; `undefined` only on an error built without one. */
  readonly limit: ExcelLimit | undefined;

  /** The value of that option: the most items allowed. */
  readonly allowed: number;

  constructor(
    /** What was counted (`"row"`, `"column"`, `"part"`), as the message names it. */
    public readonly itemType: string,
    /** @deprecated Use {@link allowed}. Removed in 1.0. */
    public readonly maxItems: number,
    options?: ExcelLimitExceededErrorOptions
  ) {
    const { limit, ...errorOptions } = options ?? {};
    super(limitMessage(itemType, maxItems, limit), errorOptions);
    this.limit = limit;
    this.allowed = maxItems;
  }
}

function limitMessage(itemType: string, allowed: number, limit: ExcelLimit | undefined): string {
  if (limit === "maxEntrySize" || limit === "maxTotalUncompressedSize") {
    return `The ${itemType} exceeds ${limit} (${allowed} bytes)`;
  }
  if (limit === "maxEntries") {
    return `Max ${itemType} count (${allowed}) exceeded (maxEntries)`;
  }
  return `Max ${itemType} count (${allowed}) exceeded`;
}

/**
 * Error thrown when a BIFF12 (`.xlsb`) record stream is malformed.
 *
 * Carries the part it came from as well as the message, because a record stream
 * offset is meaningless without knowing which `.bin` it indexes into — and a
 * package has one per worksheet.
 */
export class XlsbParseError extends ExcelError {
  override name = "XlsbParseError";

  constructor(
    public readonly part: string,
    detail: string,
    options?: BaseErrorOptions
  ) {
    super(`${part}: ${detail}`, options);
  }
}

/**
 * Error thrown when a formula expression cannot be decoded and the caller asked for `formulas: "error"`.
 *
 * Carries the sheet and the addresses rather than only a count, because the useful next step is to look at those cells
 * — and a caller who decides the failure is acceptable can re-read with `"preserve"` and compare against this list.
 * Truncated in the message at ten addresses so a sheet that fails wholesale stays readable; `addresses` holds them all.
 */
export class XlsbFormulaDecodeError extends ExcelError {
  override name = "XlsbFormulaDecodeError";

  constructor(
    public readonly sheet: string,
    public readonly addresses: readonly string[],
    public readonly source: string,
    options?: BaseErrorOptions
  ) {
    const shown = addresses.slice(0, 10).join(", ");
    const rest = addresses.length > 10 ? `, and ${addresses.length - 10} more` : "";
    super(
      `${source}: ${addresses.length} formula expression(s) on sheet "${sheet}" could not be decoded (${shown}${rest})`,
      options
    );
  }
}

// --- Deprecated names (renamed to carry the module prefix). Removed in 1.0. ---

/** @deprecated Use {@link ExcelWorksheetNameError}. Removed in 1.0. */
export const WorksheetNameError = ExcelWorksheetNameError;
/** @deprecated Use {@link ExcelWorksheetNameError}. Removed in 1.0. */
export type WorksheetNameError = ExcelWorksheetNameError;
/** @deprecated Use {@link ExcelMergeConflictError}. Removed in 1.0. */
export const MergeConflictError = ExcelMergeConflictError;
/** @deprecated Use {@link ExcelMergeConflictError}. Removed in 1.0. */
export type MergeConflictError = ExcelMergeConflictError;
/** @deprecated Use {@link ExcelInvalidAddressError}. Removed in 1.0. */
export const InvalidAddressError = ExcelInvalidAddressError;
/** @deprecated Use {@link ExcelInvalidAddressError}. Removed in 1.0. */
export type InvalidAddressError = ExcelInvalidAddressError;
/** @deprecated Use {@link ExcelColumnOutOfBoundsError}. Removed in 1.0. */
export const ColumnOutOfBoundsError = ExcelColumnOutOfBoundsError;
/** @deprecated Use {@link ExcelColumnOutOfBoundsError}. Removed in 1.0. */
export type ColumnOutOfBoundsError = ExcelColumnOutOfBoundsError;
/** @deprecated Use {@link ExcelRowOutOfBoundsError}. Removed in 1.0. */
export const RowOutOfBoundsError = ExcelRowOutOfBoundsError;
/** @deprecated Use {@link ExcelRowOutOfBoundsError}. Removed in 1.0. */
export type RowOutOfBoundsError = ExcelRowOutOfBoundsError;
/** @deprecated Use {@link ExcelInvalidValueTypeError}. Removed in 1.0. */
export const InvalidValueTypeError = ExcelInvalidValueTypeError;
/** @deprecated Use {@link ExcelInvalidValueTypeError}. Removed in 1.0. */
export type InvalidValueTypeError = ExcelInvalidValueTypeError;
/** @deprecated Use {@link ExcelPivotTableError}. Removed in 1.0. */
export const PivotTableError = ExcelPivotTableError;
/** @deprecated Use {@link ExcelPivotTableError}. Removed in 1.0. */
export type PivotTableError = ExcelPivotTableError;
/** @deprecated Use {@link ExcelChartOptionsError}. Removed in 1.0. */
export const ChartOptionsError = ExcelChartOptionsError;
/** @deprecated Use {@link ExcelChartOptionsError}. Removed in 1.0. */
export type ChartOptionsError = ExcelChartOptionsError;
/** @deprecated Use {@link ExcelTableError}. Removed in 1.0. */
export const TableError = ExcelTableError;
/** @deprecated Use {@link ExcelTableError}. Removed in 1.0. */
export type TableError = ExcelTableError;
/** @deprecated Use {@link ExcelDefinedNameError}. Removed in 1.0. */
export const DefinedNameError = ExcelDefinedNameError;
/** @deprecated Use {@link ExcelDefinedNameError}. Removed in 1.0. */
export type DefinedNameError = ExcelDefinedNameError;
/** @deprecated Use {@link ExcelImageError}. Removed in 1.0. */
export const ImageError = ExcelImageError;
/** @deprecated Use {@link ExcelImageError}. Removed in 1.0. */
export type ImageError = ExcelImageError;
/** @deprecated Use {@link ExcelLimitExceededError}. Removed in 1.0. */
export const MaxItemsExceededError = ExcelLimitExceededError;
/** @deprecated Use {@link ExcelLimitExceededError}. Removed in 1.0. */
export type MaxItemsExceededError = ExcelLimitExceededError;
/**
 * @deprecated Use {@link ExcelUnsupportedFeatureError} — the same class, named as Word's `DocxUnsupportedFeatureError` is.
 * Removed in 1.0.
 */
export const ExcelNotSupportedError = ExcelUnsupportedFeatureError;
/** @deprecated Use {@link ExcelUnsupportedFeatureError}. Removed in 1.0. */
export type ExcelNotSupportedError = ExcelUnsupportedFeatureError;
/** @deprecated Use {@link ExcelPositionErrorOptions}. Removed in 1.0. */
export type PositionErrorOptions = ExcelPositionErrorOptions;
/** @deprecated Use {@link ExcelInvalidValueTypeErrorOptions}. Removed in 1.0. */
export type InvalidValueTypeErrorOptions = ExcelInvalidValueTypeErrorOptions;
