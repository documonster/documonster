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
export class WorksheetNameError extends ExcelError {
  override name = "WorksheetNameError";
}

/**
 * Error thrown when trying to merge already merged cells.
 */
export class MergeConflictError extends ExcelError {
  override name = "MergeConflictError";

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
 * Options for the position errors — {@link InvalidAddressError}, {@link ColumnOutOfBoundsError} and
 * {@link RowOutOfBoundsError}: the base options, the argument path the position was given at, and the public call.
 */
export interface PositionErrorOptions extends BaseErrorOptions {
  /**
   * The path of the refused position from the parameter name in the public signature — `addr`, `row`, `col`,
   * `cell.r`. Defaults to `addr`, `row` and `col` respectively.
   */
  field?: string;
  /** The public call as the caller wrote it (`Address.encodeCol`), which the message then starts with. */
  where?: string;
}

function positionMessage(text: string, options: PositionErrorOptions | undefined): string {
  return options?.where ? `${options.where}: ${text}` : text;
}

function positionCause(options: PositionErrorOptions | undefined): BaseErrorOptions | undefined {
  return options?.cause === undefined ? undefined : { cause: options.cause };
}

/**
 * Error thrown when a cell address or range is invalid.
 *
 * An {@link ExcelArgumentError} (since this release), so the `catch` a caller writes for bad input catches it;
 * `field` names the parameter the address was given as (`addr` unless the call says otherwise).
 */
export class InvalidAddressError extends ExcelArgumentError {
  override name = "InvalidAddressError";

  constructor(
    public readonly address: string,
    details?: string,
    options?: PositionErrorOptions
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
export class ColumnOutOfBoundsError extends ExcelArgumentError {
  override name = "ColumnOutOfBoundsError";

  constructor(
    public readonly column: number | string,
    details?: string,
    options?: PositionErrorOptions
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
export class RowOutOfBoundsError extends ExcelArgumentError {
  override name = "RowOutOfBoundsError";

  constructor(
    public readonly row: number,
    details?: string,
    options?: PositionErrorOptions
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

/** Options for {@link InvalidValueTypeError}: the base options plus the argument path the value was given at. */
export interface InvalidValueTypeErrorOptions extends BaseErrorOptions {
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
export class InvalidValueTypeError extends ExcelArgumentError {
  override name = "InvalidValueTypeError";

  constructor(
    public readonly valueType: string,
    details?: string,
    options?: InvalidValueTypeErrorOptions
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
export class ExcelNotSupportedError extends ExcelError {
  override name = "ExcelNotSupportedError";

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
 * option of `readFile`, `readStream`, `readWithDiagnostics` and `Stream.WorkbookReader` — as Word's `Io.read` takes it.
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
export class PivotTableError extends ExcelError {
  override name = "PivotTableError";
}

/**
 * Error thrown when chart options, a reference, a preset or a topology is not valid.
 *
 * An {@link ExcelArgumentError}, so it carries the `field` at fault (a path from the parameter
 * name, e.g. `options.series[0].values`) and every `catch` for either class, and
 * `isExcelError`, still answers.
 */
export class ChartOptionsError extends ExcelArgumentError {
  override name = "ChartOptionsError";

  constructor(field: string, message: string, options?: BaseErrorOptions);
  /**
   * @deprecated Pass the field first: `new ChartOptionsError(field, message, options)`. This
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
export class TableError extends ExcelError {
  override name = "TableError";
}

/**
 * Error thrown when a defined name cannot be added — it is a table's name, and tables and defined names share one
 * namespace in Excel.
 */
export class DefinedNameError extends ExcelError {
  override name = "DefinedNameError";
}

/**
 * Error thrown when image processing fails.
 */
export class ImageError extends ExcelError {
  override name = "ImageError";
}

/**
 * Error thrown when max items limit is exceeded.
 */
export class MaxItemsExceededError extends ExcelError {
  override name = "MaxItemsExceededError";

  constructor(
    public readonly itemType: string,
    public readonly maxItems: number,
    options?: BaseErrorOptions
  ) {
    super(`Max ${itemType} count (${maxItems}) exceeded`, options);
  }
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
