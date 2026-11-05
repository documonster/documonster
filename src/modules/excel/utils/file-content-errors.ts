/**
 * How a reader reports a position in a file's content that is outside the grid or not where it is written.
 *
 * The readers decode positions with the same checked helpers the API uses (`colCache`, `getCell`, `addRow`, the
 * `<row r>` / `<c r>` decoders), and those throw the position errors — `ExcelRowOutOfBoundsError`,
 * `ExcelColumnOutOfBoundsError`, `ExcelInvalidAddressError` — which are `ExcelArgumentError`s, because at an API call they are a
 * wrong argument. A file is not an argument: the same error from a read is translated where the reader knows what it
 * was reading — a part, a sheet's model, a row — into the error that reader uses for malformed content, the position
 * error kept as `cause` and its text in the message.
 *
 * The translation happens at those boundaries rather than in the helpers, so a helper shared by an API call and a
 * reader keeps meaning "wrong argument" when the caller wrote the position.
 */
import {
  ExcelColumnOutOfBoundsError,
  ExcelFileError,
  ExcelInvalidAddressError,
  ExcelRowOutOfBoundsError,
  XlsbParseError,
  XlsxParseError
} from "@excel/errors";

/** A position error: a row, a column or an address outside the grid, or not where it is written. */
export function isPositionError(
  error: unknown
): error is ExcelRowOutOfBoundsError | ExcelColumnOutOfBoundsError | ExcelInvalidAddressError {
  return (
    error instanceof ExcelRowOutOfBoundsError ||
    error instanceof ExcelColumnOutOfBoundsError ||
    error instanceof ExcelInvalidAddressError
  );
}

/** The error an XLSX reader reports for `error` raised reading `part`: an `XlsxParseError` for a position error. */
export function xlsxPartError(error: unknown, part: string): unknown {
  return isPositionError(error)
    ? new XlsxParseError(part, `${part}: ${error.message}`, { cause: error })
    : error;
}

/** The error an XLSB reader reports for `error` raised reading `part`: an `XlsbParseError` for a position error. */
export function xlsbPartError(error: unknown, part: string): unknown {
  return isPositionError(error) ? new XlsbParseError(part, error.message, { cause: error }) : error;
}

/**
 * The error a table reader (`ExcelCsv`, `ExcelMarkdown`) reports for `error` raised adding the `row`-th row of the
 * input (1-based, as a sheet numbers rows): an `ExcelFileError` naming the row for a position error — a row wider than
 * a sheet's 16 384 columns, or past its last row.
 */
export function tableRowError(error: unknown, row: number, fileKind: "CSV" | "Markdown"): unknown {
  return isPositionError(error)
    ? new ExcelFileError(
        "<input>",
        "read",
        `row ${row} does not fit a worksheet: ${error.message}`,
        {
          cause: error,
          fileKind
        }
      )
    : error;
}
