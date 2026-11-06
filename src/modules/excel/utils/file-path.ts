/**
 * The checks and errors every `readFile` / `writeFile` of the Excel entries shares: a path is a non-empty string,
 * refused at the call with the argument error (a number reached the file system as an untyped `TypeError`), and a
 * file the system would not let us read is an `ExcelFileError` that keeps the system's error as its `cause`.
 */
import { ExcelArgumentError, ExcelFileError } from "@excel/errors";

/** Refuse a path that is not a non-empty string, naming the call and the argument. */
export function checkFilePath(where: string, filename: unknown, argument = "filename"): void {
  if (typeof filename !== "string" || filename === "") {
    const shown =
      typeof filename === "string"
        ? '""'
        : filename === null || typeof filename !== "object"
          ? String(filename)
          : "an object";
    throw new ExcelArgumentError(
      argument,
      `${where}: ${argument} must be a non-empty file path, got ${shown}.`
    );
  }
}

/**
 * The `ExcelFileError` for a file the system would not let us read or write, with its error as the cause. `fileKind`
 * names the file as the message reads it: a CSV file is not an "Excel file".
 */
export function fileSystemError(
  filename: string,
  operation: "read" | "write",
  error: unknown,
  fileKind?: string
): ExcelFileError {
  const code = (error as { code?: unknown } | null)?.code;
  const detail =
    code === "ENOENT" ? "File not found" : error instanceof Error ? error.message : String(error);
  return new ExcelFileError(filename, operation, detail, { cause: error, fileKind });
}
