/**
 * How every Word reader — `Document.read`, `Stream.DocumentReader`, the `Package` functions — reports a read that
 * failed below the document model, so they report it the same way: a package that is damaged or not one is a
 * `DocxPackageError` whose `reason` says which, with the archive's diagnosis as the cause; anything else that is not
 * already a Word error is a `DocxParseError` naming it.
 */
import { ZipParseError } from "@archive/core/errors";
import { isArchiveDataError, startsLikeZip } from "@archive/core/read-failure";
import { DocxError, DocxPackageError, DocxParseError } from "@word/errors";

/** The Word error for `error`, raised reading the package `input` (the bytes after any decryption). */
export function packageReadError(error: unknown, input: Uint8Array): DocxError {
  if (error instanceof DocxError) {
    return error;
  }
  const message = error instanceof Error ? error.message : String(error);
  if (error instanceof ZipParseError && !startsLikeZip(input)) {
    return new DocxPackageError("not-a-document", `Not a DOCX package — ${message}`, {
      cause: error
    });
  }
  if (isArchiveDataError(error)) {
    return new DocxPackageError("damaged", `The DOCX package is damaged — ${message}`, {
      cause: error
    });
  }
  return new DocxParseError(`Failed to read DOCX: ${message}`, { cause: error });
}

/**
 * The error for a main document part whose root, `rootName`, is not Word's `w:document`: the package is another kind
 * of Office document — a workbook, a presentation — or its main part is in a namespace Word does not read.
 */
export function notWordDocumentError(rootName: string): DocxPackageError {
  return new DocxPackageError(
    "not-a-document",
    `Not a Word document: the package's main part has the root <${rootName}>, not a <w:document> in Word's ` +
      "namespace — it may be a workbook or a presentation"
  );
}
