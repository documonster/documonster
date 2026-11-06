/**
 * Which of the two containers a package is.
 *
 * A leaf on purpose, and the reason is the same one that put `core/unsupported.ts` beside it: the
 * answer is needed **synchronously** — `detectFormat` is not async and cannot be, because the
 * canonical `read`/`toBuffer` decide the format before they do anything else — while the reader
 * that acts on the answer is loaded on demand. Asking `xlsb/read/package.ts` for it, which is where
 * this lived, meant a static import of the whole XLSB reader from `core/workbook-format.ts`: every
 * consumer of `Workbook.read` paid for the binary reader in order to find out it was not needed,
 * and the `await import()` the streaming reader uses for the same modules could never take effect.
 */

import { ArchiveLimitExceededError } from "@archive/core/errors";
import { ZipParser } from "@archive/unzip/zip-parser";
import { containerLimitError } from "@excel/utils/container-errors";

/**
 * The workbook part that distinguishes the containers.
 *
 * XLSB stores the workbook as BIFF12 records in `xl/workbook.bin`; XLSX stores it as XML in
 * `xl/workbook.xml`. Everything else about the two packages — the ZIP, the content types, the
 * relationship graph — is the same shape, which is why presence of this part is the test.
 */
export const XLSB_WORKBOOK_PART = "xl/workbook.bin";

/**
 * Whether these bytes are an XLSB package.
 *
 * Reads the ZIP central directory only, so it is O(entries) rather than O(bytes), and answers
 * `false` for anything it cannot parse — a caller that wants a diagnosis gets a better one from the
 * loader that then tries to read the file than from a sniffer that declined to.
 *
 * `maxEntries` is the read's own bound, applied here because this is the first walk of the directory: checked only
 * by the loader afterwards, the records it exists to keep from being allocated already had been. A package past it
 * is the read's `ExcelLimitExceededError`, not "not XLSB".
 */
export function isXlsbPackage(bytes: Uint8Array, maxEntries?: number): boolean {
  try {
    return new ZipParser(bytes, maxEntries === undefined ? {} : { maxEntries })
      .getEntries()
      .some(entry => entry.path.toLowerCase() === XLSB_WORKBOOK_PART);
  } catch (error) {
    if (error instanceof ArchiveLimitExceededError) {
      throw containerLimitError(error);
    }
    return false;
  }
}
