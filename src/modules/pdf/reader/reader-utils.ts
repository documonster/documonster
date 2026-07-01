/**
 * Shared utility functions for PDF reader modules.
 */

import type { PdfDocument } from "@pdf/reader/pdf-document";
import type { PdfDictValue } from "@pdf/reader/pdf-parser";
import { decodePdfStringBytes } from "@pdf/reader/pdf-parser";

/**
 * Safely extract a string value from a PDF dictionary entry.
 * Handles both name strings and Uint8Array PDF strings (with BOM/encoding detection).
 *
 * @param dict - The PDF dictionary
 * @param key - The key to look up
 * @param doc - The PDF document for resolving indirect references
 * @returns The string value, or empty string if not found or not a string
 */
export function getDictStringValue(dict: PdfDictValue, key: string, doc: PdfDocument): string {
  const val = dict.get(key);
  if (!val) {
    return "";
  }

  const resolved = doc.deref(val);
  if (typeof resolved === "string") {
    return resolved;
  }
  if (resolved instanceof Uint8Array) {
    return decodePdfStringBytes(resolved);
  }
  return "";
}
