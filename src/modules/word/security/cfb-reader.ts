/**
 * DOCX Module - CFB (Compound File Binary) container, as `WordCrypto.Cfb` publishes it.
 *
 * The reader and writer live in `@archive/cfb/cfb`, shared with Excel's encrypted workbooks. This module checks the
 * arguments with Word's checks and reports a malformed container as `DocxParseError`, as it always has.
 */

import { CfbError, readCfb as readSharedCfb, writeCfb as writeSharedCfb } from "@archive/cfb/cfb";
import type { CfbEntry } from "@archive/cfb/cfb";
import {
  checkBytes,
  checkNodeArray,
  checkString,
  checkStringArray
} from "@word/core/argument-checks";
import { DocxParseError } from "@word/errors";

export type { CfbEntry } from "@archive/cfb/cfb";

function docxCfbError(error: unknown): unknown {
  return error instanceof CfbError ? new DocxParseError(error.message, { cause: error }) : error;
}

/**
 * Read a CFB (OLE2 Compound File) and extract all stream entries. Reads v3 and v4 files (sector sizes 512 and 4096).
 *
 * @param buffer - The CFB file data.
 * @returns Array of named stream entries.
 * @throws {DocxParseError} If the bytes are not a readable compound file.
 */
export function readCfb(buffer: Uint8Array): CfbEntry[] {
  try {
    return readSharedCfb(buffer);
  } catch (error) {
    throw docxCfbError(error);
  }
}

/**
 * Write a set of named stream entries into a CFB (OLE2 Compound File) container: a v3 file with 512-byte sectors,
 * streams under 4096 bytes in the mini-stream as Office stores them. An entry's `path` nests it inside storages —
 * required for the `\u0006DataSpaces` structure Office demands in encrypted documents.
 *
 * @param entries - Named stream entries to include.
 * @returns The CFB file as a Uint8Array.
 */
export function writeCfb(entries: readonly CfbEntry[]): Uint8Array {
  const where = "WordCrypto.Cfb.write";
  checkNodeArray(entries, "entries", where);
  // Every entry is checked before any is written; its data is read as the bytes it views.
  const streams = entries.map((entry, i) => {
    checkString(entry.name, `entries[${i}].name`, where);
    checkStringArray(entry.path, `entries[${i}].path`, where, { optional: true });
    return { ...entry, data: checkBytes(entry.data, `entries[${i}].data`, where) };
  });
  try {
    return writeSharedCfb(streams);
  } catch (error) {
    throw docxCfbError(error);
  }
}
