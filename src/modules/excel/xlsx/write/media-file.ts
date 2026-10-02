/**
 * Reads an image a workbook names by `filename` rather than carrying as bytes.
 *
 * A platform variant rather than an injected callback: the class this replaced took a
 * `readFileAsync` hook that only its Node subclass set, so whether a filename image could be
 * written depended on which constructor had run. Resolution now decides it.
 */
import { readFileBytes } from "@utils/fs";

export function readMediaFile(filename: string): Promise<Uint8Array> {
  return readFileBytes(filename);
}
