/**
 * Browser variant: there is no filesystem, so an image named by `filename` cannot be written.
 */
import { ExcelNotSupportedError } from "@excel/errors";

export function readMediaFile(_filename: string): Promise<Uint8Array> {
  return Promise.reject(
    new ExcelNotSupportedError("Loading images from filename", "not supported in this environment")
  );
}
