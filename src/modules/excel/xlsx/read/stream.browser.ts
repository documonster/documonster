import type { Workbook } from "@excel/core/workbook.browser";
import { readXlsxBytesInto } from "@excel/xlsx/read/package";
import type { IParseStream, XlsxReadOptions } from "@excel/xlsx/types";
import { concatUint8Arrays } from "@utils/binary";

/**
 * Read a workbook from a parse stream — browser variant.
 *
 * There is no streaming ZIP parser on this platform, so the stream is collected and handed to the
 * byte reader. The Node variant (`read/stream.ts`) parses entries as they arrive instead.
 */
export async function readXlsxStreamInto(
  workbook: Workbook,
  stream: IParseStream,
  options?: XlsxReadOptions
): Promise<Workbook> {
  // Collect all stream data into a single buffer
  const chunks: Uint8Array[] = [];

  await new Promise<void>((resolve, reject) => {
    const onData = (chunk: Uint8Array) => {
      chunks.push(chunk);
    };

    const onEnd = () => {
      stream.off("data", onData);
      stream.off("end", onEnd);
      stream.off("error", onError);
      resolve();
    };

    const onError = (err: Error) => {
      stream.off("data", onData);
      stream.off("end", onEnd);
      stream.off("error", onError);
      reject(err);
    };

    stream.on("data", onData);
    stream.on("end", onEnd);
    stream.on("error", onError);
  });

  return readXlsxBytesInto(workbook, concatUint8Arrays(chunks), options);
}
