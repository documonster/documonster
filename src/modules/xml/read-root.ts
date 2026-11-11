/**
 * `Xml.readRoot`: the root element's start tag, and the prolog before it, read from the start of
 * a file and no further.
 */

import { checkByteSource, checkReadRootOptions, snapshotReadRootOptions } from "@xml/arguments";
import { ByteReader } from "@xml/byte-source";
import { XmlLimitExceededError, XmlParseError } from "@xml/errors";
import { SaxParser } from "@xml/sax";
import type { XmlByteSource, XmlReadRootOptions, XmlRootElement } from "@xml/types";

/**
 * Read a file's root element — its name and attributes — and the prolog before it, which
 * `Xml.readRecords` needs as `context` to read a range that starts later in the file.
 *
 * Reading stops at the root element's `>`: nothing after it is parsed, even in the same chunk, and
 * the source is closed. The prolog must be well-formed — any error before the root's `>` is thrown,
 * with its byte offset — and the DOCTYPE internal subset is read for its entities; nothing outside
 * the file is.
 *
 * The limits are on byte positions, not on how the source happens to be chunked, so a file gives
 * the same result however it arrives: the root's `<` must begin before `maxStartByte` and its start
 * tag be at most `maxTagBytes` long, or an {@link XmlLimitExceededError} is thrown; no more than
 * `maxStartByte + maxTagBytes` bytes are read, and up to three more to finish a character that
 * straddles that. A source that ends before the root's `>` — an empty
 * one included — is an {@link XmlTruncatedError}.
 *
 * @example
 * ```ts
 * const root = await Xml.readRoot(firstBytes);
 * root.name; // "DSR"
 * root.attributes["xmlns"];
 * ```
 */
async function readRoot(
  source: XmlByteSource,
  given?: XmlReadRootOptions
): Promise<XmlRootElement> {
  const where = "Xml.readRoot";
  const checkedSource = checkByteSource(where, source);
  // Read once, so what is checked is what is used.
  const options = snapshotReadRootOptions(given);
  checkReadRootOptions(where, options);
  const maxStartByte = options?.maxStartByte ?? 1_000_000;
  const maxTagBytes = options?.maxTagBytes ?? 65_536;
  /**
   * The most bytes fed: the limits, and up to three more to finish a character that straddles them —
   * whether it is one, and whether it is an error, decides what comes first. Once the parser has read
   * past `maxStartByte + maxTagBytes` a limit has been passed, so the read always ends by then.
   */
  const limit = maxStartByte + maxTagBytes + 3;

  const parser = SaxParser.withCheckedOptions({
    position: false,
    internalEntities: true,
    maxContextBytes: options?.maxContextBytes,
    maxEntityExpansions: options?.maxEntityExpansions,
    maxEntityExpansionLength: options?.maxEntityExpansionLength
  });
  let root: XmlRootElement | undefined;
  let startByte = 0;
  let endByte = 0;
  parser.on("error", err => {
    throw err;
  });
  parser.on("opentag", tag => {
    root = { name: tag.name, attributes: tag.attributes, context: parser.getPrologContext()! };
    startByte = parser.startByte!;
    endByte = parser.endByte!;
    parser.stop();
  });

  const reader = new ByteReader(checkedSource, where, options?.signal);
  let succeeded = false;
  try {
    const noRoot = (): XmlLimitExceededError =>
      new XmlLimitExceededError(
        `${where}: no root element begins before maxStartByte (${maxStartByte})`,
        "rootTag",
        maxStartByte,
        { byte: maxStartByte }
      );
    const tooLong = (open: number): XmlLimitExceededError =>
      new XmlLimitExceededError(
        `${where}: the root element's start tag at byte ${open} is longer than maxTagBytes (${maxTagBytes})`,
        "rootTag",
        maxTagBytes,
        { byte: open }
      );
    /**
     * The limit the read has passed once everything before `byte` is read and the root's `>` is not:
     * the root can no longer begin before `maxStartByte`, or its start tag — begun at `open` — can no
     * longer end within `maxTagBytes`. A limit and a parse error are told apart by which comes first
     * in the file, so the result does not depend on how the file is chunked.
     */
    const passed = (byte: number): XmlLimitExceededError | undefined => {
      const open = parser.inStartTag ? parser.openMarkupByte : undefined;
      if (open !== undefined && open < maxStartByte) {
        // A `<` alone may still begin a comment or a declaration rather than the root.
        return byte - open >= Math.max(maxTagBytes, 2) ? tooLong(open) : undefined;
      }
      return byte >= maxStartByte ? noRoot() : undefined;
    };
    let read = 0;
    while (root === undefined) {
      const chunk = await reader.next();
      // The signal may have fired since the chunk arrived: nothing is read after it has.
      reader.checkAborted();
      const piece =
        chunk === undefined || chunk.length <= limit - read
          ? chunk
          : chunk.subarray(0, limit - read);
      try {
        if (piece === undefined) {
          parser.close();
        } else {
          parser.write(piece);
        }
      } catch (err) {
        const limitFirst =
          err instanceof XmlParseError && err.byte !== undefined ? passed(err.byte) : undefined;
        throw limitFirst ?? err;
      }
      if (piece === undefined) {
        break;
      }
      read += piece.length;
      if (root === undefined) {
        // Known already: no need to wait for more of a source that may never send it. Judged at the
        // parser's position, not at `read`: the start of a character not yet whole is held back
        // undecoded, and may yet be an error that comes before the limit.
        const limitPassed = passed(parser.byte!);
        if (limitPassed !== undefined) {
          throw limitPassed;
        }
      }
    }
    if (root === undefined) {
      throw noRoot();
    }
    if (startByte >= maxStartByte) {
      throw noRoot();
    }
    if (endByte - startByte > maxTagBytes) {
      throw tooLong(startByte);
    }
    succeeded = true;
    return root;
  } finally {
    // A limit known to be passed is reported without waiting for the source to close; a root that
    // was read is returned once the source has closed.
    await reader.close(!succeeded);
  }
}

export { readRoot };
