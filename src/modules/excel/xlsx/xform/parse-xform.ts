/**
 * The loops that drive an xform through a parse: SAX events in, the xform's `parseOpen` / `parseText` /
 * `parseClose` callbacks called, its model out.
 *
 * Free functions rather than `BaseXform` methods so the XML parser belongs to the reader alone. Every
 * xform the writer renders extends `BaseXform`, and a class keeps all of its methods — so while these
 * lived on it, writing a workbook retained `@xml/sax` without ever parsing anything.
 */
import type { BaseXform } from "@excel/xlsx/xform/base-xform";
import { toError } from "@utils/errors";
import { SaxParser } from "@xml/sax";
import type { SaxEvent, SaxTag } from "@xml/types";

// HAN CELL namespace prefix normalization
// HAN CELL uses non-standard namespace prefixes (ep:, cp:, dc:, etc.)
const HAN_CELL_PREFIXES = new Set(["ep", "cp", "dc", "dcterms", "dcmitype", "vt"]);
const SPREADSHEETML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

// Detect HAN CELL mode from first tag. Returns:
// - undefined: normal file (no prefix handling needed)
// - null: HAN CELL file without spreadsheetml prefix (uses static prefixes only)
// - string: HAN CELL file with spreadsheetml prefix (e.g., "x")
function detectHanCellPrefix(
  tagName: string,
  attrs: Record<string, string>
): string | null | undefined {
  let hasHanCellPrefix = false;
  for (const key in attrs) {
    if (key.length > 6 && key.startsWith("xmlns:")) {
      const prefix = key.slice(6);
      // Check for spreadsheetml namespace prefix — always wins
      if (attrs[key] === SPREADSHEETML_NS) {
        return prefix;
      }
      // Note if we see a known HAN CELL prefix, but don't return yet
      // because spreadsheetml might appear later in the iteration
      if (HAN_CELL_PREFIXES.has(prefix)) {
        hasHanCellPrefix = true;
      }
    }
  }
  if (hasHanCellPrefix) {
    return null;
  }
  // Check if tag name has a known static prefix
  const i = tagName.indexOf(":");
  return i !== -1 && HAN_CELL_PREFIXES.has(tagName.slice(0, i)) ? null : undefined;
}

// Strip known namespace prefix from element name
function stripPrefix(name: string, nsPrefix: string | null): string {
  const i = name.indexOf(":");
  if (i === -1) {
    return name;
  }
  const p = name.slice(0, i);
  return p === nsPrefix || HAN_CELL_PREFIXES.has(p) ? name.slice(i + 1) : name;
}

/**
 * Drive `xform` from an iterable of SAX event batches (`parseSax`).
 */
export async function parseXformEvents<TModel>(
  xform: BaseXform<TModel>,
  saxParser: AsyncIterable<SaxEvent[]>
): Promise<TModel | undefined> {
  // IMPORTANT:
  // Do not return early once parsing is "done".
  // In true streaming scenarios, `parseSax(stream)` is backed by a Node.js
  // Readable async iterator. Returning early would close the iterator, which
  // destroys the underlying stream and can surface as AbortError (ABORT_ERR).
  let done = false;
  let finalModel: TModel | undefined;

  // HAN CELL compatibility: 0 = not checked, 1 = normal file, 2 = HAN CELL file
  let nsMode = 0;
  let nsPrefix: string | null = null;

  for await (const events of saxParser) {
    if (done) {
      continue;
    }
    for (const event of events) {
      if (event.eventType === "opentag") {
        const value = event.value;
        // Fast path for normal Excel files (majority case)
        if (nsMode === 1) {
          xform.parseOpen(value);
          continue;
        }
        // First tag - detect mode
        if (nsMode === 0) {
          const prefix = detectHanCellPrefix(value.name, value.attributes);
          if (prefix === undefined) {
            nsMode = 1;
            xform.parseOpen(value);
            continue;
          }
          nsMode = 2;
          nsPrefix = prefix;
        }
        // HAN CELL mode - strip prefix without mutating the SAX tag object
        const strippedName = stripPrefix(value.name, nsPrefix);
        if (strippedName !== value.name) {
          xform.parseOpen({
            name: strippedName,
            attributes: value.attributes,
            isSelfClosing: value.isSelfClosing
          });
        } else {
          xform.parseOpen(value);
        }
      } else if (event.eventType === "text") {
        xform.parseText(event.value);
      } else if (event.eventType === "cdata") {
        xform.parseCdata(event.value);
      } else if (event.eventType === "closetag") {
        const value = event.value;
        // Fast path for normal files
        if (nsMode === 1) {
          if (!xform.parseClose(value.name)) {
            done = true;
            finalModel = xform.model;
            break;
          }
          continue;
        }
        // HAN CELL mode - strip prefix
        if (!xform.parseClose(stripPrefix(value.name, nsPrefix))) {
          done = true;
          finalModel = xform.model;
          break;
        }
      }
    }
  }

  return done ? finalModel : xform.model;
}

/**
 * High-performance stream parsing using direct SAX callbacks.
 * Eliminates per-event object allocation and async generator overhead.
 * Use this instead of `parseXformEvents(xform, parseSax(stream))` for hot paths.
 */
export async function parseXformStream<TModel>(
  xform: BaseXform<TModel>,
  stream: AsyncIterable<unknown> | Iterable<unknown>
): Promise<TModel | undefined> {
  const parser = new SaxParser({ invalidCharHandling: "skip" });
  const decoder = new TextDecoder("utf-8", { fatal: true });

  let done = false;
  let finalModel: TModel | undefined;

  // HAN CELL compatibility: 0 = not checked, 1 = normal file, 2 = HAN CELL file
  let nsMode = 0;
  let nsPrefix: string | null = null;

  // Suppress errors that occur after we're done parsing (the SAX parser will
  // encounter unmatched tags when we stop processing close tags).
  // IMPORTANT: We set error handler FIRST, before any write() calls, to ensure
  // it's always present when fail() is called from within callbacks.
  let parseError: Error | undefined;
  parser.on("error", (err: Error) => {
    if (!done) {
      parseError = err;
    }
    // When done, silently swallow all SAX errors
  });

  parser.on("opentag", (tag: SaxTag) => {
    if (done) {
      return;
    }
    // Fast path for normal Excel files (majority case)
    if (nsMode === 1) {
      xform.parseOpen(tag);
      return;
    }
    // First tag - detect mode
    if (nsMode === 0) {
      const prefix = detectHanCellPrefix(tag.name, tag.attributes);
      if (prefix === undefined) {
        nsMode = 1;
        xform.parseOpen(tag);
        return;
      }
      nsMode = 2;
      nsPrefix = prefix;
    }
    // HAN CELL mode - strip prefix without mutating the SAX tag object
    // (the SAX parser reuses the tag on its internal stack for close-tag matching)
    const strippedName = stripPrefix(tag.name, nsPrefix);
    if (strippedName !== tag.name) {
      xform.parseOpen({
        name: strippedName,
        attributes: tag.attributes,
        isSelfClosing: tag.isSelfClosing
      });
    } else {
      xform.parseOpen(tag);
    }
  });

  parser.on("text", (text: string) => {
    if (!done) {
      xform.parseText(text);
    }
  });

  parser.on("cdata", (text: string) => {
    if (!done) {
      xform.parseCdata(text);
    }
  });

  parser.on("closetag", (tag: SaxTag) => {
    if (done) {
      return;
    }
    const name = nsMode === 2 ? stripPrefix(tag.name, nsPrefix) : tag.name;
    if (!xform.parseClose(name)) {
      done = true;
      finalModel = xform.model;
    }
  });

  // IMPORTANT: Do not return early from the async iterator.
  // In true streaming scenarios the iterator is backed by a Node.js Readable.
  // Returning early would close/destroy the stream (AbortError).
  // We must consume all chunks, but once done we skip writing to the parser.
  for await (const chunk of stream) {
    if (done) {
      continue;
    }
    const chunkStr =
      typeof chunk === "string" ? chunk : decoder.decode(chunk as Uint8Array, { stream: true });
    parser.write(chunkStr);
    if (parseError) {
      throw toError(parseError);
    }
  }

  if (!done) {
    // Flush trailing bytes from streaming decoder (catches truncated UTF-8)
    const trailing = decoder.decode();
    if (trailing) {
      parser.write(trailing);
      if (parseError) {
        throw toError(parseError);
      }
    }

    parser.close();
    if (parseError) {
      throw toError(parseError);
    }
  }

  return done ? finalModel : xform.model;
}
