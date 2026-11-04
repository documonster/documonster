import { createPartDecoder } from "@excel/utils/part-decoder";
/**
 * The loops that drive an xform through a parse: SAX events in, the xform's `parseOpen` / `parseText` /
 * `parseClose` callbacks called, its model out.
 *
 * Free functions rather than `BaseXform` methods so the XML parser belongs to the reader alone. Every
 * xform the writer renders extends `BaseXform`, and a class keeps all of its methods — so while these
 * lived on it, writing a workbook retained `@xml/sax` without ever parsing anything.
 */
import type { BaseXform } from "@excel/xlsx/xform/base-xform";
import {
  createNamespaceContext,
  withNamespaceContext
} from "@excel/xlsx/xform/fragment-namespaces";
import { createPrefixNormalizer } from "@excel/xlsx/xform/namespace-prefixes";
import { toError } from "@utils/errors";
import { SaxParser } from "@xml/sax";
import type { ParseOpenTag, SaxEvent, SaxTag } from "@xml/types";

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

  // Prefixes are normalised to the ones the xforms match first — see `namespace-prefixes.ts`.
  const prefixes = createPrefixNormalizer();
  // A capture an xform starts asks the normaliser for text — see `fragment-namespaces.ts`.
  const namespaces = createNamespaceContext(prefixes, xform);

  const dispatch = (events: SaxEvent[]): void => {
    for (const event of events) {
      if (event.eventType === "opentag") {
        const value = prefixes.open(event.value);
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
        const name = prefixes.close(event.value.name);
        // Fast path for normal files
        if (nsMode === 1) {
          if (!xform.parseClose(name)) {
            done = true;
            finalModel = xform.model;
            break;
          }
          continue;
        }
        // HAN CELL mode - strip prefix
        if (!xform.parseClose(stripPrefix(name, nsPrefix))) {
          done = true;
          finalModel = xform.model;
          break;
        }
      }
    }
  };

  for await (const events of saxParser) {
    if (!done) {
      withNamespaceContext(namespaces, () => dispatch(events));
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
  const decoder = createPartDecoder();

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

  // Prefixes are normalised to the ones the xforms match first — see `namespace-prefixes.ts`. The
  // normaliser stays out of the handlers until the text declares a prefix it has to rewrite, so a part
  // with Excel's prefixes is parsed exactly as if it did not exist.
  const prefixes = createPrefixNormalizer({ lazy: true, hinted: true });
  // A capture an xform starts asks the normaliser for text — see `fragment-namespaces.ts`.
  const namespaces = createNamespaceContext(prefixes, xform);
  // The parser's count of prefixed attribute names when the previous tag opened: a tag that moved it has one.
  let prefixedAttributes = 0;
  // Tested in the one handler rather than installing a second: two handler functions make the parser's call
  // site polymorphic for every part read afterwards, which costs more than the test.
  let normalizing = false;

  // Once the parser has reported an error, no xform hears anything more. The parser keeps going after an error and
  // delivers close tags for elements it never opened, and an xform handed one reads state it never set up — a
  // `TypeError` from inside the xform instead of the parser's `XmlParseError`.
  const onOpenTag = (raw: SaxTag): void => {
    if (done || parseError !== undefined) {
      return;
    }
    let tag: ParseOpenTag = raw;
    if (normalizing) {
      const count = parser.prefixedAttributes;
      tag = prefixes.open(raw, count !== prefixedAttributes);
      prefixedAttributes = count;
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
  };
  parser.on("opentag", onOpenTag);

  parser.on("text", (text: string) => {
    if (!done && parseError === undefined) {
      xform.parseText(text);
    }
  });

  parser.on("cdata", (text: string) => {
    if (!done && parseError === undefined) {
      xform.parseCdata(text);
    }
  });

  const onCloseTag = (tag: SaxTag): void => {
    if (done || parseError !== undefined) {
      return;
    }
    const raw = normalizing ? prefixes.close(tag.name) : tag.name;
    const name = nsMode === 2 ? stripPrefix(raw, nsPrefix) : raw;
    if (!xform.parseClose(name)) {
      done = true;
      finalModel = xform.model;
    }
  };
  parser.on("closetag", onCloseTag);

  const write = (text: string): void => {
    // Scanned even once active: the normaliser counts declarations to skip looking for them per tag.
    if (prefixes.scan(text) && !normalizing) {
      normalizing = true;
      parser.countPrefixedAttributes = true;
      // Counting starts here, so a tag the parser is part-way through may have read a prefixed name uncounted. It
      // is the one tag that cannot be vouched for, and a count one ahead makes it report having one.
      prefixedAttributes = -1;
    }
    withNamespaceContext(namespaces, () => parser.write(text));
  };

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
    write(chunkStr);
    if (parseError) {
      throw toError(parseError);
    }
  }

  if (!done) {
    // Flush trailing bytes from streaming decoder (catches truncated UTF-8)
    const trailing = decoder.decode();
    if (trailing) {
      write(trailing);
      if (parseError) {
        throw toError(parseError);
      }
    }

    withNamespaceContext(namespaces, () => parser.close());
    if (parseError) {
      throw toError(parseError);
    }
  }

  return done ? finalModel : xform.model;
}
