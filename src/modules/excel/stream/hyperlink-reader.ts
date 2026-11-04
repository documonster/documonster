/**
 * HyperlinkReader - Streaming Hyperlink Reader
 *
 * Cross-platform implementation.
 */

import { Enums } from "@excel/core/enums";
import type { InternalWorksheetOptions } from "@excel/stream/workbook-reader.browser";
import { createPartDecoder } from "@excel/utils/part-decoder";
import { RelType } from "@excel/xlsx/rel-type";
import { createPrefixNormalizer } from "@excel/xlsx/xform/namespace-prefixes";
import { EventEmitter } from "@utils/event-emitter";
import { SaxParser } from "@xml/sax";
import type { SaxTag } from "@xml/types";

export interface HyperlinkReaderOptions<TWorkbook = unknown> {
  workbook: TWorkbook;
  id: number;
  iterator: AsyncIterable<unknown>;
  options: InternalWorksheetOptions;
}

/** Hyperlink relationship parsed from worksheet rels */
export interface Hyperlink {
  type: number;
  rId: string;
  target: string;
  targetMode: string;
}

/**
 * A sheet's hyperlink relationships, from a streaming workbook reader with `hyperlinks: "emit"`: `read()` reads them,
 * raising a `hyperlink` event each, and `each` visits the ones read.
 */
class HyperlinkReader extends EventEmitter {
  workbook: unknown;
  id: number;
  iterator: AsyncIterable<unknown>;
  options: InternalWorksheetOptions;
  hyperlinks?: Record<string, Hyperlink>;
  private _hyperlinkCount = 0;

  constructor({ workbook, id, iterator, options }: HyperlinkReaderOptions) {
    super();

    this.workbook = workbook;
    this.id = id;
    this.iterator = iterator;
    this.options = options;
  }

  get count(): number {
    return this.hyperlinks ? this._hyperlinkCount : 0;
  }

  each(fn: (hyperlink: Hyperlink, rId: string) => void): void {
    const hyperlinks = this.hyperlinks;
    if (!hyperlinks) {
      return;
    }

    for (const rId in hyperlinks) {
      fn(hyperlinks[rId], rId);
    }
  }

  async read(): Promise<void> {
    const { iterator } = this;
    const hyperlinkMode = this.options.hyperlinks;

    const emitHyperlinks = hyperlinkMode === "emit";
    const cacheHyperlinks = hyperlinkMode === "cache";

    let cachedHyperlinks: Record<string, Hyperlink> | null = null;
    if (cacheHyperlinks) {
      this._hyperlinkCount = 0;
      this.hyperlinks = cachedHyperlinks = Object.create(null) as Record<string, Hyperlink>;
    }

    if (!emitHyperlinks && !cacheHyperlinks) {
      this.emit("finished");
      return;
    }

    try {
      const parser = new SaxParser({ position: false, invalidCharHandling: "skip" });
      // Prefixes normalised to the ones matched below — see `namespace-prefixes.ts`.
      const parserPrefixes = createPrefixNormalizer({ lazy: true });
      const decoder = createPartDecoder();

      parser.on("closetag", (rawTag: SaxTag) => {
        parserPrefixes.close(rawTag.name);
      });
      parser.on("opentag", (rawNode: SaxTag) => {
        const node = parserPrefixes.open(rawNode);
        if (node.name !== "Relationship") {
          return;
        }

        const attributes = node.attributes;
        if (attributes.Type !== RelType.Hyperlink) {
          return;
        }

        const relationship: Hyperlink = {
          type: Enums.RelationshipType.Hyperlink,
          rId: attributes.Id,
          target: attributes.Target,
          targetMode: attributes.TargetMode
        };

        if (emitHyperlinks) {
          this.emit("hyperlink", relationship);
          return;
        }

        // cache mode
        const rId = relationship.rId;
        if (cachedHyperlinks && cachedHyperlinks[rId] === undefined) {
          this._hyperlinkCount += 1;
        }
        cachedHyperlinks![rId] = relationship;
      });

      for await (const chunk of iterator) {
        const chunkStr =
          typeof chunk === "string" ? chunk : decoder.decode(chunk as Uint8Array, { stream: true });
        parserPrefixes.scan(chunkStr);
        parser.write(chunkStr);
      }

      // Flush trailing bytes (catches truncated UTF-8)
      const trailing = decoder.decode();
      if (trailing) {
        parserPrefixes.scan(trailing);
        parser.write(trailing);
      }

      parser.close();

      this.emit("finished");
    } catch (error) {
      this.emit("error", error);
    }
  }
}

export { HyperlinkReader };
