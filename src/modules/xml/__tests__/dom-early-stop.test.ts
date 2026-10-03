import { parseXml } from "@xml/dom";
import { XmlParseError } from "@xml/errors";
import { SaxParser } from "@xml/sax";
import { parseXmlToObject } from "@xml/to-object";
import { afterEach, describe, expect, it, vi } from "vitest";

// Every error the SAX parser raises goes through `fail()`. Before the DOM
// builders stopped at the first error, each violation after the first (every
// deeper tag, every further entity, every later malformed construct) called
// `fail()` again while the rest of the document was still being processed.
// Counting calls proves the parse stopped, without timing anything.

afterEach(() => {
  vi.restoreAllMocks();
});

function catchError(fn: () => unknown): XmlParseError {
  try {
    fn();
  } catch (err) {
    expect(err).toBeInstanceOf(XmlParseError);
    return err as XmlParseError;
  }
  throw new Error("expected a parse error");
}

/** 13 nested levels (limit 10), then a malformed tail that would fail again. */
const DEEP_THEN_BROKEN =
  "<a>".repeat(13) + "<b/>".repeat(1000) + "</mismatch><c x=1/>&undefined;" + "</a>".repeat(13);

describe("DOM builders stop at the first error", () => {
  for (const [name, parse] of [
    ["parseXml", (xml: string, maxDepth: number) => parseXml(xml, { maxDepth })],
    ["parseXmlToObject", (xml: string, maxDepth: number) => parseXmlToObject(xml, { maxDepth })]
  ] as const) {
    it(`${name}: depth limit stops parsing`, () => {
      const fail = vi.spyOn(SaxParser.prototype, "fail");
      const opentag = vi.fn();
      const on = SaxParser.prototype.on;
      vi.spyOn(SaxParser.prototype, "on").mockImplementation(function (
        this: SaxParser,
        event: string,
        handler: (...args: never[]) => void
      ) {
        if (event === "opentag") {
          const wrapped = (...args: never[]) => {
            opentag();
            handler(...args);
          };
          return on.call(this, event as never, wrapped as never);
        }
        return on.call(this, event as never, handler as never);
      } as never);

      const err = catchError(() => parse(DEEP_THEN_BROKEN, 10));
      expect(err.limit).toBe("depth");
      expect(fail).toHaveBeenCalledTimes(1);
      // 10 levels were built, then the 11th was rejected; the 1000 siblings never were.
      expect(opentag.mock.calls.length).toBeLessThanOrEqual(11);
    });
  }

  it("entity expansion limit stops parsing", () => {
    const fail = vi.spyOn(SaxParser.prototype, "fail");
    const write = SaxParser.prototype.write;
    vi.spyOn(SaxParser.prototype, "write").mockImplementation(function (
      this: SaxParser,
      chunk: string | null
    ) {
      this.ENTITIES["x"] = "X";
      return write.call(this, chunk);
    });

    const xml = "<root>" + "&x;".repeat(10_000) + "</mismatch></root>";
    const err = catchError(() => parseXml(xml, { maxEntityExpansions: 5 }));
    expect(err.limit).toBe("entityExpansions");
    expect(fail).toHaveBeenCalledTimes(1);
  });

  it("non-limit errors still report the first error", () => {
    const fail = vi.spyOn(SaxParser.prototype, "fail");
    const err = catchError(() => parseXml("<a></b><c x=1/></a>"));
    expect(err.limit).toBeUndefined();
    expect(fail).toHaveBeenCalledTimes(1);
  });

  it("SAX with an error handler still continues after a depth violation", () => {
    const parser = new SaxParser({ maxDepth: 1 });
    const errors: XmlParseError[] = [];
    const names: string[] = [];
    parser.on("error", e => errors.push(e as XmlParseError));
    parser.on("opentag", t => names.push(t.name));
    parser.write("<a><b><c/></b><d/></a>").close();
    expect(errors.length).toBeGreaterThan(0);
    expect(errors[0].limit).toBe("depth");
    expect(names).toContain("d");
  });
});
