/**
 * The SAX parser given bytes: UTF-8 decoded across chunk boundaries, byte positions on every event
 * and error, `resume` from inside a document, the lexical state that tells two parsers they agree,
 * `stop()`, and the prolog context a later parser needs.
 */

import { parseXml } from "@xml/dom";
import {
  XmlArgumentError,
  XmlLimitExceededError,
  XmlParseError,
  XmlTruncatedError
} from "@xml/errors";
import { SaxParser } from "@xml/sax";
import { parseXmlToObject } from "@xml/to-object";
import type { XmlPrologContext, XmlSaxOptions } from "@xml/types";
import { describe, expect, it } from "vitest";

import { generateDocument } from "./helpers/generated-documents";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

interface Event {
  readonly kind: string;
  readonly value: string;
  readonly start?: number;
  readonly end?: number;
}

/** Every event of `input` written in chunks of `size`, with its byte range. */
function events(input: Uint8Array | string, size: number, options?: XmlSaxOptions): Event[] {
  const out: Event[] = [];
  const parser = new SaxParser({ position: false, ...options });
  const push = (kind: string, value: string): void => {
    out.push({ kind, value, start: parser.startByte, end: parser.endByte });
  };
  parser.on("opentag", tag => push("open", tag.name));
  parser.on("closetag", tag => push("close", tag.name));
  parser.on("text", text => push("text", text));
  parser.on("cdata", text => push("cdata", text));
  parser.on("comment", text => push("comment", text));
  parser.on("pi", (target, body) => push("pi", `${target} ${body}`));
  parser.on("error", err => out.push({ kind: "error", value: err.message }));
  for (let at = 0; at < input.length; at += size) {
    parser.write(input.slice(at, at + size) as Uint8Array & string);
  }
  parser.close();
  return out;
}

describe("SaxParser given bytes", () => {
  const xml =
    '<?xml version="1.0" encoding="UTF-8"?>\r\n<!-- é -->\n<d a="ü">日本<![CDATA[😀]]><e/>x&amp;y<?pi €?></d>\n';
  const bytes = encoder.encode(xml);

  it("reports the byte range of every event, the same however the bytes are split", () => {
    const whole = events(bytes, bytes.length);
    for (const size of [1, 2, 3, 7]) {
      expect(events(bytes, size)).toEqual(whole);
    }
    const expected: Record<string, string> = {
      open: '<d a="ü">',
      comment: "<!-- é -->",
      cdata: "<![CDATA[😀]]>",
      pi: "<?pi €?>"
    };
    for (const event of whole) {
      const text = decoder.decode(bytes.subarray(event.start!, event.end!));
      if (event.kind === "open" && event.value === "e") {
        expect(text).toBe("<e/>");
      } else if (event.kind === "close" && event.value === "d") {
        expect(text).toBe("</d>");
      } else if (event.kind === "text") {
        expect(text.replace(/\r\n/g, "\n").replace("&amp;", "&")).toBe(event.value);
      } else if (event.kind === "open" && event.value === "d") {
        expect(text).toBe(expected.open);
      } else if (event.kind === "pi" && !event.value.startsWith("pi ")) {
        expect(text).toBe('<?xml version="1.0" encoding="UTF-8"?>');
      } else if (expected[event.kind] && event.kind !== "open") {
        expect(text).toBe(expected[event.kind]);
      }
    }
  });

  it("gives events the same text as the same document given as a string", () => {
    const strip = (list: Event[]): string[] => list.map(e => `${e.kind}:${e.value}`);
    expect(strip(events(bytes, 5))).toEqual(strip(events(xml, xml.length)));
  });

  it("has no byte positions when given text", () => {
    for (const event of events(xml, 4)) {
      expect(event.start).toBeUndefined();
      expect(event.end).toBeUndefined();
    }
  });

  it("counts byteOffset into every position", () => {
    const shifted = events(bytes, 9, { byteOffset: 1000 });
    expect(shifted.map(e => e.start)).toEqual(events(bytes, 9).map(e => (e.start ?? 0) + 1000));
  });

  it("strips a byte order mark only at offset 0", () => {
    const withBom = new Uint8Array([0xef, 0xbb, 0xbf, ...encoder.encode("<d/>")]);
    expect(events(withBom, 1)[0]).toEqual({ kind: "open", value: "d", start: 3, end: 7 });
    const later = events(new Uint8Array([0xef, 0xbb, 0xbf, ...encoder.encode("<d/>")]), 1, {
      byteOffset: 10,
      resume: true
    });
    expect(later[0]).toEqual({ kind: "text", value: "\uFEFF", start: 10, end: 13 });
  });

  it("reads a byte that is not UTF-8 as invalidCharHandling says, one byte at a time", () => {
    const bad = new Uint8Array([
      ...encoder.encode("<d>a"),
      0xe2,
      0x28,
      0xa1,
      ...encoder.encode("b</d>")
    ]);
    for (const size of [1, 2, 64]) {
      expect(events(bad, size, { invalidCharHandling: "replace" })[1].value).toBe(
        "a\uFFFD(\uFFFDb"
      );
      expect(events(bad, size, { invalidCharHandling: "skip" })[1].value).toBe("a(b");
      const errors = events(bad, size).filter(e => e.kind === "error");
      expect(errors.map(e => e.value)).toEqual([
        "invalid UTF-8: byte 0xe2 is not part of a well-formed sequence (at byte 4)",
        "invalid UTF-8: byte 0xa1 is not part of a well-formed sequence (at byte 6)"
      ]);
    }
  });

  it("refuses a declared encoding other than UTF-8", () => {
    const latin = encoder.encode('<?xml version="1.0" encoding="ISO-8859-1"?><d/>');
    expect(events(latin, 64)[0].value).toMatch(
      /declares encoding ISO-8859-1; bytes are read as UTF-8 only/
    );
    expect(
      events(encoder.encode('<?xml version="1.0" encoding="us-ascii"?><d/>'), 64)[0].kind
    ).toBe("pi");
  });

  it("reports an input that ends inside a character as truncated", () => {
    const parser = new SaxParser();
    const errors: Error[] = [];
    parser.on("error", err => errors.push(err));
    parser.write(new Uint8Array([...encoder.encode("<d>"), 0xe6, 0x97]));
    parser.close();
    expect(errors[0]).toBeInstanceOf(XmlTruncatedError);
    expect(errors[0].message).toMatch(/inside a UTF-8 character/);
  });

  it("puts the byte offset on errors", () => {
    const parser = new SaxParser();
    const err = (() => {
      try {
        parser.write(encoder.encode("<d>ü<1/></d>"));
      } catch (e) {
        return e as XmlParseError;
      }
      return undefined;
    })();
    expect(err).toBeInstanceOf(XmlParseError);
    expect(err!.byte).toBe(6);
    expect(err!.message).toMatch(/\(at byte 6\)$/);
  });

  it("refuses text and bytes in one document, and anything else", () => {
    const parser = new SaxParser();
    parser.write(encoder.encode("<d>"));
    expect(() => parser.write("</d>")).toThrow(XmlArgumentError);
    expect(() => new SaxParser().write(42 as never)).toThrow(XmlArgumentError);
  });

  it("reads names outside the Basic Multilingual Plane, in start and end tags", () => {
    const name = "a\u{10000}b";
    const list = events(encoder.encode(`<${name} ${name}="1"></${name}>`), 1);
    expect(list.filter(e => e.kind !== "error").map(e => `${e.kind}:${e.value}`)).toEqual([
      `open:${name}`,
      `close:${name}`
    ]);
    expect(list.some(e => e.kind === "error")).toBe(false);
  });
});

describe("SaxParser.stop", () => {
  it("reads nothing after the event it is called in, the rest of the chunk included", () => {
    const parser = new SaxParser();
    const seen: string[] = [];
    parser.on("opentag", tag => {
      seen.push(tag.name);
      parser.stop();
    });
    parser.on("error", err => seen.push(err.message));
    parser.write(encoder.encode("<a x='1'><b/><<<broken"));
    expect(seen).toEqual(["a"]);
    expect(parser.stopped).toBe(true);
    expect(parser.byte).toBe(9);
    parser.write(encoder.encode("more"));
    expect(seen).toEqual(["a"]);
    // Closing reports nothing about what was left unread, and the next write starts afresh.
    parser.close();
    parser.write("<c/>");
    parser.close();
    expect(seen).toEqual(["a", "c"]);
  });
});

describe("SaxParser resume", () => {
  it("passes over end tags of elements opened before it and does not ask for a root", () => {
    const list = events(encoder.encode("text</x></y><r>1</r></z>tail"), 3, {
      resume: true,
      byteOffset: 50
    });
    expect(list.filter(e => e.kind === "error")).toEqual([]);
    expect(list.filter(e => e.kind === "open").map(e => e.start)).toEqual([62]);
  });

  it("drops the rest of a character it starts inside", () => {
    const bytes = encoder.encode("ü<r/>").subarray(1);
    const list = events(bytes, 1, { resume: true, byteOffset: 1 });
    expect(list).toEqual([
      { kind: "open", value: "r", start: 2, end: 6 },
      { kind: "close", value: "r", start: 2, end: 6 }
    ]);
  });

  it("still reports what is wrong whatever came before", () => {
    const list = events(encoder.encode("<a><b></a>"), 64, { resume: true });
    expect(list.filter(e => e.kind === "error").map(e => e.value)).toEqual([
      "unclosed tag: b (at byte 9)"
    ]);
  });

  it("expands the entities of the context it is given", () => {
    const list = events(encoder.encode("<r a='&e;'>&e;</r></d>"), 2, {
      resume: true,
      context: { root: "d", entities: { e: "x&amp;y" } }
    });
    expect(list.map(e => e.value)).toEqual(["r", "x&y", "r"]);
  });
});

describe("SaxParser.getLexicalState", () => {
  /** The lexical state after each byte of `bytes` from `from`, for a parser begun at `start`. */
  function states(bytes: Uint8Array, start: number, resume: boolean): string[] {
    const parser = new SaxParser({ position: false, byteOffset: start, resume });
    parser.on("error", () => {});
    const out: string[] = [];
    for (let at = start; at < bytes.length; at++) {
      parser.write(bytes.subarray(at, at + 1));
      out.push(parser.getLexicalState());
    }
    return out;
  }

  it("meets the state of a reading from the start once both have read the same markup", () => {
    const bytes = encoder.encode('<d>\n  <r id="1"><t>a</t></r>\n  <r id="2"><t>b</t></r>\n</d>');
    const whole = states(bytes, 0, false);
    for (let start = 4; start < bytes.length - 6; start++) {
      const later = states(bytes, start, true);
      const meet = later.findIndex((state, i) => state === whole[start + i]);
      expect(meet, `start ${start}`).toBeGreaterThanOrEqual(0);
      // Once met, they stay met.
      expect(later.slice(meet)).toEqual(whole.slice(start + meet));
    }
  });

  it("gives equal parsers equal events from there on, on generated documents", () => {
    for (let seed = 1; seed <= 20; seed++) {
      const bytes = encoder.encode(
        generateDocument(seed, { decoys: true, sameNameNesting: true }).xml
      );
      const whole = states(bytes, 0, false);
      const start = Math.floor(bytes.length / 3);
      const later = states(bytes, start, true);
      const meet = later.findIndex((state, i) => state === whole[start + i]);
      if (meet >= 0) {
        expect(later.slice(meet)).toEqual(whole.slice(start + meet));
        const at = start + meet + 1;
        const tail = (parser: SaxParser): Event[] => {
          const out: Event[] = [];
          parser.on("opentag", tag =>
            out.push({ kind: "open", value: tag.name, start: parser.startByte })
          );
          parser.on("text", text =>
            out.push({ kind: "text", value: text, start: parser.startByte })
          );
          parser.write(bytes.subarray(at));
          return out;
        };
        const a = new SaxParser({ position: false });
        a.on("error", () => {});
        a.write(bytes.subarray(0, at));
        const b = new SaxParser({ position: false, byteOffset: start, resume: true });
        b.on("error", () => {});
        b.write(bytes.subarray(start, at));
        expect(tail(b), `seed ${seed}`).toEqual(tail(a));
      }
    }
  });
});

describe("SaxParser internalEntities and getPrologContext", () => {
  const xml =
    '<!DOCTYPE d [ <!ENTITY a "A"> <!ENTITY b "[&a;]"> <!ENTITY ext SYSTEM "x.xml"> <!ENTITY a "ignored"> ]><d x="&b;">&b;</d>';

  it("expands declared entities, the first declaration binding", () => {
    const list = events(xml, xml.length, { internalEntities: true });
    expect(list.map(e => e.value)).toEqual(["d", "[A]", "d"]);
  });

  it("carries the root and the entities, ready for JSON, once the root is read", () => {
    const parser = new SaxParser({ internalEntities: true });
    expect(parser.getPrologContext()).toBeUndefined();
    let context: unknown;
    parser.on("opentag", () => {
      context = parser.getPrologContext();
    });
    parser.write(xml).close();
    expect(JSON.parse(JSON.stringify(context))).toEqual({
      root: "d",
      entities: { a: "A", b: "[&a;]" }
    });
  });

  it("does not read an external entity, and says so", () => {
    const list = events('<!DOCTYPE d [ <!ENTITY ext SYSTEM "x.xml"> ]><d>&ext;</d>', 64, {
      internalEntities: true
    });
    expect(list.find(e => e.kind === "error")!.value).toBe(
      "undefined entity: ext (external entities are not read)"
    );
  });

  it("processes no declaration after a parameter-entity reference", () => {
    const list = events('<!DOCTYPE d [ <!ENTITY % p "x"> %p; <!ENTITY e "E"> ]><d>&e;</d>', 64, {
      internalEntities: true
    });
    expect(list.find(e => e.kind === "error")!.value).toBe("undefined entity: e");
  });

  it("refuses an entity whose replacement text holds markup", () => {
    const list = events('<!DOCTYPE d [ <!ENTITY e "<b/>"> ]><d>&e;</d>', 64, {
      internalEntities: true
    });
    expect(list.find(e => e.kind === "error")!.value).toBe(
      "entity e contains markup, which is not expanded"
    );
  });

  it("counts nested expansions against maxEntityExpansions before building them", () => {
    let doctype = '<!DOCTYPE d [ <!ENTITY l0 "ha">';
    for (let i = 1; i <= 9; i++) {
      doctype += ` <!ENTITY l${i} "${`&l${i - 1};`.repeat(10)}">`;
    }
    const parser = new SaxParser({ internalEntities: true });
    expect(() => parser.write(`${doctype} ]><d>&l9;</d>`)).toThrow(XmlLimitExceededError);
  });

  it("refuses a context larger than maxContextBytes", () => {
    const big = "x".repeat(200);
    const parser = new SaxParser({ internalEntities: true, maxContextBytes: 100 });
    const err = (() => {
      try {
        parser.write(`<!DOCTYPE d [ <!ENTITY e "${big}"> ]><d/>`);
      } catch (e) {
        return e;
      }
      return undefined;
    })();
    expect(err).toBeInstanceOf(XmlLimitExceededError);
    expect((err as XmlLimitExceededError).limit).toBe("prologContext");
  });

  it("refuses a declaration too long to hold and check, unless the bound is raised", () => {
    const values = Array.from({ length: 20000 }, (_, i) => `v${i}`).join("|");
    const doc = `<!DOCTYPE d [ <!ATTLIST d a (${values}) #IMPLIED> <!ENTITY e "E"> ]><d>&e;</d>`;
    // Longer than maxContextBytes, it cannot be checked without being held: refused as past the
    // limit, rather than taken unchecked — whatever it declares, and whether or not it is used.
    for (const internalEntities of [true, false]) {
      expect(
        () => new SaxParser({ internalEntities }).write(doc),
        String(internalEntities)
      ).toThrow(expect.objectContaining({ limit: "prologContext" }));
    }
    expect(() =>
      new SaxParser({ internalEntities: true }).write(
        `<!DOCTYPE d [ <!ENTITY e "${"x".repeat(70000)}"> ]><d/>`
      )
    ).toThrow(XmlLimitExceededError);
    // Raised, it is held and checked.
    expect(
      events(doc, 4096, { internalEntities: true, maxContextBytes: 1 << 20 }).map(e => e.value)
    ).toEqual(["d", "E", "d"]);
    // Read whole, the input is all in memory, and no bound applies.
    expect(parseXml(doc, { internalEntities: true }).root.name).toBe("d");
  });

  it("refuses a long chain of entities by the limit, and expands it without one", () => {
    const entities: Record<string, string> = {};
    const length = 20000;
    for (let i = 0; i < length; i++) {
      entities[`e${i}`] = i === length - 1 ? "end" : `&e${i + 1};`;
    }
    const context = { root: "d", entities };
    const limited = new SaxParser({
      resume: true,
      context,
      maxEntityExpansions: 10,
      maxContextBytes: 1e7
    });
    expect(() => limited.write(encoder.encode("<r>&e0;</r>"))).toThrow(XmlLimitExceededError);
    const unlimited = new SaxParser({
      resume: true,
      context,
      maxEntityExpansions: 0,
      maxContextBytes: 1e7
    });
    let text = "";
    unlimited.on("text", t => (text += t));
    unlimited.write(encoder.encode("<r>&e0;</r>"));
    expect(text).toBe("end");
  });

  it("counts the use of each entity against maxEntityExpansions, the first included", () => {
    const doc = '<!DOCTYPE R [ <!ENTITY a "A"> <!ENTITY b "B"> ]><R>&a;&b;</R>';
    expect(() =>
      new SaxParser({ internalEntities: true, maxEntityExpansions: 1 }).write(doc)
    ).toThrow(XmlLimitExceededError);
    let text = "";
    const two = new SaxParser({ internalEntities: true, maxEntityExpansions: 2 });
    two.on("text", t => (text += t));
    two.write(doc).close();
    expect(text).toBe("AB");
  });

  it("places an error found by the attribute fast path where the state machine does", () => {
    const bytes = encoder.encode('<R a="123456789" a="x"/>');
    const places = [1, 3, bytes.length].map(size => {
      const parser = new SaxParser({ position: false });
      try {
        for (let at = 0; at < bytes.length; at += size) {
          parser.write(bytes.subarray(at, at + size));
        }
        parser.close();
      } catch (err) {
        return (err as XmlParseError).byte;
      }
      return undefined;
    });
    expect(places).toEqual([23, 23, 23]);
  });

  it("reports byte as everything written once a chunk is read through", () => {
    const bytes = encoder.encode("  <?pi x?> <R a='1'>");
    const parser = new SaxParser({ position: false });
    for (let at = 0; at < bytes.length; at++) {
      parser.write(bytes.subarray(at, at + 1));
      expect(parser.byte, `after byte ${at}`).toBe(at + 1);
    }
  });

  it("counts position as the characters read, between writes and inside handlers", () => {
    const parser = new SaxParser();
    const inside: number[] = [];
    parser.on("opentag", () => inside.push(parser.position));
    parser.write("<R>");
    expect(parser.position).toBe(3);
    parser.write("é<a/>");
    expect(parser.position).toBe(8);
    expect(inside).toEqual([3, 8]);
  });

  it("measures what entities expand to, and refuses a text too long before building it", () => {
    // `c` expands to 27 UTF-16 units — the emoji is two, each reference one character — used twice.
    const doc =
      '<!DOCTYPE R [<!ENTITY a "x&#x1F600;&amp;&#65;"><!ENTITY b "&a;-&a;\ty"><!ENTITY c "&b;&lt;&b;">]><R k="&c;">&c;</R>';
    const read = (maxEntityExpansionLength: number): string => {
      try {
        parseXmlToObject(doc, { internalEntities: true, maxEntityExpansionLength });
        return "ok";
      } catch (err) {
        return (err as XmlLimitExceededError).limit;
      }
    };
    expect([read(54), read(53), read(0)]).toEqual(["ok", "entityExpansionLength", "ok"]);
    // 90 KB that would expand to 539,500,000 characters, well within maxEntityExpansions.
    const bomb = `<!DOCTYPE R [<!ENTITY e "${"x".repeat(65000)}">]><R a="${"&e;".repeat(8300)}"/>`;
    const err = (() => {
      try {
        new SaxParser({ internalEntities: true }).write(bomb);
      } catch (e) {
        return e;
      }
      return undefined;
    })();
    expect(err).toBeInstanceOf(XmlLimitExceededError);
    expect((err as XmlLimitExceededError).limit).toBe("entityExpansionLength");
    expect((err as XmlLimitExceededError).allowed).toBe(10_000_000);
  });

  it("reports text outside the root where it is, before anything after it", () => {
    for (const doc of ["x \u0000<R/>", "x\n\n <R/>"]) {
      const bytes = encoder.encode(doc);
      for (const size of [1, bytes.length]) {
        const parser = new SaxParser();
        let error: XmlParseError | undefined;
        try {
          for (let at = 0; at < bytes.length; at += size) {
            parser.write(bytes.subarray(at, at + size));
          }
        } catch (err) {
          error = err as XmlParseError;
        }
        expect(error?.message, `${JSON.stringify(doc)} in chunks of ${size}`).toContain(
          "text data outside of root node"
        );
        expect([error?.byte, error?.line, error?.column]).toEqual([0, 1, 1]);
      }
    }
  });

  it("parses a DOCTYPE of any size when the whole input is in memory", () => {
    const doc = `<!DOCTYPE R [<!ENTITY e "${"x".repeat(70000)}">]><R>&e;</R>`;
    expect((parseXmlToObject(doc, { internalEntities: true }).R as string).length).toBe(70000);
    expect(parseXml(doc, { internalEntities: true }).root.name).toBe("R");
    // A streamed parse holds the declarations, so it keeps the bound.
    expect(() => new SaxParser({ internalEntities: true }).write(doc)).toThrow(
      XmlLimitExceededError
    );
  });

  it("Xml.parse and Xml.parseToObject use the options they checked, read once", () => {
    let reads = 0;
    const options = {
      get textKey(): string {
        return ++reads === 1 ? "#t" : (42 as unknown as string);
      }
    };
    expect(parseXmlToObject("<r>x<b/></r>", options)).toEqual({ r: { b: "", "#t": "x" } });
    expect(reads).toBe(1);
    let fragmentReads = 0;
    const parseOptions = {
      get fragment(): boolean {
        return ++fragmentReads === 1 ? false : (7 as unknown as boolean);
      }
    };
    expect(parseXml("<r/>", parseOptions).root.name).toBe("r");
    expect(fragmentReads).toBe(1);
  });

  it("uses the options it checked, read once", () => {
    let reads = 0;
    const options = {
      get maxDepth(): number {
        return ++reads === 1 ? 1 : -1;
      }
    };
    expect(() => new SaxParser(options).write("<R><A/></R>")).toThrow(XmlLimitExceededError);
    expect(reads).toBe(1);
    // Inherited options are options.
    expect(() =>
      new SaxParser(Object.create({ maxDepth: 1 }) as object).write("<R><A/></R>")
    ).toThrow(XmlLimitExceededError);
  });

  it("bounds the context by the bytes of its JSON form, exactly", async () => {
    for (const doc of [
      '<!DOCTYPE Rü [ <!ENTITY é "€&quot;"> ]><Rü/>',
      // Bytes declared ASCII: the context carries `encoding` too.
      encoder.encode('<?xml version="1.0" encoding="ascii"?><!DOCTYPE R [ <!ENTITY e "v"> ]><R/>')
    ]) {
      const parser = new SaxParser({ internalEntities: true });
      parser.write(doc);
      const size = encoder.encode(JSON.stringify(parser.getPrologContext())).length;
      expect(() =>
        new SaxParser({ internalEntities: true, maxContextBytes: size }).write(doc)
      ).not.toThrow();
      expect(() =>
        new SaxParser({ internalEntities: true, maxContextBytes: size - 1 }).write(doc)
      ).toThrow(XmlLimitExceededError);
    }
  });

  it("measures a context it is given by its JSON form, every field counted", () => {
    for (const context of [
      { root: "r", entities: {}, attributes: {} },
      { root: "r", entities: {}, encoding: "ascii" as const },
      { root: "r", entities: { e: "x" }, encoding: "ascii" as const, attributes: {} }
    ] as XmlPrologContext[]) {
      const size = encoder.encode(JSON.stringify(context)).length;
      expect(() => new SaxParser({ resume: true, context, maxContextBytes: size })).not.toThrow();
      expect(() => new SaxParser({ resume: true, context, maxContextBytes: size - 1 })).toThrow(
        XmlArgumentError
      );
    }
  });

  it("accepts the same processing instructions whether or not anyone is given them", () => {
    const docs = [
      `<?${"x".repeat(300)}!?><R/>`, // a bad character far into a long target
      `<?${"x".repeat(300)} body?><R/>`, // a long, valid target
      `<!DOCTYPE R [ <?${"x".repeat(300)}! ?> ]><R/>` // the same, in the internal subset
    ];
    for (const doc of docs) {
      const outcome = (listen: boolean): string => {
        const parser = new SaxParser();
        if (listen) {
          parser.on("pi", () => {});
        }
        try {
          parser.write(doc).close();
          return "accepted";
        } catch (err) {
          return (err as Error).message;
        }
      };
      expect(outcome(false), doc.slice(0, 20)).toBe(outcome(true));
    }
    // An XML declaration's encoding is checked however far into it.
    const declaration = encoder.encode(
      `<?xml version="1.0"${" ".repeat(300)}encoding="ISO-8859-1"?><R/>`
    );
    expect(() => new SaxParser().write(declaration)).toThrow(/ISO-8859-1/);
  });

  it("holds no comment, CDATA or processing instruction text that nobody is given", () => {
    const parser = new SaxParser({ position: false });
    parser.write(`<r/><!--${"x".repeat(100000)}`);
    expect((parser as unknown as { text: string }).text).toBe("");
    const told = new SaxParser({ position: false });
    let comment = "";
    told.on("comment", c => (comment = c));
    told.write(`<r/><!--${"x".repeat(1000)}-->`);
    expect(comment).toHaveLength(1000);
  });

  it("keeps the old report without internalEntities", () => {
    const list = events('<!DOCTYPE d [ <!ENTITY e "E"> ]><d>&e;</d>', 64);
    expect(list.find(e => e.kind === "error")!.value).toBe(
      "undefined entity: e (entities declared in a DOCTYPE internal subset are not supported)"
    );
  });
});

describe("SaxParser — a declared encoding and standalone", () => {
  it("reads a byte past 0x7F as no character of a document declared ASCII, however it is chunked", () => {
    const bytes = encoder.encode('<?xml version="1.0" encoding="US-ASCII"?>\r\n<r a="x">é\r</r>');
    for (const mode of ["error", "replace", "skip"] as const) {
      const outcomes = new Set<string>();
      for (let size = 1; size <= bytes.length; size++) {
        const parser = new SaxParser({ invalidCharHandling: mode });
        const seen: string[] = [];
        parser.on("error", err => seen.push(`error at ${(err as XmlParseError).byte}`));
        parser.on("text", text => seen.push(`text ${JSON.stringify(text)} ${parser.startByte}`));
        parser.on("closetag", tag => seen.push(`close ${tag.name} ${parser.startByte}`));
        for (let at = 0; at < bytes.length; at += size) {
          parser.write(bytes.subarray(at, at + size));
        }
        parser.close();
        outcomes.add(JSON.stringify(seen));
      }
      expect([...outcomes], mode).toHaveLength(1);
      const seen = JSON.parse([...outcomes][0]) as string[];
      if (mode === "error") {
        // Both bytes of é are refused where they are.
        expect(seen.filter(event => event.startsWith("error"))).toEqual([
          "error at 52",
          "error at 53"
        ]);
      } else {
        expect(seen).toContain(
          `text ${JSON.stringify(mode === "replace" ? "\uFFFD\uFFFD\n" : "\n")} 52`
        );
      }
      expect(seen[seen.length - 1]).toBe("close r 55");
    }
    // ASCII itself reads as it always did, and so does UTF-8 declared as such.
    expect(() =>
      new SaxParser()
        .write(encoder.encode('<?xml version="1.0" encoding="ascii"?><r>&#233;</r>'))
        .close()
    ).not.toThrow();
    expect(() =>
      new SaxParser()
        .write(encoder.encode('<?xml version="1.0" encoding="UTF-8"?><r>é</r>'))
        .close()
    ).not.toThrow();
  });

  it("holds a standalone document to declaring the parameter entities it refers to", () => {
    const standalone = '<?xml version="1.0" standalone="yes"?>';
    expect(() =>
      new SaxParser().write(`${standalone}<!DOCTYPE r SYSTEM "x" [%missing;]><r/>`).close()
    ).toThrow(/undeclared parameter entity/);
    expect(() =>
      new SaxParser().write(`${standalone}<!DOCTYPE r [<!ENTITY % p "x"> %p;]><r/>`).close()
    ).not.toThrow();
    // Not standalone, the external subset may declare it.
    expect(() =>
      new SaxParser().write('<!DOCTYPE r SYSTEM "x" [%missing;]><r/>').close()
    ).not.toThrow();
  });
});
